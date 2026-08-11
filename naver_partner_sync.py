"""
네이버쇼핑 파트너센터 — 가격비교 매칭 상품 배치 다운로드 + 로컬DB 동기화

배경: 네이버 쇼핑 검색 API 종료로 스마트스토어 자유검색을 대체할 방법이 없다.
대신 PM이 다른 서비스에서 검증한 방식(파트너센터 "가격비교 매칭" 상품을
Excel로 배치 다운로드 → 병합 → 로컬DB화)을 가져와, "몰별 비교"가 아니라
"네이버 참고 최저가" 숫자 하나를 보여주는 용도로 축소 적용한다.

로그인 방식: 비밀번호를 이 코드/서버에 저장하거나 흘려보내지 않는다.
1회는 사용자가 헤드풀 브라우저에서 직접 로그인(2단계인증 포함)하고,
그 로그인 세션(storage_state)만 로컬 파일로 저장해 재사용한다.
세션이 만료되면 `python naver_partner_sync.py login`을 다시 실행해 갱신한다.

주의: 이 파일의 파트너센터 화면 진입 경로(엑셀 다운 → 전체 엑셀 분할
다운로드 → 배치별 "다운로드"→"파일다운" 클릭)는 사용자가 준 스크린샷
기준으로 작성했다. 실제 배포/실행 환경에서 셀렉터가 안 맞으면 로그를
보고 조정이 필요하다 — 작성 시점에 실제 사이트를 열어 검증할 수 없었다
(자동화 도구의 도메인 접근이 차단되어 있음).
"""
import os
import re
import glob
import json
import time
import sqlite3
import threading

try:
    from playwright.sync_api import sync_playwright
    _HAS_PLAYWRIGHT = True
except ImportError:
    _HAS_PLAYWRIGHT = False

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_DATA_DIR = "/data" if os.path.isdir("/data") else BASE_DIR

STATE_PATH     = os.path.join(_DATA_DIR, "naver_partner_state.json")
DOWNLOAD_DIR   = os.path.join(_DATA_DIR, "naver_partner_downloads")
DB_PATH        = os.path.join(_DATA_DIR, "naver_ref_price.db")

PARTNER_ENTRY_URL = "https://center.shopping.naver.com/product/manage"

# 사용자가 확인해준 다운로드 Excel 컬럼 헤더 (그대로 매핑에 사용)
COL_PRODUCT_NAME   = "상품명"
COL_MALL_CATEGORY  = "쇼핑몰 카테고리"
COL_MALL_PRODUCT_ID = "쇼핑몰 상품ID"
COL_NAVER_ITEM_ID  = "네이버 가격비교 상품ID"
COL_BRAND          = "브랜드(네이버쇼핑)"
COL_SELL_PRICE     = "판매가"
COL_REG_DATE       = "등록일자"
COL_CATALOG_NAME   = "가격비교명"
COL_LOWEST_PRICE   = "가격비교 최저가"
COL_CATALOG_ID     = "가격비교 ID"

SYNC_INTERVAL_SEC = int(os.environ.get("NAVER_PARTNER_SYNC_INTERVAL_SEC", str(90 * 60)))
_BATCH_TIMEOUT_MS = 3 * 60 * 1000   # 배치당 최대 대기 (344개 연속 배치 실측 검증 완료, 3분으로 여유있게 설정)


def _init_db():
    os.makedirs(_DATA_DIR, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS naver_ref_price (
            mall_product_id TEXT PRIMARY KEY,
            product_name     TEXT,
            brand            TEXT,
            catalog_name     TEXT,
            catalog_id       TEXT,
            naver_item_id    TEXT,
            sell_price       INTEGER,
            lowest_price     INTEGER,
            reg_date         TEXT,
            updated_at       REAL
        )
    """)
    con.execute("CREATE INDEX IF NOT EXISTS idx_product_name ON naver_ref_price(product_name)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_brand ON naver_ref_price(brand)")
    con.commit()
    return con


# ── 1회 수동 로그인 → 세션 저장 ────────────────────────────────────────────────
def login_and_save_state():
    """헤드풀 브라우저를 띄워 사용자가 직접 로그인(2단계인증 포함)하도록 하고,
    로그인 완료 후 터미널에서 Enter를 누르면 storage_state를 저장한다.
    비밀번호는 이 스크립트를 거치지 않고 사용자가 브라우저에 직접 입력한다."""
    if not _HAS_PLAYWRIGHT:
        raise RuntimeError("playwright가 설치되어 있지 않습니다.")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.goto(PARTNER_ENTRY_URL, wait_until="domcontentloaded", timeout=60000)
        print(
            "\n브라우저 창에서 네이버 계정으로 로그인해주세요 (2단계 인증 포함).\n"
            "로그인이 완료되어 파트너센터 화면이 보이면, 이 터미널로 돌아와 Enter를 누르세요."
        )
        input()
        os.makedirs(_DATA_DIR, exist_ok=True)
        context.storage_state(path=STATE_PATH)
        print(f"로그인 세션 저장 완료: {STATE_PATH}")
        browser.close()


def _get_authed_context(pw, headless=True):
    if not os.path.exists(STATE_PATH):
        raise RuntimeError(
            f"로그인 세션 파일이 없습니다 ({STATE_PATH}). "
            "먼저 `python naver_partner_sync.py login`을 실행해 로그인하세요."
        )
    browser = pw.chromium.launch(headless=headless)
    context = browser.new_context(locale="ko-KR", storage_state=STATE_PATH,
                                   accept_downloads=True)
    return browser, context


# ── 배치 다운로드 ─────────────────────────────────────────────────────────────
def _row_has(row, label):
    return row.get_by_text(label, exact=True).count() > 0


def _click_row_label(row, label):
    """행 안에서 label 텍스트 요소를 찾아, 가장 가까운 클릭 가능한 조상
    (button/a/[onclick]/[role=button]/li/td)을 JS로 직접 클릭한다.
    Playwright의 가시성 액션어빌리티 체크를 완전히 건너뛰어, 접근성용
    숨김 텍스트(<span class="blind">) 패턴에도 안전하게 동작한다."""
    target = row.get_by_text(label, exact=True).first
    target.evaluate(
        "(el) => { "
        "  const c = el.closest('button, a, [onclick], [role=button], li, td'); "
        "  (c || el).click(); "
        "}"
    )


def _open_download_popup(page, inner):
    with page.context.expect_page(timeout=30000) as popup_info:
        inner.get_by_text("엑셀 다운", exact=False).first.click()
        inner.get_by_text("전체 엑셀 분할 다운로드", exact=False).first.click()
    popup = popup_info.value
    popup.wait_for_load_state("domcontentloaded", timeout=30000)
    return popup


def _dismiss_modal_if_present(popup):
    """'이미 진행중인 다운로드가 있습니다' 같은 커스텀 확인 모달이 떠 있으면 닫는다."""
    try:
        ok_btn = popup.get_by_text("확인", exact=True)
        if ok_btn.count() > 0:
            ok_btn.first.click()
            popup.wait_for_timeout(500)
            return True
    except Exception:
        pass
    return False


_BATCH_COOLDOWN_SEC = 3   # 배치 사이 서버 쿨다운 (5초로 3배치 연속 성공 확인됨, 여유 두고 3~5초 사용)
_MODAL_BACKOFF_SEC = [10, 30, 60]   # '이미 진행중' 모달을 만났을 때 재시도 전 대기(점증)

def _download_all_batches(page, inner, out_dir):
    """'전체 엑셀 분할 다운로드' 팝업 하나를 계속 열어둔 채 배치를 순서대로 처리한다.
    배치 사이에 서버 쿨다운(_BATCH_COOLDOWN_SEC)이 필요함이 실측으로 확인됨 —
    쿨다운 없이 바로 다음 배치를 요청하면 '이미 진행중인 다운로드가 있습니다'
    모달이 뜨며 막힌다. 그래도 모달이 뜨면 점증 백오프로 재시도한다.
    inner: 상품 관리 화면이 들어있는 iframe의 frame_locator (엑셀 다운 버튼도 이 안에 있음).
    저장된 파일 경로 리스트 반환."""
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    seen_ranges = set()
    stalled_rounds = 0
    is_first_batch = True

    popup = _open_download_popup(page, inner)
    _dismiss_modal_if_present(popup)

    while stalled_rounds < 3:
        rows = popup.locator("table tr")
        row_count = rows.count()
        target = None
        for i in range(row_count):
            row = rows.nth(i)
            text = row.inner_text()
            m = re.search(r"(\d+)\s*[~\-]\s*(\d+)", text)
            if not m:
                continue
            range_key = m.group(0)
            if range_key in seen_ranges:
                continue
            target = (row, range_key)
            break

        if target is None:
            popup.wait_for_timeout(3000)
            stalled_rounds += 1
            continue

        row, range_key = target
        progressed = False
        try:
            if not is_first_batch:
                popup.wait_for_timeout(_BATCH_COOLDOWN_SEC * 1000)
            is_first_batch = False

            if _row_has(row, "다운로드") and not _row_has(row, "파일다운"):
                _click_row_label(row, "다운로드")
                deadline = time.time() + _BATCH_TIMEOUT_MS / 1000
                backoff_used = 0
                while not _row_has(row, "파일다운") and time.time() < deadline:
                    if _dismiss_modal_if_present(popup):
                        wait_sec = _MODAL_BACKOFF_SEC[min(backoff_used, len(_MODAL_BACKOFF_SEC) - 1)]
                        backoff_used += 1
                        popup.wait_for_timeout(wait_sec * 1000)
                        _click_row_label(row, "다운로드")
                    popup.wait_for_timeout(2000)
                if not _row_has(row, "파일다운"):
                    raise TimeoutError("파일다운 대기 타임아웃")

            with popup.expect_download(timeout=_BATCH_TIMEOUT_MS) as dl_info:
                _click_row_label(row, "파일다운")
            download = dl_info.value

            fname = f"batch_{range_key.replace('~', '_').replace('-', '_')}.xlsx"
            fpath = os.path.join(out_dir, fname)
            download.save_as(fpath)
            saved.append(fpath)
            seen_ranges.add(range_key)
            progressed = True
        except Exception as e:
            print(f"[naver-partner-sync] 배치 {range_key} 처리 실패: {e}")
            try:
                os.makedirs(_DATA_DIR, exist_ok=True)
                popup.screenshot(
                    path=os.path.join(_DATA_DIR, f"naver_partner_stall_{range_key.replace('~', '_')}.png"),
                    full_page=True,
                )
            except Exception:
                pass
            seen_ranges.add(range_key)

        stalled_rounds = 0 if progressed else stalled_rounds + 1

    popup.close()
    return saved


# ── 엑셀 → DB 병합 ────────────────────────────────────────────────────────────
def _upsert_excel_file(con, filepath):
    # read_only 모드는 이 사이트가 내려주는 xlsx에서 내용을 제대로 못 읽는
    # 경우가 확인되어 일반 모드로 로드한다 (배치당 1000행 수준이라 메모리 문제 없음).
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    required = [COL_PRODUCT_NAME, COL_MALL_PRODUCT_ID, COL_LOWEST_PRICE]
    header, idx = None, {}
    # 1행은 보통 "서비스 상품_partN" 같은 제목 행이라, 앞쪽 몇 줄을 살펴
    # 필수 컬럼이 다 있는 실제 헤더 행을 찾는다.
    for _ in range(5):
        try:
            candidate = next(rows_iter)
        except StopIteration:
            break
        cand_header = [str(h).strip() if h else "" for h in candidate]
        cand_idx = {name: cand_header.index(name) for name in cand_header if name}
        if all(c in cand_idx for c in required):
            header, idx = cand_header, cand_idx
            break

    if header is None:
        raise ValueError(f"{filepath}: 필수 컬럼을 가진 헤더 행을 찾지 못함")

    def _num(v):
        if v is None:
            return None
        s = re.sub(r"[^\d]", "", str(v))
        return int(s) if s else None

    now = time.time()
    n = 0
    for row in rows_iter:
        mall_pid = row[idx[COL_MALL_PRODUCT_ID]] if idx.get(COL_MALL_PRODUCT_ID) is not None else None
        if not mall_pid:
            continue
        con.execute(
            """INSERT INTO naver_ref_price
               (mall_product_id, product_name, brand, catalog_name, catalog_id,
                naver_item_id, sell_price, lowest_price, reg_date, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(mall_product_id) DO UPDATE SET
                 product_name=excluded.product_name, brand=excluded.brand,
                 catalog_name=excluded.catalog_name, catalog_id=excluded.catalog_id,
                 naver_item_id=excluded.naver_item_id, sell_price=excluded.sell_price,
                 lowest_price=excluded.lowest_price, reg_date=excluded.reg_date,
                 updated_at=excluded.updated_at""",
            (
                str(mall_pid),
                row[idx[COL_PRODUCT_NAME]] if idx.get(COL_PRODUCT_NAME) is not None else None,
                row[idx[COL_BRAND]] if idx.get(COL_BRAND) is not None else None,
                row[idx[COL_CATALOG_NAME]] if idx.get(COL_CATALOG_NAME) is not None else None,
                str(row[idx[COL_CATALOG_ID]]) if idx.get(COL_CATALOG_ID) is not None and row[idx[COL_CATALOG_ID]] else None,
                str(row[idx[COL_NAVER_ITEM_ID]]) if idx.get(COL_NAVER_ITEM_ID) is not None and row[idx[COL_NAVER_ITEM_ID]] else None,
                _num(row[idx[COL_SELL_PRICE]]) if idx.get(COL_SELL_PRICE) is not None else None,
                _num(row[idx[COL_LOWEST_PRICE]]) if idx.get(COL_LOWEST_PRICE) is not None else None,
                str(row[idx[COL_REG_DATE]]) if idx.get(COL_REG_DATE) is not None and row[idx[COL_REG_DATE]] else None,
                now,
            ),
        )
        n += 1
    con.commit()
    wb.close()
    return n


# ── 전체 동기화 1회 실행 ───────────────────────────────────────────────────────
def run_sync():
    """로그인 세션 재사용 → 배치 다운로드 → DB 병합. 예외 발생 시 그대로 올림
    (호출부에서 로깅/재시도 정책 결정)."""
    if not _HAS_PLAYWRIGHT:
        raise RuntimeError("playwright가 설치되어 있지 않습니다.")
    with sync_playwright() as pw:
        browser, context = _get_authed_context(pw, headless=True)
        try:
            page = context.new_page()
            page.goto(PARTNER_ENTRY_URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(1500)
            try:
                # 실제 상품 관리 화면은 iframe 안에 렌더링됨 (embrace-token-at-url 경유)
                page.wait_for_selector("iframe", timeout=15000)
                inner = page.frame_locator("iframe")
                inner.get_by_role("link", name="서비스 상품", exact=True).click(timeout=15000)
                batch_dir = os.path.join(DOWNLOAD_DIR, str(int(time.time())))
                files = _download_all_batches(page, inner, batch_dir)
            except Exception:
                os.makedirs(_DATA_DIR, exist_ok=True)
                shot_path = os.path.join(_DATA_DIR, "naver_partner_debug.png")
                html_path = os.path.join(_DATA_DIR, "naver_partner_debug.html")
                try:
                    page.screenshot(path=shot_path, full_page=True)
                    with open(html_path, "w", encoding="utf-8") as f:
                        f.write(page.content())
                    print(f"[naver-partner-sync] 실패 — 디버그용 스크린샷/HTML 저장: {shot_path}, {html_path}")
                except Exception as _e2:
                    print(f"[naver-partner-sync] 디버그 캡처도 실패: {_e2}")
                raise
            print(f"[naver-partner-sync] {len(files)}개 배치 파일 다운로드 완료")
        finally:
            browser.close()

    con = _init_db()
    total = 0
    for f in files:
        try:
            total += _upsert_excel_file(con, f)
        except Exception as e:
            print(f"[naver-partner-sync] {f} 병합 실패: {e}")
    con.close()
    print(f"[naver-partner-sync] DB 병합 완료: {total}건 upsert")

    for f in files:
        try:
            os.remove(f)
        except OSError:
            pass
    return total


# ── 조회 (server.py에서 사용) ──────────────────────────────────────────────────
def lookup_naver_reference_price(keyword):
    """검색어(공백으로 여러 토큰 가능)와 상품명/브랜드가 모두 매칭되는 행 중
    최저가(lowest_price)가 가장 낮은 값을 반환. 매칭 없으면 None."""
    if not keyword or not os.path.exists(DB_PATH):
        return None
    tokens = [t for t in keyword.strip().split() if t]
    if not tokens:
        return None
    con = sqlite3.connect(DB_PATH)
    try:
        clauses, params = [], []
        for t in tokens:
            clauses.append("(product_name LIKE ? OR brand LIKE ?)")
            params.extend([f"%{t}%", f"%{t}%"])
        sql = (
            "SELECT MIN(lowest_price), COUNT(*) FROM naver_ref_price "
            f"WHERE lowest_price IS NOT NULL AND lowest_price > 0 AND {' AND '.join(clauses)}"
        )
        row = con.execute(sql, params).fetchone()
        if not row or row[1] == 0:
            return None
        return {"price": row[0], "match_count": row[1]}
    finally:
        con.close()


# ── 백그라운드 주기 실행 ───────────────────────────────────────────────────────
_sync_thread = None

def start_background_sync_thread():
    """서버 프로세스 안에서 주기적으로 run_sync()를 실행하는 데몬 스레드 시작.
    로그인 세션 파일이 없으면 조용히 건너뜀(최초 1회 `login` 실행 전 상태)."""
    global _sync_thread
    if _sync_thread is not None:
        return
    if not os.path.exists(STATE_PATH):
        print(f"[naver-partner-sync] 로그인 세션 없음({STATE_PATH}) — 백그라운드 동기화 건너뜀")
        return

    def _loop():
        while True:
            try:
                run_sync()
            except Exception as e:
                print(f"[naver-partner-sync] 동기화 실패: {e}")
            time.sleep(SYNC_INTERVAL_SEC)

    _sync_thread = threading.Thread(target=_loop, daemon=True)
    _sync_thread.start()
    print(f"[naver-partner-sync] 백그라운드 동기화 시작 (주기: {SYNC_INTERVAL_SEC}초)")


if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "login":
        login_and_save_state()
    elif cmd == "sync":
        run_sync()
    elif cmd == "serve":
        start_background_sync_thread()
        while True:
            time.sleep(3600)
    else:
        print("사용법: python naver_partner_sync.py [login|sync|serve]")

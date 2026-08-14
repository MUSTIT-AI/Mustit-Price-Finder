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
import sys
import re
import glob
import json
import time
import sqlite3
import threading

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import requests

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
SERVICE_LIST_URL  = "https://adcenter.shopping.naver.com/iframe/product/manage/service/list.nhn"
IFRAME_URL_HINT   = "adcenter.shopping.naver.com/iframe/product/manage"

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


# ── 배치 다운로드 (PM이 검증한 다른 서비스 PoC 방식 그대로 채용 — 1000건당 3.4초 실측) ──
def _wait_for_content_frame(page, timeout_ms=15000):
    """상품관리 iframe(실제 URL이 IFRAME_URL_HINT를 포함하는 프레임)을 찾는다.
    page.frame_locator('iframe')로 접근하는 것보다, 실제 Frame 객체를 직접
    찾는 이 방식이 frame.goto() 이후에도 안정적으로 재사용 가능하다."""
    elapsed, step = 0, 300
    while elapsed < timeout_ms:
        for frame in page.frames:
            if IFRAME_URL_HINT in frame.url:
                return frame
        page.wait_for_timeout(step)
        elapsed += step
    return None


def _get_last_row_label(popup):
    els = popup.locator("#downloadList tr .num")
    n = els.count()
    if n == 0:
        return None
    return els.nth(n - 1).inner_text().strip()


def _open_download_popup(page):
    """상품관리 iframe에서 '가격비교매칭완료' 필터를 적용한 뒤 전체 엑셀 분할
    다운로드 팝업을 연다. 이 필터를 먼저 적용하는 것이 PM이 확인한 다른
    서비스 PoC와의 핵심 차이 — 필터 없이 하면 배치당 처리시간이 10배 이상
    느려짐(전체 336만건 중 미매칭 상품까지 서버가 훑는 것으로 추정)."""
    content_frame = _wait_for_content_frame(page)
    if content_frame is None:
        raise RuntimeError("상품관리 iframe을 찾지 못함")

    content_frame.goto(SERVICE_LIST_URL)
    content_frame.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    # goto 이후 프레임 참조가 끊길 수 있어 재확인
    content_frame = _wait_for_content_frame(page)
    if content_frame is None:
        raise RuntimeError("상품관리 iframe(goto 이후)을 찾지 못함")

    try:
        content_frame.locator('a[status="MODEL_MATCHED"]').first.click()
        page.wait_for_timeout(1500)
    except Exception as e:
        print(f"[naver-partner-sync] 가격비교매칭완료 필터 클릭 실패(무시하고 진행): {e}")

    with page.context.expect_page(timeout=30000) as popup_info:
        content_frame.locator("#excelDown a.tab_toggle").click()
        page.wait_for_timeout(500)
        content_frame.locator("#excelDown li[key='whole'] a").click()
    popup = popup_info.value
    try:
        popup.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    popup.wait_for_timeout(2000)
    return popup


def _download_all_batches(page, out_dir):
    """전체 엑셀 분할 다운로드 팝업 하나를 계속 열어둔 채 배치를 순서대로 처리.
    각 배치는 '#downloadList a.btn_down'의 마지막 행을 클릭 → 새 행(다음 배치)이
    나타날 때까지 대기, 이 한 번의 클릭으로 처리+다운로드가 함께 일어남
    (다운로드는 context 레벨 'download' 이벤트로 수집). '이미 진행중인 다운로드'
    다이얼로그가 뜨면 점증 대기(3×시도횟수초) 후 같은 행을 재클릭.
    저장된 파일 경로 리스트 반환."""
    os.makedirs(out_dir, exist_ok=True)
    downloaded = []
    page.context.on("download", lambda d: downloaded.append(d))

    last_dialog_message = {"text": None}

    def _on_dialog(dialog):
        last_dialog_message["text"] = dialog.message
        dialog.accept()
    page.on("dialog", _on_dialog)

    def _dismiss_html_modal(p):
        """'이미 진행중인 다운로드가 있습니다' 등은 네이티브 브라우저 다이얼로그가
        아니라 커스텀 HTML 모달(.swal-overlay)로 뜨는 경우가 있어, page.on('dialog')
        로는 못 잡는다. 확인 버튼을 직접 클릭해서 닫는다."""
        try:
            overlay = p.locator(".swal-overlay--show-modal")
            if overlay.count() > 0:
                p.locator(".swal-button--confirm").first.click(timeout=3000)
                p.wait_for_timeout(300)
                return True
        except Exception:
            pass
        return False

    popup = _open_download_popup(page)
    popup.on("dialog", _on_dialog)

    MAX_RETRY = 5
    i = 0
    while True:
        if popup.locator("#downloadList a.btn_down").count() == 0:
            print(f"[naver-partner-sync] 더 이상 받을 배치가 없음 - 총 {i}개로 종료")
            break

        i += 1
        label_before = _get_last_row_label(popup)
        new_row_ok = False
        for attempt in range(1, MAX_RETRY + 1):
            last_dialog_message["text"] = None
            _dismiss_html_modal(popup)
            try:
                popup.locator("#downloadList a.btn_down").last.click()
            except Exception as e:
                print(f"[naver-partner-sync] 배치 {i} 다운로드 버튼 클릭 실패: {e}")
                break
            try:
                popup.wait_for_function(
                    "prevLabel => { "
                    "  const els = document.querySelectorAll('#downloadList tr .num'); "
                    "  if (!els.length) return false; "
                    "  const last = els[els.length - 1].textContent.trim(); "
                    "  return last !== prevLabel; }",
                    arg=label_before,
                    timeout=15000,
                )
                new_row_ok = True
                break
            except Exception:
                modal_seen = _dismiss_html_modal(popup)
                is_conflict = modal_seen or (
                    last_dialog_message["text"] and "이미 진행중인 다운로드" in last_dialog_message["text"]
                )
                if is_conflict:
                    wait_s = 3 * attempt
                    popup.wait_for_timeout(wait_s * 1000)
                    continue
                popup.wait_for_timeout(2000)

        if not new_row_ok:
            print(f"[naver-partner-sync] 배치 {i}에서 정지(재시도 {MAX_RETRY}회 소진) - 종료")
            try:
                os.makedirs(_DATA_DIR, exist_ok=True)
                popup.screenshot(path=os.path.join(_DATA_DIR, "naver_partner_stall.png"), full_page=True)
                with open(os.path.join(_DATA_DIR, "naver_partner_stall.html"), "w", encoding="utf-8") as f:
                    f.write(popup.content())
            except Exception:
                pass
            break

    saved = []
    for d in downloaded:
        fpath = os.path.join(out_dir, d.suggested_filename)
        try:
            d.save_as(fpath)
            saved.append(fpath)
        except Exception as e:
            print(f"[naver-partner-sync] 다운로드 저장 실패: {e}")

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
            page.goto(PARTNER_ENTRY_URL, wait_until="networkidle", timeout=60000)
            page.wait_for_timeout(1500)
            try:
                batch_dir = os.path.join(DOWNLOAD_DIR, str(int(time.time())))
                files = _download_all_batches(page, batch_dir)
            except Exception:
                os.makedirs(_DATA_DIR, exist_ok=True)
                shot_path = os.path.join(_DATA_DIR, "naver_partner_debug.png")
                html_path = os.path.join(_DATA_DIR, "naver_partner_debug.html")
                try:
                    page.screenshot(path=shot_path, full_page=True)
                    with open(html_path, "w", encoding="utf-8") as f:
                        f.write(page.content())
                    print(f"[naver-partner-sync] 실패 - 디버그용 스크린샷/HTML 저장: {shot_path}, {html_path}")
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


def lookup_naver_reference_price_by_mustit_ids(item_nos):
    """머스트잇 검색 결과의 itemNo(=엑셀 '쇼핑몰 상품ID')로 정확히 매칭해 조회.
    텍스트 매칭(lookup_naver_reference_price)보다 정확함 — 상품코드 검색처럼
    상품명 텍스트로는 못 찾는 경우에도 머스트잇 자체 검색이 이미 정확한 상품을
    찾아준 경우라면 그 itemNo로 바로 매칭 가능. 매칭 없으면 None."""
    item_nos = [str(x) for x in (item_nos or []) if x]
    if not item_nos or not os.path.exists(DB_PATH):
        return None
    con = sqlite3.connect(DB_PATH)
    try:
        placeholders = ",".join("?" * len(item_nos))
        sql = (
            "SELECT MIN(lowest_price), COUNT(*) FROM naver_ref_price "
            f"WHERE lowest_price IS NOT NULL AND lowest_price > 0 "
            f"AND mall_product_id IN ({placeholders})"
        )
        row = con.execute(sql, item_nos).fetchone()
        if not row or row[1] == 0:
            return None
        return {"price": row[0], "match_count": row[1]}
    finally:
        con.close()


# ── 결과 DB를 웹 서비스로 전송 (별도 Railway 서비스로 분리 배포된 경우) ────────────
def push_db_to_web_service():
    """동기화 완료 후, 이 프로세스와 별도로 배포된 웹 서비스에 내부망(private
    networking)으로 최신 DB 파일을 전송한다. NAVER_PARTNER_UPLOAD_URL 환경변수가
    없으면(로컬 개발 등) 조용히 건너뜀. ACCESS_PASSWORD가 설정돼 있으면 웹 서비스의
    Basic Auth를 통과하도록 같은 값을 헤더에 실어 보낸다."""
    url = os.environ.get("NAVER_PARTNER_UPLOAD_URL", "").strip()
    if not url or not os.path.exists(DB_PATH):
        return
    pw = os.environ.get("ACCESS_PASSWORD", "").strip()
    auth = ("sync", pw) if pw else None
    try:
        with open(DB_PATH, "rb") as f:
            r = requests.post(url, files={"file": ("naver_ref_price.db", f)}, auth=auth, timeout=180)
        if r.status_code == 200:
            print(f"[naver-partner-sync] DB 전송 완료 -> {url}")
        else:
            print(f"[naver-partner-sync] DB 전송 실패({r.status_code}): {r.text[:200]}")
    except Exception as e:
        print(f"[naver-partner-sync] DB 전송 오류: {e}")


# ── 백그라운드 주기 실행 ───────────────────────────────────────────────────────
_sync_thread = None

def start_background_sync_thread():
    """서버 프로세스 안에서 주기적으로 run_sync()를 실행하는 데몬 스레드 시작.
    로그인 세션 파일이 없으면 조용히 건너뜀(최초 1회 `login` 실행 전 상태)."""
    global _sync_thread
    if _sync_thread is not None:
        return
    if not os.path.exists(STATE_PATH):
        print(f"[naver-partner-sync] 로그인 세션 없음({STATE_PATH}) - 백그라운드 동기화 건너뜀")
        return

    def _loop():
        while True:
            try:
                run_sync()
                push_db_to_web_service()
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

# Mustit Price Finder — 개발자 온보딩 가이드

머스트잇·트렌비·SSG·롯데온·스마트스토어 **5개 명품몰 가격비교** 웹앱.
Flask 백엔드(`server.py`) + 단일 페이지 프론트(`index.html`) 구조이며 **Railway** 에 배포됩니다.

---

## ⚠️ 먼저 읽어주세요 — 현재 상태

- **네이버 쇼핑 검색 API가 종료**되어, 기존의 상품 검색·수집 로직(`call_api` 등)이 **동작하지 않습니다.**
- 직접 크롤링도 시도했으나 네이버가 자동화 트래픽을 강하게 차단(418 / 로그인벽 / IP 일시제한)합니다. **서버(데이터센터 IP)에서는 더 심하게 막힙니다.**
- 따라서 현재 핵심 과제는 **데이터 소스 교체**입니다 (아래 둘 중 방향 협의 필요):
  - **C안**: 유료/공식 API(네이버 커머스API 또는 유료 가격비교/SERP)
  - **D안**: 5개몰 각 사이트 검색을 직접 연동 (상세 스크래퍼는 이미 `server.py`에 존재)
- 진단용 스크립트: `naver_probe.py` (네이버 접근 가능성 테스트, 참고용)

> 요약: **UI·판매자매칭·상세스크래퍼는 살아있고, "상품 목록을 어디서 가져올지"만 새로 붙이면 됩니다.**

---

## 1. 로컬 실행

### 사전 준비
- **Python 3.11+**
- (선택) 상세 스크래핑용 **Playwright Chromium**

### 설치 & 실행
```bash
# 1) 의존성 설치
pip install -r requirements.txt

# 2) Playwright 브라우저 설치 (머스트잇 등 상세 스크래핑에 필요)
python -m playwright install chromium

# 3) 실행 (개발 모드)
python server.py
```
→ 브라우저에서 **http://127.0.0.1:5050** 접속.

### 환경변수 / 설정
| 변수 | 용도 | 비고 |
|---|---|---|
| `NAVER_CLIENT_ID` / `NAVER_CLIENT_SECRET` | (구)네이버 검색 API 키 | **API 종료로 현재 무의미** |
| `ACCESS_PASSWORD` | 접속 비밀번호(Basic Auth) | 미설정 시 인증 없음(로컬 개발용) |
| `PORT` | 서버 포트 | 미설정 시 5050(로컬) |

로컬에서는 환경변수 대신 `api_keys.json` 파일로도 키를 넣을 수 있습니다(이 파일은 `.gitignore`로 커밋 제외됨).

---

## 2. 코드 구조
| 파일 | 역할 |
|---|---|
| `server.py` | Flask 백엔드 전부 — 검색·플랫폼분류·판매자 스크래핑·매칭·CSV·디버그 라우트 |
| `index.html` | 프론트 전체 (HTML+CSS+JS 인라인) |
| `seller_config.json` | 판매자 × 5개몰 셀러ID 매핑표 |
| `Dockerfile` / `Procfile` | Railway 배포 설정 |
| `requirements.txt` | Python 의존성 |

주요 함수(참고): `search_by_platform`(핵심 흐름), `enrich_sellers_in_place`(판매자 병렬 스크래핑), `_fetch_trenbe_detail`/`_fetch_lotteon_detail`/`_fetch_ssg_detail`/`_fetch_mustit_detail`(몰별 상세).

---

## 3. Git & 배포 워크플로우

```
코드 수정(로컬) → commit → git push origin main → Railway 자동 빌드(2~5분) → 라이브 반영
```

- **기본 브랜치: `main`.** 여기에 push하면 **자동 배포**됩니다(별도 PR/리뷰 없음).
- **보호 설정:** `main`의 **force-push·브랜치 삭제는 차단**됩니다(일반 push는 자유). 실수로 이력을 날릴 수 없습니다.
- **롤백:** 배포가 잘못되면 Railway 대시보드에서 **이전 배포로 1클릭 롤백** 가능.
- **커밋 금지 대상:** `api_keys.json`, `*.db`, `pw_profile/`, 크롤링 산출물 → 이미 `.gitignore` 처리됨.

```bash
git clone https://github.com/MUSTIT-AI/Mustit-Price-Finder.git
cd Mustit-Price-Finder
# ... 수정 ...
git add -A && git commit -m "설명" && git push origin main
```

---

## 4. Claude Code + GitHub MCP (바이브 코딩 세팅)

[Claude Code](https://claude.com/claude-code)로 AI 페어 개발 시, GitHub MCP를 붙이면 이슈·PR·레포를 Claude가 직접 다룰 수 있습니다.

### 준비물 — GitHub 토큰(PAT)
https://github.com/settings/personal-access-tokens 에서 **Fine-grained token** 발급, 이 레포에 대해:
- `Contents` : Read and write
- `Pull requests` : Read and write
- `Issues` : Read and write
- (선택) `Actions` : Read

### 방법 A — 원격(호스티드) MCP · 가장 간단
```bash
claude mcp add --transport http github https://api.githubcopilot.com/mcp/ \
  --header "Authorization: Bearer YOUR_GITHUB_PAT"
```

### 방법 B — 로컬 Docker MCP (Docker 필요)
```bash
claude mcp add --transport stdio github-local \
  -e GITHUB_PERSONAL_ACCESS_TOKEN=YOUR_GITHUB_PAT \
  -- docker run --rm -i ghcr.io/github/github-mcp-server:latest
```

### 팀 공유용 `.mcp.json` (선택)
프로젝트 루트에 아래 파일을 두면 팀원이 각자 토큰만 넣고 바로 사용합니다.
각자 실행 전 `export GITHUB_PAT=<본인토큰>` (Windows PowerShell: `$env:GITHUB_PAT="<본인토큰>"`).
```json
{
  "mcpServers": {
    "github": {
      "type": "http",
      "url": "https://api.githubcopilot.com/mcp/",
      "headers": { "Authorization": "Bearer ${GITHUB_PAT}" }
    }
  }
}
```
> `.mcp.json`에는 **토큰 값을 직접 쓰지 말 것**(환경변수 참조만). 토큰이 박힌 파일은 커밋 금지.

### 연결 확인
```bash
claude mcp list          # ✔ Connected 확인
```
세션 안에서 `/mcp` 입력 → github 서버 선택 → 도구 목록 확인. 테스트: "내 오픈 PR 보여줘".

---

## 5. 주의사항 요약
- 🔴 **네이버 직접 크롤링 금지** — IP 차단(사무실 IP까지 막힘) 유발. 데이터 소스는 C/D안으로 협의.
- 🔒 **시크릿·데이터 커밋 금지** — `.gitignore` 확인. 실 키/판매자DB는 Railway 환경변수·볼륨에만.
- 🚀 **`main` push = 즉시 프로덕션 배포** — 큰 변경은 배포 후 라이브(URL) 확인 습관.

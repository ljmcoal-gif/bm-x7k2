# CLAUDE.md

이 저장소에서 작업하는 AI 어시스턴트(Claude Code 및 기타)를 위한 가이드입니다.

## 이 프로젝트는

배터리·딥테크 산업 뉴스를 실시간으로 모니터링하는 **단일 파일 클라이언트
사이드 웹 대시보드**입니다. UI와 모든 문구는 **한국어**(`lang="ko"`)입니다.

- 🔋 **배터리·딥테크 모니터링**
- 카테고리별 뉴스 추적: 수주/계약, 생산 차질, 운송 차질, M&A·딜,
  기술 동향, 방산·딥테크.
- 한국어(🇰🇷)·영어(🇺🇸) 검색어 모두에 대해 **Google News RSS**를 수집하고,
  각 기사를 분류한 뒤 차트, 요약 카드, 기업 언급 빈도 목록, 소스 탭,
  필터 가능한 뉴스 피드를 렌더링합니다.

## 저장소 구성

이 저장소는 **단 하나의 파일**로 이루어져 있습니다:

```
index.html   ← 애플리케이션 전체 (HTML + CSS + JS, 모두 인라인)
```

**빌드 단계, 패키지 매니저, 설치할 의존성, 테스트 스위트가 전혀 없습니다.**
유일한 런타임 외부 의존성은 CDN으로 로드되는 Chart.js
(`cdn.jsdelivr.net/npm/chart.js@4.4.0`)뿐입니다.

> **파일 헤더에 대한 참고:** `index.html` 상단 주석은 이 파일을 "독립형
> (standalone)" 버전으로 설명하며, "공식" Flask 백엔드(`app.py`, 5001 포트,
> Naver·TheBell 지원)를 언급합니다. **그러나 그 `app.py`는 이 저장소에
> 존재하지 않습니다** — 독립형 `index.html`만 있습니다. 백엔드가 있다고
> 가정하지 마세요. 작업에 백엔드가 필요하면 먼저 사용자에게 확인하세요.

## 실행 방법

모든 것이 클라이언트 사이드이므로 파일을 그냥 열면 됩니다:

- 브라우저에서 `index.html`을 직접 열거나(`file://`로도 동작), **또는**
- 로컬에서 서빙. 예: `python3 -m http.server 8000` 실행 후
  `http://localhost:8000/index.html` 접속.

로드 즉시 뉴스를 가져오고, 이후 **10분**마다 자동 새로고침합니다
(`AUTO_REFRESH = 10 * 60 * 1000`).

## 아키텍처 (모두 `index.html` 내부)

파일은 세 부분으로 구성됩니다: `<style>` 블록, 정적 HTML 마크업, 그리고 두
개의 `<script>` 블록(설정, 그다음 엔진).

### 1. 설정 스크립트 (첫 번째 `<script>`)

순수 JS 배열/객체 — **콘텐츠/동작을 수정할 주요 위치입니다**:

- `BATTERY_COMPANIES` — 기사 텍스트와 매칭되는 기업명(한국어 + 영어 별칭,
  예: `"LG에너지솔루션"`, `"LGES"`).
- `CATEGORY_KEYWORDS` — 카테고리별 키워드 목록. 기사 텍스트에 키워드가 하나라도
  나타나면 해당 카테고리가 부여됩니다.
- `SEVERITY_KEYWORDS` — `critical` / `warning` / `info` 키워드 목록. **순서가
  중요** — 가장 먼저 매칭되는 단계가 적용됩니다(critical을 먼저 검사).

### 2. 엔진 스크립트 (두 번째 `<script>`)

- `CORS_PROXIES` — 공개 CORS 프록시 빌더의 순서 목록(codetabs, allorigins,
  corsproxy.io). `fetchWithProxy()`가 순서대로 시도하여 RSS처럼 보이는 텍스트를
  반환할 때까지 진행. 개방형 프록시를 앞에 둬서 `file://`에서도 작동합니다.
- `SEARCH_QUERIES`(한국어)와 `SEARCH_QUERIES_EN`(영어) — Google News 검색어.
  여기에서 쿼리를 추가/삭제하여 수집 범위를 조정합니다.
- `fetchGoogleNewsRSS(query, locale)`는 알맞은 로케일 파라미터(`hl/gl/ceid`)로
  Google News RSS URL을 만들고 결과를 파싱합니다.
- `parseRSS()`는 `DOMParser`로 RSS XML을 파싱하고, `hashStr()`는 중복 제거를
  위해 `title + source`로 안정적인 id를 생성합니다.
- `fetchAllNews()`는 모든 쿼리를 병렬 실행하고, id로 중복을 제거하며, 분류한 뒤
  결과가 도착하는 대로 점진적으로 렌더링합니다.
- 분류: `classifyCategory()`, `classifySeverity()`, `findCompanies()`.
- 필터링: `applyFilters()`와 `filterBy*` 헬퍼들(카테고리, 심각도, 기업, 소스)이
  드롭다운, 검색 박스, 요약 카드, 소스 탭으로 구동됩니다. 이벤트 처리는
  **이벤트 위임**(DOMContentLoaded 리스너에서 설정)을 사용 — 동적으로
  렌더링되는 요소는 인라인 핸들러 대신 `data-*` 속성을 가집니다.
- 렌더링: `renderDashboard()`가 `renderSummaryCards`, `renderCharts`(Chart.js
  도넛 + 바), `renderCompanyList`, `renderSourceTabs`, `renderAlertBanner`,
  `renderNewsFeed`를 조율합니다.
- 번역: 영어 기사에는 🌐 번역 버튼이 붙으며, 무료 **MyMemory** API
  (`translateArticle` / `_mymemory`, en→ko, 키 불필요)를 호출합니다. 메모리 내
  `_translationCache`와 원문/번역 토글을 사용합니다.
- 유틸리티: `esc()`(HTML 이스케이프 — **`innerHTML`에 텍스트를 주입할 때 항상
  사용**), `truncate()`, `fmtDate()`(상대 시간 한국어 표기), `showToast()`,
  `showLoading()`.

## 따라야 할 규칙

- **단일 파일 유지.** 사용자가 명시적으로 분리를 요청하지 않는 한 모든 HTML,
  CSS, JS는 `index.html`에 인라인으로 둡니다.
- **프레임워크/번들러 금지.** 바닐라 JS만 사용하며, 유일한 CDN 의존성은
  Chart.js입니다. 요청 없이 npm/빌드 도구를 도입하지 마세요.
- **한국어 우선 UI.** 사용자 대상 문자열은 한국어입니다. 기존 톤과 이모지 접두
  레이블(📝 수주, 🏭 생산 차질 등)을 맞추세요.
- **보안:** 사용자/피드에서 유래한 텍스트를 `innerHTML`로 삽입할 때는 반드시
  `esc()`를 거쳐야 합니다(셀렉터에는 `translateArticle`처럼 `CSS.escape()` 사용).
- **스타일링:** `:root`의 기존 CSS 커스텀 속성(예: `--bg-card`, `--accent-blue`)과
  확립된 다크 테마 팔레트를 사용하세요. 레이아웃은 1200/768/480px 분기점으로
  반응형입니다.
- **소스 주석**은 한국어로 작성하며 박스 드로잉 구분선(`═══`, `──`)을 사용합니다.
  주변에 주석을 추가할 때 이 스타일을 유지하세요.
- **카테고리는 고정 집합.** 카테고리를 추가하면 다음을 *모두* 갱신하세요:
  `CATEGORY_KEYWORDS`, 요약 카드 마크업, 필터 `<select>`, `renderCharts`의 차트
  레이블/색상, `computeStats` 통계 객체, `catBadges` 맵, `updateActiveCards`.

## Git 워크플로

- AI 작업용 활성 개발 브랜치: **`claude/claude-md-docs-4032wb`**
  (명시적 허가 없이 `main`에 푸시하지 마세요).
- `git push -u origin <branch-name>`으로 푸시합니다.
- 사용자가 명시적으로 요청하지 않는 한 풀 리퀘스트를 열지 **마세요**.

## 주의 사항 (Gotchas)

- 이 앱은 **공개 CORS 프록시**와 **Google News RSS**의 접근 가능성에
  의존합니다. 뉴스가 로드되지 않으면 대개 코드 버그가 아니라 프록시/네트워크
  문제입니다(📡 메시지와 함께 빈 상태가 표시됩니다).
- 쿼리에는 `encodeURIComponent`를 의도적으로 사용합니다. 원본 코드는 `+`의
  이중 인코딩 함정을 언급합니다 — 쿼리 문자열에서는 공백을 선호하세요.
- 영속성이 없습니다 — 모든 상태는 메모리에 있으며 새로고침 시 초기화됩니다.
</content>

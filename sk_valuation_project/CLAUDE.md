# CLAUDE.md — SK LFP/ESS DC블록 플랫폼 밸류에이션 프로젝트

이 파일은 Claude Code가 이 프로젝트에서 작업할 때 가장 먼저 읽는 컨텍스트다.
미국 LFP/ESS 배터리 플랫폼(SKBA 조지아 + SKOT 테네시) 인수 검토를 위한 재무모델·파이낸싱 구조·투자 덱을 다룬다.

---

## 이 프로젝트가 하는 일

1. **재무모델** (`model/build_model3.py`): 2개 공장(SKBA·SKOT)의 FCFF/FCFE를 라인 단위로 생성하는 Python 스크립트. 결과는 4-시트 Excel(`output/FCFF_model_lines.xlsx`).
2. **투자 덱** (`deck/gen.js`, `deck/gen_v2.js`): pptxgenjs로 PPTX를 생성. v1은 클래식 스타일, v2는 SK 실적발표 템플릿(명조+IBM Plex) A4 가로.
3. **분석 문서** (`docs/`): 밸류에이션·파이낸싱 논의의 핵심 결론과 원칙.

---

## 빠른 시작 (실행법)

### 모델 생성 + 검산
```bash
cd model
python3 build_model3.py                 # → output 경로에 FCFF_model_lines.xlsx 생성
# 수식 재계산·에러 검사 (LibreOffice 필요)
pkill -9 soffice; sleep 1
python3 /path/to/recalc.py FCFF_model_lines.xlsx 95   # total_errors가 0이어야 정상
```
- `build_model3.py` 맨 아래 `wb.save("...FCFF_model_lines.xlsx")` 경로를 환경에 맞게 수정할 것.
- 의존성: `openpyxl`. 재계산엔 LibreOffice(`soffice`)가 필요.

### 덱 생성
```bash
cd deck
NODE_PATH=$(npm root -g) node gen_v2.js  # → SK_Valuation_v2.pptx (SK 템플릿, A4 가로, 권장)
NODE_PATH=$(npm root -g) node gen.js      # → SK_Valuation_Financing.pptx (v1 클래식)
# PDF 변환
soffice --headless --convert-to pdf SK_Valuation_v2.pptx
```
- 의존성: `pptxgenjs` (전역 설치). v2는 폰트 필요 — 아래 "디자인 시스템" 참조.

---

## 모델 구조 (build_model3.py)

4개 시트를 생성한다: **Assumptions, SKBA, SKOT, Consolidated**.

- **Assumptions**: 모든 입력값(★ 표시 = 핵심 가정, 노란셀). ASP·마진·AMPC·WACC·자본구조·NOL·그랜트 등.
- **SKBA / SKOT**: 공장별 P&L → NOPAT(NOL 반영) → FCFF. `plant_common()` 함수가 양쪽 공용.
- **Consolidated**: 합산 → DCF(EV) + 공장별 FCFE 브리지 + 메자닌 DSCR + NAV(SOTP).

### 핵심 함수
- `plant_common(ws, PROW, rev_row, ampc_gwh_keys, capex_key, head, nol_key)`: 공장 P&L·FCFF·NOL 생성.
- `kv(row, label, value, fmt, note)`: Assumptions 입력셀 생성(파란=입력).
- `CL(...)`: Consolidated 라인 생성. `L(...)`: 공장시트 라인 생성.

### 주의해야 할 기술적 함정 (반드시 숙지)
1. **순환참조**: 전년 동일행을 `PROW[key]`로 참조하면 KeyError(forward ref). → 행 오프셋을 변수에 저장 후 셀주소 직접참조(예: `nolbop_r+3`). NOL·메자닌 잔액이 이 패턴.
2. **L() 함수는 `green=` 인자 없음** (TypeError). 공장시트 L()과 Consolidated CL()의 시그니처가 다름.
3. **C열 노트가 "="로 시작하면 수식으로 파싱됨** → 노트 앞 "=" 제거.
4. **재계산은 시간이 오래 걸림**(110초+ 가능). `pkill -9 soffice` 후 충분한 타임아웃.

---

## 현재 기준 가정 (Base Case)

| 항목 | 값 |
|---|---|
| DC블록 ASP | $155/kWh |
| AMPC (45X 셀+모듈) | $45/kWh → 0 (2033, 법정 phase-out) |
| ex-AMPC EBITDA (개조분) | 9.0% (gross margin 14~16%에서 역산) |
| 가동률 (run-rate) ★ | BA 85% / OT 82% (DC센터 수요 확보; eff 토글=1.0) |
| SKBA 기저 EBITDA (pre-SOP floor) ★ | $0.2bn/yr, 2026~ (양산 전 기존 사업 하한; 개조분>기저면 미적용) |
| run-rate ex-AMPC EBITDA | $554mm (BA 241 + OT 313, 9% 마진) |
| WACC / Ke / TGR | 10.5% / 13% / 2% |
| 세율 | 23% (AMPC 비과세, NOL/tax carryforward 반영) |
| SKBA 명판 / SKOT 명판 | 21.2 GWh / 29.7 GWh |
| DOE loan | $4.0bn, 이자만 3.5%(또는 5%) |
| 메자닌 | **제거** (OT 소요 $1bn은 BA 증자/equity-down) |
| 순차입금 | $6.4bn = BA 3.7 + OT 2.6 (합산) |
| SKBA 기초 NOL ★ | $4.0bn (모델 입력; TCJA 80% 한도 적용) |
| CapEx 그랜트 | capex의 10% (≈$142mm), 시나리오상 최대 20% |

**결과**: DCF EV ≈ $5.0bn (가동률 85% + pre-SOP 기저 + NOL 반영), Equity(FCFE) ≈ +$2.4bn, NAV(SOTP) $5.0bn.

**역산 시나리오 (implied min EV/EBITDA, run-rate $554mm 기준)**: BASE(그랜트10/희석20) 15.4x → ① 그랜트20 14.1x → ② 희석30 14.1x → ③ 희석49 13.1x → 최대결합(그랜트20+희석49) 12.6x(peer 밴드 진입). 희석 캡을 49%까지 열면 peer 8~13x 안으로 들어오나 지배력 약화 트레이드오프(협상 상한). DCF는 WACC 10.5%·TGR 2%·Ke 13% 기준 EV $5.0bn. `model/scenarios.py`로 재현.

---

## 절대 원칙 (작업 시 반드시 지킬 것)

자세한 내용은 `docs/밸류에이션_원칙.md` 참조. 핵심만:

1. **Single source of truth = Excel.** 가정 바꾸면 덱·풋볼필드·민감도 전부 재계산. 슬라이드가 Excel과 다르면 버그.
2. **마진은 역산**: `GM = 목표 EBITDA% + SG&A%`.
3. **한시 보조금(AMPC)은 영구 자본화 금지**: 멀티플은 ex-AMPC EBITDA에만.
4. **매출↑ ≠ 가치↑** (박리다매 함정): ex-보조금 EBITDA 절대액을 본다.
5. **EV·Equity·NAV·FCFE는 다른 층위**: 순차입금은 한 번만 차감. NAV는 Equity와 비교.
6. **부채는 실제 상환구조 반영**: DOE론은 "이자만"(원금상환 0).
7. **세금은 NOL 반영**: 단 과세소득이 작으면 NOL이 많아도 무용(7장 참조).
8. **멀티엔티티는 법인별 분리**: 세금·자본구조·격리를 명시적으로.
9. **메자닌 적격성은 DSCR로**: CFADS ÷ 상환액 ≥ 1.3x.
10. **추정과 확인을 구분**: ★는 미검증. 불확실하면 민감도로 범위 제시.

---

## 자주 하는 작업 (이렇게 요청하면 됨)

- **"DOE 이자율 5%로 바꿔"** → `build_model3.py`에서 `R['doe_rate']` 입력값 수정 → 재생성 → DSCR 재확인.
- **"메자닌을 $700mm로 줄여"** → `R['pref_bal']` 수정 → DSCR이 1.3x 넘는지 검산.
- **"ex-AMPC EBITDA를 11%로"** → gross margin 4곳(1-1동/1-2동/각형/파우치) 각 +2%p → ex-AMPC EBITDA가 정확히 11%인지 Consolidated에서 확인.
- **"풋볼필드가 모델이랑 달라"** → Excel에서 EV·EBITDA·NAV 실제값 추출 → `deck/gen.js`의 rows·결론선·민감도 갱신 → PDF로 시각 검증.
- **민감도 스윕** → `build_model3.py`를 복제해 입력만 바꿔가며 배치 실행, EV/DSCR 추출(과거 `/tmp/*sweep*.py` 패턴 참조).
- **모델 재계산 + 핵심값 추출** → `python3 model/extract.py` (LibreOffice 재계산 후 run-rate ex-AMPC EBITDA·EV·Equity·total_errors 출력).
- **조달 시나리오(그랜트/희석 → implied 멀티플)** → `python3 model/scenarios.py` (재계산된 xlsx의 run-rate EBITDA를 읽어 명명 시나리오 + 3×3 민감도 매트릭스 출력).
- **"SKBA 기저 EBITDA 바꿔"** → `R['ba_base_ebitda']` 수정 → 재생성 → `extract.py`로 run-rate 확인 → 덱 S2/S4 숫자 갱신.

---

## 디자인 시스템 (deck/gen_v2.js — SK 실적발표 템플릿)

- **규격**: PPT A4 가로 (11.69 × 8.27 in)
- **배경** #FBF9F5 / **텍스트·기준선** #1B1714
- **헤드라인** Noto Serif CJK KR Bold(명조) / **본문** IBM Plex Sans KR / **수치** IBM Plex Mono
- **강조 레드** #C1121F(핵심) / **웜그레이** #D8CFC2(비교) / **그린** #1E7A4D(증감)
- **차트**: 가로 그리드만(#ECE6DC 1px), 기준선 2px, **그라데이션·3D·그림자 금지**, 값 라벨 직접 표기, 계열 최대 3개.
- 폰트 설치(없을 경우): IBM Plex Mono/Sans는 github.com/IBM/plex, IBM Plex Sans KR은 github.com/google/fonts/raw/main/ofl/ibmplexsanskr/ 에서 받아 `~/.fonts`에 두고 `fc-cache`.

---

## 알려진 한계·미해결 (다음 작업 후보)

- **SKBA NOL $4.0bn은 모델 입력(추정)** — DART 연결주석 종속기업 명세 또는 데이터룸 세무신고서로 확정 필요. NOL/tax carryforward는 `plant_common`에서 반영(TCJA 80% 한도, 당기손실 이월가산). 단 과세소득이 작아 NOL을 키워도 EV 변화 거의 없음(구조적).
- **SKBA 기저 EBITDA $0.2bn은 가정 (pre-SOP floor)** — 양산(SOP) 전 기존 사업이 매년 ~$0.2bn EBITDA를 낸다는 하한 전제(2026~). `R['ba_base_ebitda']` 입력, `plant_common`에서 `MAX(0, 기저−(gp+sga))`로 부족분만 보정 → 개조분이 기저를 넘는 run-rate 연도엔 미적용(run-rate 불변). 데이터룸으로 실제 기존 사업 EBITDA 확인 필요.
- **2033년 AMPC 소멸 후 SKOT 적자 회귀** — 메자닌은 AMPC 창 안에서 회수하는 구조라 그 이후는 미해결. 근본 해법은 마진(EBIT) 개선.
- **HTML 팩**(`sk_valuation_pack.html`)은 구버전(DC블록·9% 미반영). 필요시 갱신.
- **v2 덱**은 핵심 10슬라이드만 — v1의 GTM·Crucible 상세 일부는 미이전.

---

## 파일 맵

```
sk_valuation_project/
├── CLAUDE.md                  ← (이 파일) Claude Code 진입점
├── README.md                  ← 사람용 개요
├── docs/
│   ├── 핵심결론_종합노트.md     ← 14섹션 전체 분석 (왜 그런가)
│   └── 밸류에이션_원칙.md       ← 14개 모델링 원칙 (어떻게 하는가)
├── model/
│   └── build_model3.py        ← 모델 생성 스크립트 (현행)
├── deck/
│   ├── gen.js                 ← v1 덱 (클래식)
│   └── gen_v2.js              ← v2 덱 (SK 템플릿, A4 가로) ★ 권장
└── output/
    ├── FCFF_model_lines.xlsx  ← 모델 산출물
    ├── SK_Valuation_v2.pptx   ← 덱 산출물
    └── SK_Valuation_v2.pdf
```

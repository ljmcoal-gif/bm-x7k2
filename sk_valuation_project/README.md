# SK LFP/ESS DC블록 플랫폼 — 밸류에이션·파이낸싱 프로젝트

미국 LFP/ESS 배터리 플랫폼(SKBA 조지아 + SKOT 테네시) 인수 검토를 위한 재무모델·파이낸싱 구조·투자 덱.

> **Claude Code로 작업하려면** `CLAUDE.md`를 먼저 읽으세요. 전체 맥락·실행법·작업 규칙이 거기 있습니다.

## 무엇이 들어있나

| 폴더 | 내용 |
|---|---|
| `docs/` | 분석 결론(`핵심결론_종합노트.md`)과 모델링 원칙(`밸류에이션_원칙.md`) |
| `model/` | FCFF/FCFE 모델 생성 스크립트 (`build_model3.py`) |
| `deck/` | 투자 덱 생성기 — `gen_v2.js`(SK 템플릿·권장), `gen.js`(클래식) |
| `output/` | 산출물 — Excel 모델, PPTX/PDF 덱 |

## 핵심 결론 한 줄

> 이 딜의 가치는 "보조금(AMPC)을 걷어낸 사업 본연의 EBITDA"와 "보조금이 살아있는 동안의 부채상환능력(DSCR)"에서 나온다. 가장 근본적으로는 EBIT=0을 벗어나 D&A를 덮는 진짜 마진(ex-AMPC EBITDA 9%+)을 만들지 못하면, 2033년 AMPC 소멸 후 영구히 외부 자금에 의존하게 된다.

## 빠른 실행

```bash
# 모델
cd model && python3 build_model3.py

# 덱 (SK 템플릿)
cd deck && NODE_PATH=$(npm root -g) node gen_v2.js
```

자세한 실행법·의존성·함정은 `CLAUDE.md` 참조.

## 기준 결과

- DCF EV ≈ **$3.45bn** (NOL 반영), 밴드 $3.0~4.0bn
- NAV(SOTP) $5.0bn
- 메자닌 $1bn @ ex-AMPC EBITDA 9% → DSCR 2029~31 = 1.53/1.60/1.39x (적격)

★ 표시 수치는 미검증 추정 — 데이터룸/DART 확인 필요.

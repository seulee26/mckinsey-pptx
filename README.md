# axlabs-mckinsey-pptx

> **[AX Labs](https://theaxlabs.com/)가 만든 Claude Code 플러그인입니다.**
> 맥킨지 스타일 PPTX 생성기 — 원본 디자인(딥네이비/브라이트 블루 팔레트, 굵은 제목 + 하단 라인, 점선 섹션 마커, 출처·페이지 번호가 있는 바닥 규칙)을 충실히 재현한 40개의 프로덕션 레벨 슬라이드 템플릿을 제공합니다. `python-pptx` 기반으로 모든 도형을 네이티브로 그려 레퍼런스와 동일한 룩앤필을 유지합니다.

**Claude Code 서브에이전트**(`mckinsey-slide-agent`)가 함께 제공되어, 슬라이드마다 가장 적합한 템플릿을 선택하고 그 선택 이유를 설명한 뒤, 한 문단 짜리 브리프로부터 실제 `.pptx` 파일을 빌드합니다. `/mckinsey-deck` 슬래시 커맨드도 함께 제공됩니다.

![Author](https://img.shields.io/badge/author-AX%20Labs-0b1f3a)
![License](https://img.shields.io/badge/license-MIT-blue)
![Templates](https://img.shields.io/badge/templates-40-brightgreen)
![Platform](https://img.shields.io/badge/platform-Claude%20Code-6b4bff)

---

## Claude Code 플러그인으로 설치하기

Claude Code 세션 안에서:

```
/plugin marketplace add seulee26/mckinsey-pptx
/plugin install axlabs-mckinsey-pptx@axlabs
```

그다음 Python 의존성을 한 번만 설치하세요:

```bash
pip install -r "${CLAUDE_PLUGIN_ROOT:-.}/requirements.txt"

# 선택 사항 — 생성된 덱의 PNG 미리보기가 필요하다면
brew install --cask libreoffice   # soffice
brew install poppler              # pdftoppm
```

> 플러그인은 에이전트와 함께 Python 소스(`mckinsey_pptx/`)도 함께 배포합니다.
> 에이전트는 생성하는 모든 빌드 스크립트에서 자동으로 플러그인 루트를
> `sys.path`에 추가하므로 패키지를 전역 설치할 필요가 없습니다.

### 설치 확인

Claude Code 세션에서:

```
/agents
```

목록에 `mckinsey-slide-agent`가 보여야 합니다. 그런 다음:

```
/mckinsey-deck Q4 사업 리뷰 데크. 매출 1,200억(전년 1,050억), KPI 지연 2건, 진출 영역 3개 검토 중.
```

---

## 사용하는 두 가지 방법

### 1. Claude Code 서브에이전트 (권장) 🤖

플러그인이 설치된 Claude Code 세션에서 바로:

```
> 분기 사업 리뷰 데크 만들어줘. 매출은 1,200억(전년 1,050억),
  KPI 지연 2건, 진출 영역 3개 검토 중.
```

에이전트가 수행하는 작업:

1. `${CLAUDE_PLUGIN_ROOT}/mckinsey_pptx/agent/CATALOG.md` (템플릿 플레이북) 로드
2. 5–10 슬라이드 규모의 덱 흐름(스토리 아크) 설계
3. 각 슬라이드마다 선택한 템플릿명과 **왜 인접 템플릿이 아닌 이 템플릿을
   골랐는지** 근거 제시
4. 현재 작업 디렉터리의 `output/` 아래에 Python 빌드 스크립트를 생성하고 실행
5. PNG 미리보기를 렌더링하고, 결과 경로와 선택 근거를 리포트

에이전트가 인식하는 트리거 표현(한국어/영어):
- "맥킨지 슬라이드 / 데크 / 보고서 만들어줘"
- "전략 보고서 PPT", "사업 리뷰 데크"
- "make a deck / presentation / slides / PPTX / PowerPoint"
- "build a McKinsey-style report / strategy deck"

슬래시 커맨드로 명시적으로 호출할 수도 있습니다:

```
/mckinsey-deck <브리프>
```

### 2. Python API 직접 호출

```python
from mckinsey_pptx import PresentationBuilder

b = PresentationBuilder(default_section_marker="Q4 review")

b.add("dark_navy_summary",
      body="[Bottom line]: K-battery의 향후 5년이 글로벌 리더십을 결정합니다.")

b.add("executive_summary_takeaways",
      sections=[
          {"takeaway": "시장 YoY 22% 성장",
           "bullets": ["북미 점유율 상승", "유럽 정체"]},
      ])

b.add("column_historic_forecast",
      categories=[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026],
      values=[1035, 1108, 1153, 1148, 1206, 1265, 1381, 1430, 1535],
      forecast_from_index=5,
      historic_growth="3%", forecast_growth="6%")

b.save("output/deck.pptx")
```

한국어/혼합 콘텐츠의 경우 타이포그래피 패밀리를 `Apple SD Gothic Neo`로
변경하세요 (`examples/demo_korean.py` 참고).

---

## 실전 워크플로 — 원본 파일로 덱 만들기

실무에서는 "브리프만 한 문단"으로 끝나는 경우는 드물어요. 대개 Excel 숫자,
Word/메모 원고, PDF 회의록, 로고 이미지 등이 함께 있습니다. 이 섹션은
그런 **실제 자료를 에이전트에 물려주는 방법**을 정리합니다.

### 권장 프로젝트 디렉터리 구조

Claude Code 세션을 **덱을 만들고 싶은 프로젝트 폴더에서** 시작하세요.
에이전트는 현재 작업 디렉터리(CWD)를 기준으로 파일을 찾고, 결과물도
CWD 안 `output/`에 떨어뜨립니다.

```
my-deck-project/
├── inputs/                       # 원본 자료를 여기에 전부 모아두기
│   ├── sales_q4.xlsx             # 매출·KPI 숫자
│   ├── customer_survey.csv       # 설문 결과
│   ├── exec_brief.docx           # 경영진 브리프(Word)
│   ├── meeting_notes.md          # 회의록(마크다운/텍스트)
│   ├── board_memo.pdf            # 이사회 메모
│   └── logo.png                  # 로고 이미지
├── output/                       # 에이전트가 여기에 .pptx, 미리보기 생성
│   ├── agent_q4-review.py        # 에이전트가 만든 빌드 스크립트
│   ├── q4-review.pptx            # 최종 산출물
│   └── preview_q4-review/        # PNG 미리보기 (soffice 설치 시)
└── notes.md                      # (선택) 덱 방향성 메모
```

폴더 이름은 자유지만, **`inputs/`와 `output/`으로 분리**하는 걸 권장합니다.
에이전트는 `output/` 외에 다른 곳은 건드리지 않아요.

### 원본 파일 종류별 사용법

#### 1. Excel / CSV 데이터 (`.xlsx`, `.xls`, `.csv`)

가장 흔한 케이스. 매출, KPI, 비교표, 설문 결과 등.

**에이전트에게 지시하는 법:**

```
> inputs/sales_q4.xlsx의 '월별매출' 시트 데이터로 Q4 사업 리뷰 데크 만들어줘.
  3번째 슬라이드는 historic + forecast 차트(2026부터 전망)로 해주고,
  KPI 대시보드 슬라이드도 하나 넣어줘.
```

에이전트는:
1. `openpyxl` 또는 `pandas`로 Excel을 **직접 읽음** (Claude Code의 Bash/Read 도구 사용)
2. 필요한 숫자만 추출해서 템플릿의 `values`, `categories` 인자로 꽂음
3. 읽은 값을 리포트에 인용 ("매출 1,265억은 Sheet1 B13에서 가져왔습니다")

> 💡 **팁**: 시트명·셀 범위를 구체적으로 말해주면 추출 정확도가 올라갑니다.
> "Sheet1의 A1:D13", "월별매출 시트의 2026년 열" 등.

만약 의존성이 없으면 에이전트가 알아서 설치 시도합니다:
```bash
pip install openpyxl pandas
```

#### 2. Word 문서 (`.docx`, `.doc`)

브리프, 기안서, 보고서 초안을 Word로 들고 오는 경우.

```
> inputs/exec_brief.docx 읽고 그 내용을 바탕으로
  경영진용 10슬라이드 전략 보고서 만들어줘.
  도입부는 dark_navy_summary로, 결론은 prioritization_matrix로.
```

에이전트는 `python-docx`로 본문/제목/표를 추출하고, 문장을
**불릿으로 재구성**(bullet-ify)합니다 — 긴 문단을 슬라이드 친화적인
짧은 키워드로 압축하는 게 기본 동작.

> 💡 `.doc`(구버전)은 먼저 `.docx`로 저장 변환해주세요. python-docx는
> OOXML만 지원합니다.

#### 3. PDF (`.pdf`)

회의록, 보고서, 외부 리서치 등.

```
> inputs/board_memo.pdf 3~7페이지 요약해서
  이사회용 임원 요약 슬라이드 만들어줘.
```

에이전트는 `pdftotext`(poppler) 또는 `pypdf`로 텍스트 추출 → 요약 →
`executive_summary_takeaways` 같은 템플릿에 삽입합니다. 표가 많은
PDF는 텍스트 레이아웃이 망가지는 경우가 있어 — 그럴 땐 **해당 숫자를
Excel/CSV로도 같이 주세요**.

#### 4. 마크다운 / 텍스트 (`.md`, `.txt`)

회의록·브레인스토밍·메모. 가장 부담 없음.

```
> notes.md 참고해서 킥오프 덱 만들어줘. 워크플로 섹션은
  process_flow_horizontal으로.
```

#### 5. 이미지 / 로고 (`.png`, `.jpg`, `.svg`)

**주의**: 현재 템플릿들은 로고/이미지 삽입을 **기본 파라미터로 받지 않습니다.**
로고를 덱에 넣고 싶다면 두 가지 중 하나:

- **(A) 빌드 스크립트에서 직접 삽입**: 에이전트에게 "빌드 스크립트 만든 뒤
  `slides[0].shapes.add_picture('inputs/logo.png', ...)` 한 줄 추가해줘"라고
  지시하면 됩니다.
- **(B) 생성된 .pptx를 PowerPoint/Keynote로 열어서 수동 삽입**. 템플릿이
  로고 영역을 비워뒀다고 생각하고 열어서 드래그.

### 실제 시나리오 예시

#### 시나리오 A. 분기 사업 리뷰

```
my-deck/
├── inputs/
│   ├── q4_financials.xlsx    # 매출·영업이익·KPI 숫자
│   └── risks.md              # 리스크 6개 리스트
└── (빈 output/은 에이전트가 생성)
```

**지시:**
```
> 이 폴더의 inputs/ 자료로 Q4 사업 리뷰 데크 7슬라이드 만들어줘.
  경영진 대상, 한국어, 결론은 투자 승인 요청.
  q4_financials.xlsx의 '요약' 시트에서 매출·영업이익 숫자 써줘.
  리스크 슬라이드는 prioritization_matrix로.
```

#### 시나리오 B. 시장 진입 킥오프

```
my-deck/
├── inputs/
│   ├── market_research.pdf   # 외부 리서치 15페이지
│   ├── customer_icp.xlsx     # 타겟 고객 프로필
│   └── roadmap.md            # 12주 계획
```

**지시:**
```
> 인도네시아 시장 진출 킥오프 데크 만들어줘. 영문, 10슬라이드.
  inputs/market_research.pdf 요약 + customer_icp.xlsx ICP +
  roadmap.md를 gantt_timeline으로.
```

#### 시나리오 C. 숫자만 있고 브리프가 없는 경우

```
my-deck/
└── inputs/
    └── numbers.csv           # 5년치 매출 + 3년 전망
```

**지시:**
```
> inputs/numbers.csv만 보고 이걸 보여줄 수 있는 5슬라이드 덱 만들어줘.
  숫자가 말해주는 이야기를 찾아서 스토리아크 짜줘.
```

→ 에이전트가 숫자의 추세·변곡점을 스스로 파악하고, 슬라이드
   흐름까지 설계합니다.

### 반복 개선 (Iteration)

첫 빌드는 초안일 뿐. 다음처럼 이어서 요청하면 됩니다:

```
> 슬라이드 4를 prioritization_matrix 말고 growth_share으로 바꿔줘.
> 슬라이드 2의 takeaway 세 번째 불릿을 "NPS 42점(업계 평균 +15)"로 수정.
> 전체 덱 톤을 더 단정하게 — 이모지·느낌표 전부 제거.
> 이 덱 영문판도 하나 만들어줘. 같은 구조, 같은 숫자.
```

에이전트는 변경 요청도 **어떤 템플릿으로 왜 바꾸는지** 근거와 함께 적용합니다.

### 자주 묻는 실무 질문

**Q. Excel 시트가 수십 개면 어떻게 해요?**
A. 쓸 시트와 셀 범위를 지목해주세요. "요약" 시트의 A1:F20만 본다면
   그만큼만 읽고 다른 시트는 무시합니다.

**Q. 기밀 데이터인데 클라우드로 안 올라가죠?**
A. 플러그인은 전부 **로컬에서 실행**됩니다. Claude Code 자체가 외부 API
   호출을 하긴 하지만, 당신의 파일은 로컬에서 파싱·처리되고 API에는
   모델 응답 생성에 필요한 컨텍스트만 전달됩니다. 보안 정책은 회사
   가이드라인을 따르세요.

**Q. 결과물이 마음에 안 드는 레이아웃이면?**
A. (1) 같은 템플릿 목록을 `${CLAUDE_PLUGIN_ROOT}/mckinsey_pptx/agent/CATALOG.md`에서
   훑어보고 (2) 이름을 지정해서 "이 슬라이드 X 템플릿으로 다시"라고
   말하세요. 40개 템플릿 중 가장 적절한 걸 고르면 레이아웃이 바뀝니다.

**Q. 회사 로고/브랜드 색상으로 바꾸고 싶어요.**
A. `mckinsey_pptx/theme.py`의 `Theme` 데이터클래스를 복제·수정해서 전달하면
   됩니다. 엔터프라이즈 버전(프라이빗 템플릿)은 AX Labs로 문의하세요.

---

## 슬라이드 템플릿 카탈로그 (총 40개)

전체 카탈로그(**언제 써야 하는지 / 언제 쓰면 안 되는지 / 필수 입력 / 예시**)는 [`mckinsey_pptx/agent/CATALOG.md`](mckinsey_pptx/agent/CATALOG.md)에 정리되어 있습니다.

### 임팩트 요약

| 키 | 용도 |
|---|---|
| `executive_summary_paragraph` | 제목 + 2–4개의 본문 블록 |
| `executive_summary_takeaways` | 굵은 takeaway → 불릿, 선택적 최종 결론 |
| `dark_navy_summary` | 풀블리드 딥네이비 한 줄 임팩트 스테이트먼트 |

### 상태 · 진단

| 키 | 용도 |
|---|---|
| `assessment_table` | 카테고리 × KPI 표 + 신호등 도트 |

### 차트

| 키 | 용도 |
|---|---|
| `column_comparison` | 정렬 막대 + 포커스 하이라이트 + 우측 패널 |
| `column_simple_growth` | 시계열 + 단일 성장 화살표 |
| `column_split_growth` | 시계열 + 2단계 성장 구분 |
| `column_historic_forecast` | 실적(네이비) + 전망(라이트 블루) 막대 |
| `bubble_chart` | 풀폭 XY 산점도, 그룹별 색상 |
| `bubble_chart_takeaways` | 위와 동일 + 우측 불릿 |

### 매트릭스

| 키 | 용도 |
|---|---|
| `growth_share` (`bcg_matrix`) | 2×2 BCG 사분면 + 버블 |
| `prioritization_matrix` | 3×3 Time × Impact 그리드, 상태 색상 |

### 트렌드 · 영역

| 키 | 용도 |
|---|---|
| `three_trends_icons` | 3행: 원형 아이콘 + 라벨 + 불릿 |
| `three_trends_table` | 3행: 이름 필 + 설명 + 예시 |
| `three_trends_numbered` | 3행: 번호 + 블루 라벨 필 + 불릿 |
| `five_key_areas` | 번호가 매겨진 5행, 화살표 + 설명 |
| `overview_areas` | 5–7개의 세로 영역 카드, 알파벳 배지 |

### 조직 · 위계

| 키 | 용도 |
|---|---|
| `issue_tree` | 루트 이슈 → 메인 → 세컨더리 → 하위 드라이버 |
| `org_chart` | CEO → N명의 헤드 → 헤드별 팀원 |
| `project_team_circles` | 리더 1명 + N명의 라벨링된 팀원 원형 |
| `team_chart` | 기능 컬럼 × 역할 타일 (채움/아웃라인) |

### 로드맵 · 프로세스

| 키 | 용도 |
|---|---|
| `phases_chevron_3` | 3개의 셰브론 화살표 단계 + 산출물/담당 |
| `phases_table_4` | 4컬럼: 설명 + 활동 + 결과 |
| `waves_timeline_4` | 가로 화살표 위의 4개 웨이브, 마커 포함 |
| `gantt_timeline` | 주간 간트, 워크스트림 행 + 마일스톤 |
| `process_activities` | 3행 표: 활동 / 관리 / 산출물 |
| `process_flow_horizontal` | 4–6개의 번호가 매겨진 셰브론 타일 + 설명 |
| `funnel` | 하향식 퍼널 (TAM/SAM/SOM, 마케팅 퍼널 등) |

### 구조적 슬라이드 (덱 뼈대)

| 키 | 용도 |
|---|---|
| `cover_slide` | 타이틀 페이지, 클라이언트 + 날짜 + 악센트 스트라이프 |
| `section_divider` | 챕터 구분(큰 번호 + 챕터명) |
| `agenda` | 번호가 매겨진 목차, 활성 강조 |
| `stat_hero` | 한 개의 큰 숫자 + 라벨 + 맥락 |
| `quote_slide` | 큰 인용 + 출처 |

### 비교

| 키 | 용도 |
|---|---|
| `comparison_table` | 옵션 × 기준, 하비볼 레이팅 |
| `pros_cons` | 2컬럼 ✓ 초록 / ✗ 빨강 분석 |
| `two_column_compare` | Before/After 또는 As-is/To-be, 화살표 |

### 추가 차트

| 키 | 용도 |
|---|---|
| `stacked_column_chart` | 시간/카테고리별 구성 |
| `grouped_column_chart` | 카테고리별 나란한 비교 |
| `line_chart` | 1–4개 시리즈 시계열, 마커 포함 |
| `kpi_dashboard` | 4–8개의 KPI 타일 + 델타 (▲▼▬) |

---

## 플러그인 없이 데모 실행하기

레포지토리를 클론하고 requirements를 설치한 뒤:

```bash
# 영문 데모 (15 슬라이드)
python -m examples.demo

# 한국어 데모 (21 슬라이드 — K-배터리 글로벌 전략)
python -m examples.demo_korean
```

결과물은 `output/`에 생성됩니다.

---

## 적응형 Direct API

서브에이전트 없이도 템플릿 추론을 쓰고 싶다면,
`PresentationBuilder.add_specs()`가 dict 리스트를 받아 각 dict의 형태로부터
템플릿을 자동으로 추론합니다:

```python
b.add_specs([
    {"body": "[Bottom line]: ..."},                              # → dark_navy_summary
    {"sections": [...]},                                          # → executive_summary_takeaways
    {"categories": [{"name":"...", "rows":[...]}]},              # → assessment_table
    {"main_drivers": [...]},                                      # → issue_tree
    {"weeks": [...], "workstreams": [...]},                       # → gantt_timeline
    {"categories": [...], "values": [...], "forecast_from_index": 5},
                                                                  # → column_historic_forecast
])
```

`{"type": "<template_id>", ...}`로 명시 오버라이드도 가능합니다.

---

## 테마 커스터마이즈

`mckinsey_pptx/theme.py`를 수정하거나, `PresentationBuilder`에 커스텀
`Theme(...)`을 전달하세요:

```python
from dataclasses import replace
from mckinsey_pptx import DEFAULT_THEME, PresentationBuilder
from mckinsey_pptx.theme import Typography

KO_THEME = replace(
    DEFAULT_THEME,
    typography=replace(DEFAULT_THEME.typography, family="Apple SD Gothic Neo"),
    copyright_text="ⓒ 2026 AX Labs",
)

b = PresentationBuilder(theme=KO_THEME)
```

---

## CLI

```bash
python -m mckinsey_pptx.cli --list-types
python -m mckinsey_pptx.cli --demo -o output/demo.pptx
python -m mckinsey_pptx.cli specs.json -o deck.pptx \
       --section-marker "Strategy review"
```

---

## 프로젝트 구조

```
mckinsey-pptx/
├── .claude-plugin/
│   ├── plugin.json              # Claude Code 플러그인 매니페스트 (name, version, author)
│   └── marketplace.json         # AX Labs 마켓플레이스 엔트리 (단일 플러그인)
├── agents/
│   └── mckinsey-slide-agent.md  # 서브에이전트 정의 (플러그인에서 자동 로드)
├── commands/
│   └── mckinsey-deck.md         # /mckinsey-deck 슬래시 커맨드
├── mckinsey_pptx/               # Python 구현체
│   ├── theme.py                 # 팔레트 / 폰트 / 레이아웃 토큰
│   ├── base.py                  # 공통 도형·텍스트 헬퍼 + 슬라이드 크롬
│   ├── builder.py               # PresentationBuilder + 적응형 라우팅
│   ├── cli.py
│   ├── slides/                  # 모듈별 40개 템플릿
│   └── agent/
│       └── CATALOG.md           # 서브에이전트가 읽는 템플릿 플레이북
├── examples/
│   ├── demo.py                  # 영문 15 슬라이드 데모
│   └── demo_korean.py           # 한국어 21 슬라이드 데모
├── LICENSE                      # MIT — © 2026 AX Labs (이승필)
├── README.md
└── requirements.txt
```

---

## 제작자 소개

**AX Labs — 이승필 (Seungpil Lee)** 가 제작·유지보수합니다.

AX Labs는 팀들이 AI 네이티브 컨설팅 워크플로(AX = *AI eXperience / AI
Transformation*)를 도입하도록 돕습니다. 이 플러그인은 AX Labs가 내부적으로
수 일이 아닌 수 분 만에 컨설팅 덱을 초안화하기 위해 쓰는 툴킷의 오픈소스
일부입니다.

### 엔터프라이즈 AX 문의

조직에서 다음이 필요하다면:

- **프라이빗 슬라이드 템플릿** — 회사 고유 비주얼 아이덴티티에 맞춘 커스텀
- **도메인 튜닝 서브에이전트** — 산업별 플레이북, 내부 프레임워크,
  자체 데이터 소스 연동
- **엔드-투-엔드 AX 트랜스포메이션** — 진단 → 파일럿 → 기능 전반 배포
- **온프레미스 / VPC 배포** — 에이전트 파이프라인 전반
- **교육 및 내재화** — 파트너/컨설팅 팀 대상

→ 연락처: **help@theuxlabs.com** (제목: `[AX 문의] <회사명>`)

오픈소스 기여, 버그 리포트, 템플릿 요청은 GitHub 이슈로 환영합니다.

---

## 라이선스

[MIT](./LICENSE) © 2026 AX Labs — 이승필 (Seungpil Lee)

MIT 라이선스는 여기 공개된 오픈소스 플러그인에 적용됩니다. 커스텀 확장,
프라이빗 템플릿, 엔터프라이즈 배포는 별도의 상업적 조건으로 운영됩니다 —
AX Labs로 문의 주세요.

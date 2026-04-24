# axlabs-mckinsey-pptx

> **[AX Labs](https://theaxlabs.com/)가 만든 Claude Code 플러그인입니다.**
> 맥킨지 스타일의 고퀄리티 PPT를 **채팅으로 한 줄 말하면 자동으로 만들어 드립니다.**
> "Q4 사업 리뷰 데크 만들어줘"라고 입력하면 40개의 전문 템플릿 중에서
> 적절한 걸 골라 표지·차트·매트릭스·로드맵을 넣은 완성된 `.pptx` 파일이
> 폴더에 생깁니다. 파워포인트를 열어서 드래그할 필요도, 디자인 감각이
> 필요도 없습니다.

![Author](https://img.shields.io/badge/author-AX%20Labs-0b1f3a)
![License](https://img.shields.io/badge/license-MIT-blue)
![Templates](https://img.shields.io/badge/templates-40-brightgreen)
![Platform](https://img.shields.io/badge/platform-Claude%20Code-6b4bff)

---

## 이게 뭔가요? (1분 요약)

- **무엇을 하는 도구:** 글로만 지시해도 컨설팅 회사 스타일 PPT를 자동으로 만들어주는 도구예요.
- **누가 쓰면 좋은가:** 기획자·마케터·임원·컨설턴트·학생 — **파워포인트를 일일이 만들기 지겨운 누구나.**
- **어떻게 쓰는가:** Claude Code(채팅 AI 도구) 안에서 **한국어로 말하듯** 요청하면 됩니다.
  개발 지식이 필요 없고, 명령어 암기도 필요 없어요.
- **결과물:** 진짜 `.pptx` 파일. 파워포인트로 열어서 편집하고, 메일로 보내고,
  발표까지 바로 가능합니다.

**예시 한 줄:**
> "Q4 사업 리뷰 데크 만들어줘. 매출 1,200억, 전년 대비 14% 성장, KPI 지연 2건."

↓ 30초~1분 후 ↓

→ `output/q4-review.pptx` 파일이 생성됩니다. 표지 + 요약 + 성장 차트 + 이슈
  매트릭스 + 로드맵 + 결론까지 6~8장의 맥킨지 스타일 슬라이드가 들어가요.

---

## 처음 설치하는 분을 위한 준비 (5분)

이 도구는 **Claude Code**라는 프로그램 안에서 동작합니다. Claude Code는
"명령줄에서 쓰는 AI 비서"라고 생각하시면 돼요.

### 1단계. Claude Code 설치

아직 없다면 [Claude Code 공식 다운로드](https://claude.com/claude-code)에서
설치하세요. Mac / Windows / Linux 전부 지원합니다.

### 2단계. 이 플러그인 설치

Claude Code를 실행한 뒤, **채팅창에** 아래 두 줄을 **한 줄씩** 입력하세요:

```
/plugin marketplace add seulee26/mckinsey-pptx
```

```
/plugin install axlabs-mckinsey-pptx@axlabs
```

> "Installed axlabs-mckinsey-pptx"라는 메시지가 나오면 성공입니다.

### 3단계. 필요한 도구 한 번만 설치

Claude에게 아래처럼 말씀하세요. Claude가 알아서 필요한 걸 설치해줍니다:

```
이 플러그인이 쓸 파이썬 라이브러리들 설치해줘.
```

PPT를 이미지로 미리보기하고 싶다면 추가로(선택):

- Mac: Claude에게 "`brew install --cask libreoffice && brew install poppler` 실행해줘"
- 그 외 OS: LibreOffice를 일반 설치 방식으로 깔면 됩니다

> 💡 미리보기 도구가 없어도 `.pptx` 파일은 정상적으로 만들어집니다.
> 그냥 파워포인트나 Keynote로 열면 돼요.

### 설치 확인

Claude Code 채팅창에 다음을 입력:

```
/agents
```

목록에 `mckinsey-slide-agent`가 보이면 준비 완료입니다.

---

## 쓰는 법 — 채팅창에 말만 걸면 끝

### 가장 간단한 방법

Claude Code 채팅창에 그냥 **한국어로 원하는 덱을 설명**하세요:

```
분기 사업 리뷰 데크 만들어줘. 매출 1,200억(전년 1,050억),
KPI 지연 2건, 진출 영역 3개 검토 중.
```

↓

Claude가 알아서:

1. 덱의 흐름(표지 → 요약 → 차트 → 결론)을 **스토리 구조**로 설계
2. 각 슬라이드에 맞는 템플릿을 40개 중에서 골라줌
3. 고른 이유까지 **"왜 이 템플릿인지"** 설명
4. `output/` 폴더에 완성된 `.pptx` 파일 저장

### 이런 말도 알아듣습니다

- "맥킨지 스타일로 전략 보고서 만들어줘"
- "사업 리뷰 데크 짜줘, 경영진 대상으로"
- "킥오프 덱 10슬라이드, 영문으로"
- "5년치 매출 추이 보여주는 한 장짜리 슬라이드"
- "BCG 매트릭스로 4개 사업 비교해줘"

### 완성되면 이렇게 확인

결과 파일 경로를 Claude가 알려줍니다. 예: `output/q4-review.pptx`.
Finder(Mac) / 탐색기(Windows)에서 열어보면 됩니다.
더블클릭하면 파워포인트나 Keynote로 열립니다.

---

## 실전 가이드 — 엑셀·워드 원본으로 덱 만들기

"한 문단짜리 설명"으로 끝나는 일은 드물죠. 보통은 **엑셀 매출 데이터**,
**워드로 정리한 기획안**, **PDF 리서치 자료** 같은 원본이 있습니다.
아래처럼 폴더를 정리하고 Claude에게 보여주면 자동으로 내용을 읽어서
덱에 녹여줍니다.

### 폴더 정리 방법 (중요!)

Claude Code를 **덱을 만들고 싶은 프로젝트 폴더**에서 실행하세요.
Claude는 그 폴더 안에 있는 파일들만 읽습니다.

권장 구조:

```
내_프로젝트_폴더/
├── inputs/                    ← 원본 파일은 전부 여기에
│   ├── 매출데이터.xlsx
│   ├── 기획안.docx
│   ├── 회의록.pdf
│   └── 로고.png
└── output/                    ← 완성된 PPT는 여기에 저장됨 (자동 생성)
```

폴더 이름은 한글이든 영어든 상관없어요. **원본과 결과물을 분리**하는 게 핵심.

### 파일 종류별 활용법

#### 📊 엑셀·CSV 파일 (매출, 설문, KPI 숫자)

Claude에게 **어느 시트**의 **어느 숫자**를 쓸지 알려주면 됩니다:

```
inputs/매출데이터.xlsx의 '월별매출' 시트 숫자로 Q4 리뷰 덱 만들어줘.
1월부터 12월까지 월별 매출을 막대 차트로 보여주고,
KPI 대시보드도 한 장 넣어줘.
```

Claude가 엑셀 파일을 직접 열어 숫자를 읽고, 차트로 그려줍니다.
**"어느 숫자를 어디서 가져왔는지"도 리포트해줍니다** — 검증하기 쉬움.

> 💡 **팁:** 시트 이름·셀 범위(예: "A1부터 F20까지")를 알려주면 더 정확해요.
> 시트가 수십 개면 쓸 것만 지목해주세요.

#### 📝 워드 문서 (`.docx`)

기획안, 보고서 초안, 회의록 등. Claude가 **문장을 읽어서 슬라이드용
짧은 키워드로 자동 축약**해줍니다.

```
inputs/기획안.docx 읽고 경영진용 10슬라이드 전략 보고서로 만들어줘.
결론은 투자 승인 요청으로.
```

> ⚠️ 구버전 `.doc` 파일은 워드에서 먼저 **`.docx`로 저장**해주세요.

#### 📄 PDF 파일

회의록, 외부 리서치, 이사회 자료 등.

```
inputs/이사회메모.pdf 3~7페이지 요약해서
임원용 한 장짜리 요약 슬라이드 만들어줘.
```

> 💡 PDF에 **숫자 표가 많다면** 엑셀로도 같이 주세요. PDF 표는 가끔
> 레이아웃이 깨져서 읽혀요. 엑셀이 있으면 Claude가 엑셀 쪽을 우선 씁니다.

#### 🗒 메모·마크다운 (`.md`, `.txt`)

아이디어 메모, 브레인스토밍 초안 — 가장 부담 없어요.

```
notes.md 보고 킥오프 덱 만들어줘.
```

#### 🖼 로고·이미지

로고를 넣고 싶다면 두 가지 방법:

1. **Claude에게 한 줄 요청:** "표지에 `inputs/로고.png` 넣어줘"
2. **직접 넣기:** 완성된 `.pptx`를 파워포인트로 열어서 로고를 드래그

### 실제 사용 시나리오

#### 시나리오 A. 월말에 분기 사업 리뷰 덱 준비

```
프로젝트폴더/
└── inputs/
    ├── 4분기_재무.xlsx
    └── 리스크_리스트.md
```

Claude에게:

```
이 폴더의 inputs/ 자료로 Q4 사업 리뷰 데크 7슬라이드 만들어줘.
경영진 대상, 결론은 투자 승인 요청.
4분기_재무.xlsx의 '요약' 시트에서 매출·영업이익 숫자 써주고,
리스크는 중요도/시급성 매트릭스로 보여줘.
```

#### 시나리오 B. 해외 시장 진입 킥오프 덱

```
프로젝트폴더/
└── inputs/
    ├── 시장조사.pdf         (외부 리서치 15페이지)
    ├── 타겟고객.xlsx
    └── 로드맵.md            (12주 계획)
```

Claude에게:

```
인도네시아 시장 진출 킥오프 덱 만들어줘. 영문, 10슬라이드.
시장조사.pdf 요약 + 타겟고객.xlsx 고객 프로필 +
로드맵.md를 간트 차트로.
```

#### 시나리오 C. 숫자만 있고 뭘 보여줘야 할지 모를 때

```
프로젝트폴더/
└── inputs/
    └── 5년_매출.csv
```

Claude에게:

```
inputs/5년_매출.csv만 보고, 이 숫자들이 말하는 이야기를 찾아서
5슬라이드로 보여줘. 어떤 메시지가 중요한지 네가 판단해줘.
```

→ Claude가 추세·변곡점·이상치를 스스로 발견해서 스토리를 짜줍니다.

### 마음에 안 들면? 바로 수정 가능

첫 결과는 **초안**이에요. 그대로 채팅을 이어가면서:

```
슬라이드 4를 다른 레이아웃으로 바꿔줘. 숫자 비교가 더 잘 보이게.
```

```
슬라이드 2의 세 번째 불릿을 "NPS 42점(업계 평균 +15)"로 수정.
```

```
전체 톤을 더 단정하게. 이모지랑 느낌표 다 빼줘.
```

```
이 덱 영문판도 같이 만들어줘. 구조랑 숫자는 똑같이.
```

이런 식으로 대화하듯 계속 다듬을 수 있습니다.

---

## 자주 묻는 질문

**Q. 파이썬, 코딩 한 번도 안 해봤는데 괜찮나요?**
A. 네, **전혀 필요 없어요.** Claude Code만 설치하면 전부 한국어 대화로 끝납니다.

**Q. 기밀 데이터인데 안전한가요?**
A. 파일은 **여러분 컴퓨터 안에서만** 처리됩니다. 엑셀·워드를 클라우드에
   업로드하지 않아요. (단, Claude와의 대화 내용은 AI 응답을 받기 위해
   Anthropic 서버에 전달됩니다. 회사 보안 정책에 따라 판단해주세요.)

**Q. 레이아웃이 마음에 안 들어요.**
A. "이 슬라이드 다른 레이아웃으로 바꿔줘"라고만 말하면 40개 템플릿 중
   다른 걸 골라 다시 그려줍니다. 구체적으로 "숫자 크게 보여주는 걸로"
   같은 힌트를 주면 더 빠르게 좋은 결과가 나와요.

**Q. 회사 브랜드 색상, 로고로 바꾸고 싶어요.**
A. 기본 테마는 맥킨지 스타일(딥네이비)입니다. 커스텀 브랜드 템플릿은
   AX Labs 엔터프라이즈 계약으로 제공합니다 — 아래 **엔터프라이즈 문의**
   참고.

**Q. 슬라이드 개수를 지정할 수 있나요?**
A. 네. "7슬라이드로", "10장짜리로" 같이 말씀하시면 됩니다. 지정 안 하면
   내용에 맞춰 보통 5~10장으로 만들어줍니다.

**Q. 영어로 된 덱도 만들 수 있나요?**
A. 물론요. 한국어로 요청해도 "영문으로 만들어줘"라고 하면 영어 덱을 만들고,
   영어로 요청하면 영어 덱을 만듭니다.

**Q. 완성된 PPT를 편집할 수 있나요?**
A. 네, **일반 파워포인트 파일**이에요. 파워포인트/Keynote로 열어서
   마음대로 수정·발표하세요.

---

## 어떤 슬라이드를 만들어주나요? (40개 템플릿)

Claude가 알아서 고르긴 하지만, 원하면 직접 지정할 수도 있습니다.

### 🎯 요약·결론 슬라이드
- **핵심 메시지 한 줄** (딥네이비 풀스크린)
- **경영진용 요약** (굵은 결론 + 불릿)
- **파라그래프 요약** (제목 + 2~4개 단락)

### 📊 차트·숫자 시각화
- **시계열 성장 차트** (성장 화살표 포함)
- **실적 + 전망 차트** (과거는 진한색, 미래는 연한색)
- **비교 막대 차트** (하이라이트 강조)
- **버블 차트**, **누적 차트**, **그룹 비교 차트**, **라인 차트**
- **KPI 대시보드** (4~8개 타일, 전년 대비 ▲▼)

### 🧩 매트릭스·프레임워크
- **BCG 성장-점유율 매트릭스** (2×2 사분면)
- **우선순위 매트릭스** (3×3 시급성 × 중요도)
- **비교 테이블** (옵션 × 기준, 하비볼 평가)
- **장단점 분석** (✓ 초록 / ✗ 빨강)
- **Before/After 비교**

### 🏢 조직·팀 구조
- **조직도** (CEO → 헤드 → 팀원)
- **팀 원형 차트** (리더 + 팀원)
- **기능별 팀 매트릭스**
- **이슈 트리** (문제 → 원인 → 근본 원인)

### 🗓 로드맵·프로세스
- **3단계 셰브론 단계**
- **4단계 상세 테이블**
- **4개 웨이브 타임라인**
- **간트 차트** (주간 단위 + 마일스톤)
- **프로세스 흐름도** (4~6개 단계)
- **퍼널** (TAM/SAM/SOM 등)

### 📋 구조적 슬라이드
- **표지 슬라이드**
- **챕터 구분**
- **목차**
- **한 개 큰 숫자 강조**
- **인용 슬라이드**

### 🎨 기타
- **3가지 트렌드** (아이콘/테이블/번호 3종 스타일)
- **5가지 핵심 영역**
- **개요 영역 카드** (5~7개)
- **상태 평가 테이블** (신호등 색상)

> 전체 상세 목록은 [`mckinsey_pptx/agent/CATALOG.md`](mckinsey_pptx/agent/CATALOG.md)에 있습니다.
> 평소엔 Claude에게 "사업 비교 슬라이드" 같이 **목적만** 말해도 알아서 골라요.

---

## 이 플러그인은 어떻게 구성되어 있나요?

설치하면 이런 폴더 구조가 들어옵니다. 비개발자라면 **"그냥 이런 게 들어있구나"**
정도로 알고 계셔도 충분해요. 모두 **필수 구성요소**이니 지우지 마세요.

```
axlabs-mckinsey-pptx/
├── .claude-plugin/            ← 플러그인 명함 (Claude Code가 읽는 정보)
├── agents/                    ← ★ AI 비서의 "두뇌" (여기 있어야 말을 알아들어요)
│   └── mckinsey-slide-agent.md
├── commands/                  ← 슬래시 명령어 정의 (/mckinsey-deck 같은 것)
│   └── mckinsey-deck.md
├── mckinsey_pptx/             ← 실제 PPT를 그리는 엔진 (40개 템플릿)
├── examples/                  ← 참고용 샘플 덱
└── requirements.txt           ← 자동 설치용 도구 목록
```

**핵심 세 가지:**

- **`agents/` 폴더** — "맥킨지 슬라이드 에이전트"라는 AI 비서의 성격·지식·
  일하는 방식이 적힌 파일이 들어있어요. Claude Code는 여기 있는 파일을 읽어서
  **"이 AI가 뭘 잘하고, 어떻게 일해야 하는지"**를 학습합니다. 없으면 "덱
  만들어줘"라고 말해도 그냥 일반 AI가 대답할 뿐, 전문 에이전트는 깨어나지
  않아요.
- **`commands/` 폴더** — `/mckinsey-deck` 같은 **슬래시 커맨드**가 정의된 곳.
  채팅창에 `/`를 누르면 뜨는 명령어 자동완성의 소스예요.
- **`mckinsey_pptx/` 폴더** — 실제로 `.pptx` 파일을 그리는 **그래픽 엔진**.
  40가지 슬라이드 템플릿(차트·매트릭스·로드맵 등)의 디자인·색상·레이아웃
  코드가 여기에 있어요. 에이전트가 덱을 만들 때 내부적으로 이 엔진을
  호출합니다.

세 폴더가 **한 세트로 움직여야** 플러그인이 동작합니다. 이 중 하나라도
빠지면 "에이전트만 있고 그릴 수 없음" 또는 "그릴 수는 있는데 대화로
부를 수 없음" 상태가 돼요.

---

## 엔터프라이즈 문의 — AX Labs

이 오픈소스 플러그인이 유용하셨다면, 조직 단위로도 도입하실 수 있습니다.

**AX Labs는 이런 것도 제공합니다:**

- **회사 고유 브랜드 템플릿** — 귀사 로고·색상·폰트에 맞춘 전용 40개 템플릿
- **산업 특화 AI 에이전트** — 귀사 내부 프레임워크·플레이북·데이터 소스 학습
- **엔드투엔드 AX 도입** — 진단부터 기능 전반 배포까지
- **온프레미스 / VPC 배포** — 기밀 데이터를 절대 외부로 내지 않는 환경
- **팀 교육·내재화** — 컨설팅·기획팀 대상 워크숍

**AX = AI eXperience / AI Transformation.** 수 일 걸리던 컨설팅 덱을
수 분에 초안화하는 게 시작이고, 기획·리서치·분석 워크플로 전체를
AI 네이티브로 바꿔드립니다.

→ 연락처: **help@theuxlabs.com**
→ 웹사이트: [theaxlabs.com](https://theaxlabs.com/)
→ 제목 양식: `[AX 문의] <회사명>`

---

## 기여·문의

- **버그 리포트·템플릿 제안:** GitHub 이슈 환영합니다
- **오픈소스 기여:** Pull Request 환영
- **엔터프라이즈 문의:** help@theuxlabs.com

---

## 라이선스

[MIT](./LICENSE) © 2026 AX Labs — 이승필 (Seungpil Lee)

오픈소스 플러그인에는 MIT 라이선스가 적용됩니다. 커스텀 확장·프라이빗
템플릿·엔터프라이즈 배포는 별도의 상업적 조건으로 운영됩니다 —
AX Labs로 문의 주세요.

---

<details>
<summary>🛠 개발자용 (Python API, CLI, 프로젝트 구조)</summary>

### Python API 직접 호출

```python
from mckinsey_pptx import PresentationBuilder

b = PresentationBuilder(default_section_marker="Q4 review")
b.add("dark_navy_summary", body="[Bottom line]: K-battery의 향후 5년이 글로벌 리더십을 결정합니다.")
b.add("executive_summary_takeaways",
      sections=[{"takeaway": "시장 YoY 22% 성장",
                 "bullets": ["북미 점유율 상승", "유럽 정체"]}])
b.add("column_historic_forecast",
      categories=[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026],
      values=[1035, 1108, 1153, 1148, 1206, 1265, 1381, 1430, 1535],
      forecast_from_index=5, historic_growth="3%", forecast_growth="6%")
b.save("output/deck.pptx")
```

### 적응형 API (`add_specs`)

dict의 모양으로 템플릿 자동 추론:

```python
b.add_specs([
    {"body": "[Bottom line]: ..."},                              # dark_navy_summary
    {"sections": [...]},                                          # executive_summary_takeaways
    {"categories": [...], "values": [...], "forecast_from_index": 5},
                                                                  # column_historic_forecast
])
```

### CLI

```bash
python -m mckinsey_pptx.cli --list-types
python -m mckinsey_pptx.cli --demo -o output/demo.pptx
python -m mckinsey_pptx.cli specs.json -o deck.pptx --section-marker "Strategy review"
```

### 데모 실행

```bash
python -m examples.demo          # 영문 15 슬라이드
python -m examples.demo_korean   # 한국어 21 슬라이드
```

### 테마 커스터마이즈

```python
from dataclasses import replace
from mckinsey_pptx import DEFAULT_THEME, PresentationBuilder

KO_THEME = replace(
    DEFAULT_THEME,
    typography=replace(DEFAULT_THEME.typography, family="Apple SD Gothic Neo"),
    copyright_text="ⓒ 2026 AX Labs",
)
b = PresentationBuilder(theme=KO_THEME)
```

### 전체 프로젝트 구조 (실제 파일)

```
axlabs-mckinsey-pptx/
├── .claude-plugin/
│   ├── marketplace.json                   # AX Labs 마켓플레이스 엔트리
│   └── plugin.json                        # 플러그인 매니페스트
├── agents/
│   └── mckinsey-slide-agent.md            # 서브에이전트 정의
├── commands/
│   └── mckinsey-deck.md                   # /mckinsey-deck 슬래시 커맨드
├── mckinsey_pptx/                         # Python 엔진
│   ├── __init__.py
│   ├── theme.py                           # 팔레트·폰트·레이아웃 토큰
│   ├── base.py                            # 공통 도형·텍스트 헬퍼 + 슬라이드 크롬
│   ├── builder.py                         # PresentationBuilder + add_specs 추론
│   ├── cli.py                             # 커맨드라인 인터페이스
│   ├── agent/
│   │   └── CATALOG.md                     # 서브에이전트용 템플릿 플레이북
│   └── slides/                            # 40개 템플릿 (카테고리별 파일)
│       ├── __init__.py
│       ├── assessment_table.py
│       ├── bubble_chart.py
│       ├── column_chart.py
│       ├── comparison_slides.py
│       ├── executive_summary.py
│       ├── extra_charts.py
│       ├── org_charts.py
│       ├── process_extras.py
│       ├── structure_slides.py
│       ├── summary_slide.py
│       ├── timeline_slides.py
│       └── trends_slides.py
├── examples/
│   ├── __init__.py
│   ├── demo.py                            # 영문 15 슬라이드 데모
│   └── demo_korean.py                     # 한국어 21 슬라이드 데모
├── LICENSE                                # MIT © 2026 AX Labs
├── README.md
└── requirements.txt                       # python-pptx 등 의존성
```

</details>

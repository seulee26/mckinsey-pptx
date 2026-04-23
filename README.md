# mckinsey-pptx

> **A Claude Code plugin by [AX Labs](mailto:seungpillee94@gmail.com).**
> McKinsey-style PPTX generator — 40 production-ready slide templates that
> faithfully reproduce the source design (deep navy / bright blue palette,
> bold-and-rule title, dashed section marker, bottom rule with source + page
> number). Built on `python-pptx` — every shape is drawn natively so the look
> matches the references.

Ships with a **Claude Code subagent** (`mckinsey-slide-agent`) that picks the
right template for each slide, explains its choice, and builds a real `.pptx`
from a one-paragraph brief. Also exposes a `/mckinsey-deck` slash command.

![Author](https://img.shields.io/badge/author-AX%20Labs-0b1f3a)
![License](https://img.shields.io/badge/license-MIT-blue)
![Templates](https://img.shields.io/badge/templates-40-brightgreen)
![Platform](https://img.shields.io/badge/platform-Claude%20Code-6b4bff)

---

## Install as a Claude Code plugin

Inside Claude Code:

```
/plugin marketplace add seulee26/mckinsey-pptx
/plugin install mckinsey-pptx@axlabs
```

Then install the Python dependencies once on your machine:

```bash
pip install -r "${CLAUDE_PLUGIN_ROOT:-.}/requirements.txt"

# optional — for PNG previews of generated decks
brew install --cask libreoffice   # soffice
brew install poppler              # pdftoppm
```

> The plugin ships the Python source (`mckinsey_pptx/`) alongside the agent.
> The agent automatically adds the plugin root to `sys.path` in every build
> script it generates, so you don't have to install the package globally.

### Verify

In a Claude Code session:

```
/agents
```

You should see `mckinsey-slide-agent` listed. Then try:

```
/mckinsey-deck Q4 사업 리뷰 데크. 매출 1,200억(전년 1,050억), KPI 지연 2건, 진출 영역 3개 검토 중.
```

---

## Two ways to use

### 1. Claude Code subagent (recommended) 🤖

From any Claude Code session with the plugin installed:

```
> 분기 사업 리뷰 데크 만들어줘. 매출은 1,200억(전년 1,050억),
  KPI 지연 2건, 진출 영역 3개 검토 중.
```

The agent will:

1. Read `${CLAUDE_PLUGIN_ROOT}/mckinsey_pptx/agent/CATALOG.md` (the template playbook)
2. Plan a 5–10 slide deck arc
3. For each slide, name the template and **explain why** it picked that one
   over nearby templates
4. Generate a Python build script under `output/` in your current working
   directory and run it
5. Render PNG previews and report back with paths and rationales

Trigger phrases the agent recognizes (English + Korean):
- "make a deck / presentation / slides / PPTX / PowerPoint"
- "build a McKinsey-style report / strategy deck"
- "맥킨지 슬라이드 / 데크 / 보고서 만들어줘"
- "전략 보고서 PPT", "사업 리뷰 데크"

You can also invoke explicitly with the slash command:

```
/mckinsey-deck <your brief>
```

### 2. Direct Python API

```python
from mckinsey_pptx import PresentationBuilder

b = PresentationBuilder(default_section_marker="Q4 review")

b.add("dark_navy_summary",
      body="[Bottom line]: K-battery's next 5 years decide global leadership.")

b.add("executive_summary_takeaways",
      sections=[
          {"takeaway": "Market grew 22% YoY",
           "bullets": ["NA share rising", "Europe stagnating"]},
      ])

b.add("column_historic_forecast",
      categories=[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026],
      values=[1035, 1108, 1153, 1148, 1206, 1265, 1381, 1430, 1535],
      forecast_from_index=5,
      historic_growth="3%", forecast_growth="6%")

b.save("output/deck.pptx")
```

For Korean / mixed content, swap the typography family to `Apple SD Gothic Neo`
(see `examples/demo_korean.py`).

---

## Slide template catalog (40 total)

The full catalog with **when to use / when NOT to use / required inputs / examples**
lives in [`mckinsey_pptx/agent/CATALOG.md`](mckinsey_pptx/agent/CATALOG.md).

### Executive summary & impact

| key | what |
|---|---|
| `executive_summary_paragraph` | Title + 2–4 paragraph blocks |
| `executive_summary_takeaways` | Bold takeaway → bullets, optional final conclusion |
| `dark_navy_summary` | Full-bleed deep navy single impact statement |

### Status & assessment

| key | what |
|---|---|
| `assessment_table` | Category × KPI table with traffic-light dots |

### Charts

| key | what |
|---|---|
| `column_comparison` | Sorted bars + focus highlight + right pane |
| `column_simple_growth` | Time series with one growth arrow |
| `column_split_growth` | Time series with two-phase growth |
| `column_historic_forecast` | Actuals (navy) + forecast (light blue) bars |
| `bubble_chart` | Full-width XY scatter with group colors |
| `bubble_chart_takeaways` | Same + right-side bullets |

### Matrices

| key | what |
|---|---|
| `growth_share` (`bcg_matrix`) | 2×2 BCG quadrants with bubbles |
| `prioritization_matrix` | 3×3 Time × Impact grid, status colors |

### Trends / areas

| key | what |
|---|---|
| `three_trends_icons` | 3 rows: circular icon + label + bullets |
| `three_trends_table` | 3 rows: name pill + description + examples |
| `three_trends_numbered` | 3 rows: number + blue label pill + bullets |
| `five_key_areas` | 5 numbered rows with arrow + description |
| `overview_areas` | 5–7 vertical area cards with letter badges |

### Org & hierarchy

| key | what |
|---|---|
| `issue_tree` | Root issue → main → secondary → underlying drivers |
| `org_chart` | CEO → N heads → reports per head |
| `project_team_circles` | One leader + N labeled teammate circles |
| `team_chart` | Function columns × role tiles (filled/outline) |

### Roadmap & process

| key | what |
|---|---|
| `phases_chevron_3` | 3 chevron-arrow phases with deliverables/people |
| `phases_table_4` | 4 columns: description + activities + outcomes |
| `waves_timeline_4` | 4 waves on horizontal arrow with markers |
| `gantt_timeline` | Weekly Gantt with workstream rows + milestones |
| `process_activities` | 3-row table: Activities / Mgmt / Deliverables |
| `process_flow_horizontal` | 4–6 numbered chevron tiles with descriptions |
| `funnel` | Top-down funnel (TAM/SAM/SOM, marketing funnel) |

### Structural (deck scaffolding)

| key | what |
|---|---|
| `cover_slide` | Title page with client + date + accent stripe |
| `section_divider` | Chapter divider (big number + chapter name) |
| `agenda` | Numbered table-of-contents with active highlight |
| `stat_hero` | One huge number + label + context |
| `quote_slide` | Big pull-quote with attribution |

### Comparison

| key | what |
|---|---|
| `comparison_table` | Options × criteria with Harvey-ball ratings |
| `pros_cons` | Two-column ✓ green / ✗ red analysis |
| `two_column_compare` | Before/After or As-is/To-be with arrow |

### Additional charts

| key | what |
|---|---|
| `stacked_column_chart` | Composition over time/category |
| `grouped_column_chart` | Side-by-side category comparison |
| `line_chart` | 1–4 series time-series with markers |
| `kpi_dashboard` | 4–8 KPI tiles with deltas (▲▼▬) |

---

## Running the demos (without the plugin)

Clone the repo and install requirements, then:

```bash
# English demo (15 slides)
python -m examples.demo

# Korean demo (21 slides — K-battery global strategy)
python -m examples.demo_korean
```

Outputs land in `output/`.

---

## Adaptive direct API

If you don't want to use the subagent but still want template inference,
`PresentationBuilder.add_specs()` accepts a list of dicts and infers the
template from each dict's shape:

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

You can override with `{"type": "<template_id>", ...}`.

---

## Customizing the theme

Edit `mckinsey_pptx/theme.py` or pass a custom `Theme(...)` to
`PresentationBuilder`:

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

## Project layout

```
mckinsey-pptx/
├── .claude-plugin/
│   ├── plugin.json              # Claude Code plugin manifest (name, version, author)
│   └── marketplace.json         # AX Labs marketplace entry (single-plugin)
├── agents/
│   └── mckinsey-slide-agent.md  # subagent definition (auto-loaded by plugin)
├── commands/
│   └── mckinsey-deck.md         # /mckinsey-deck slash command
├── mckinsey_pptx/               # Python implementation
│   ├── theme.py                 # palette / fonts / layout tokens
│   ├── base.py                  # shared shape/text helpers + slide chrome
│   ├── builder.py               # PresentationBuilder + adaptive routing
│   ├── cli.py
│   ├── slides/                  # 40 templates across modules
│   └── agent/
│       └── CATALOG.md           # template playbook the subagent reads
├── examples/
│   ├── demo.py                  # English 15-slide demo
│   └── demo_korean.py           # Korean 21-slide demo
├── LICENSE                      # MIT — © 2026 AX Labs (이승필)
├── README.md
└── requirements.txt
```

---

## About the author

Built and maintained by **AX Labs — 이승필 (Seungpil Lee)**.

AX Labs helps teams adopt AI-native consulting workflows (AX = *AI
eXperience / AI Transformation*). This plugin is an open-source slice of the
toolkit AX Labs uses internally to draft consulting decks in minutes instead
of days.

### Enterprise AX inquiries

If your organization wants:

- **Private slide templates** matching your firm's visual identity
- **Domain-tuned subagents** (industry playbooks, internal frameworks,
  proprietary data sources)
- **End-to-end AX transformation** (audit → pilot → deploy across functions)
- **On-prem / VPC deployment** of the agent pipeline
- **Training & enablement** for partner/consulting teams

→ Reach out: **seungpillee94@gmail.com** (subject: `[AX inquiry] <company>`)

Open-source contributions, bug reports, and template requests are welcome
via GitHub issues.

---

## License

[MIT](./LICENSE) © 2026 AX Labs — 이승필 (Seungpil Lee)

The MIT license covers the open-source plugin shipped here. Custom
extensions, private templates, and enterprise deployments operate under
separate commercial terms — contact AX Labs.

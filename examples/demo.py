"""Build one example deck per slide template.

Run:  python -m examples.demo  -> writes output/demo_full.pptx
"""
from __future__ import annotations
from pathlib import Path

from mckinsey_pptx import PresentationBuilder


SAMPLE_PARAGRAPHS = [
    "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod "
    "tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim "
    "veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea "
    "commodo consequat. Duis aute irure dolor in reprehenderit in voluptate "
    "velit esse cillum dolore eu fugiat nulla pariatur.]"
] * 4


SAMPLE_SECTIONS = [
    {"takeaway": "[Key takeaway/conclusion in bold text]",
     "bullets": [
         "[Supporting arguments for key takeaway/conclusion in bullet points]",
         "[Supporting arguments for key takeaway/conclusion in bullet points]",
         "[Supporting arguments for key takeaway/conclusion in bullet points]",
     ]} for _ in range(3)
]


def build(output_path: str = "output/demo_full.pptx") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    b = PresentationBuilder(default_section_marker="Section marker")

    # Executive summary - paragraph
    b.add("executive_summary_paragraph",
          title="Executive summary",
          subtitle="[Insert own text description key takeaway.]",
          paragraphs=SAMPLE_PARAGRAPHS)

    # Executive summary - takeaways
    b.add("executive_summary_takeaways",
          title="Executive summary",
          sections=SAMPLE_SECTIONS,
          final_conclusion="[Final conclusion/recommendation in bold text]")

    # Assessment table
    b.add("assessment_table",
          categories=[
              {"name": "[Area/BU 1]",
               "rows": [
                   {"kpi": "[insert dimension]", "target": "xx", "actual": "xx",
                    "status_label": "[insert status]", "status": "green"},
                   {"kpi": "[insert dimension]", "target": "xx", "actual": "xx",
                    "status_label": "[insert status]", "status": "green"},
                   {"kpi": "[insert dimension]", "target": "xx", "actual": "xx",
                    "status_label": "[insert status]", "status": "red"},
                   {"kpi": "[insert dimension]", "target": "xx", "actual": "xx",
                    "status_label": "[insert status]", "status": "green"},
               ]},
              {"name": "[Area/BU 2]",
               "rows": [
                   {"kpi": "[insert dimension]", "target": "xx", "actual": "xx",
                    "status_label": "[insert status]", "status": "green"},
                   {"kpi": "[insert dimension]", "target": "xx", "actual": "xx",
                    "status_label": "[insert status]", "status": "amber"},
                   {"kpi": "[insert dimension]", "target": "xx", "actual": "xx",
                    "status_label": "[insert status]", "status": "amber"},
               ]},
          ])

    # Bubble chart
    b.add("bubble_chart",
          bubbles=[
              {"label": "Product 1", "x": 150, "y": 1850, "size": 1,
               "group": "blue_light"},
              {"label": "Product 2", "x": 280, "y": 2200, "size": 0.6,
               "group": "blue_royal"},
              {"label": "Product 3", "x": 200, "y": 1200, "size": 0.4,
               "group": "blue_royal"},
              {"label": "Product 4", "x": 280, "y": 720, "size": 4,
               "group": "blue_royal"},
              {"label": "Product 5", "x": 320, "y": 1500, "size": 3,
               "group": "blue_dark"},
              {"label": "Product 6", "x": 470, "y": 1400, "size": 0.7,
               "group": "blue_dark"},
              {"label": "Product 7", "x": 560, "y": 2300, "size": 3,
               "group": "blue_light"},
              {"label": "Product 8", "x": 700, "y": 2300, "size": 0.7,
               "group": "blue_dark"},
              {"label": "Product 9", "x": 440, "y": 1500, "size": 0.5,
               "group": "blue_royal"},
              {"label": "Product 10", "x": 240, "y": 1380, "size": 0.7,
               "group": "blue_light"},
          ])

    # Bubble chart with takeaways
    b.add("bubble_chart_takeaways",
          bubbles=[
              {"label": "Product 1", "x": 230, "y": 1700, "size": 1.5,
               "group": "blue_light"},
              {"label": "Product 2", "x": 360, "y": 2050, "size": 0.4,
               "group": "blue_dark"},
              {"label": "Product 3", "x": 240, "y": 1180, "size": 0.5,
               "group": "blue_dark"},
              {"label": "Product 4", "x": 280, "y": 720, "size": 4,
               "group": "mid_blue"},
              {"label": "Product 5", "x": 330, "y": 1500, "size": 3,
               "group": "blue_light"},
              {"label": "Product 6", "x": 460, "y": 1400, "size": 0.6,
               "group": "blue_light"},
              {"label": "Product 7", "x": 580, "y": 2200, "size": 3,
               "group": "blue_dark"},
              {"label": "Product 8", "x": 700, "y": 2200, "size": 0.5,
               "group": "blue_light"},
              {"label": "Product 9", "x": 460, "y": 1620, "size": 0.5,
               "group": "blue_dark"},
              {"label": "Product 10", "x": 230, "y": 1380, "size": 0.5,
               "group": "blue_dark"},
          ],
          takeaways=[
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "…",
          ])

    # Growth-share matrix
    b.add("growth_share",
          bus=[
              {"name": "[BU 1]", "x": 12, "y": 37, "size": 4},
              {"name": "[BU 2]", "x": 30, "y": 21, "size": 1.6},
              {"name": "[BU 3]", "x": 32, "y": 18, "size": 1.4},
              {"name": "[BU 5]", "x": 50, "y": 11, "size": 2},
              {"name": "[BU 6]", "x": 58, "y": 8, "size": 1.6},
              {"name": "[BU 7]", "x": 70, "y": 19, "size": 1.6},
              {"name": "[BU 8]", "x": 70, "y": 45, "size": 2.6},
          ])

    # Prioritization matrix
    b.add("prioritization_matrix",
          items=[
              {"name": "[Insert name]", "x_band": 1, "y_band": 0, "ox": 0.25,
               "oy": 0.5, "status": "green"},
              {"name": "[Insert name]", "x_band": 1, "y_band": 0, "ox": 0.7,
               "oy": 0.5, "status": "green"},
              {"name": "[Insert name]", "x_band": 2, "y_band": 0, "ox": 0.65,
               "oy": 0.6, "status": "green"},
              {"name": "[Insert name]", "x_band": 2, "y_band": 0, "ox": 0.20,
               "oy": 0.85, "status": "amber"},
              {"name": "[Insert name]", "x_band": 2, "y_band": 1, "ox": 0.55,
               "oy": 0.40, "status": "red"},
              {"name": "[Insert name]", "x_band": 1, "y_band": 1, "ox": 0.3,
               "oy": 0.5, "status": "amber"},
              {"name": "[Insert name]", "x_band": 1, "y_band": 1, "ox": 0.85,
               "oy": 0.85, "status": "green"},
              {"name": "[Insert name]", "x_band": 0, "y_band": 2, "ox": 0.55,
               "oy": 0.30, "status": "amber"},
              {"name": "[Insert name]", "x_band": 2, "y_band": 2, "ox": 0.75,
               "oy": 0.40, "status": "amber"},
          ])

    # Column chart variants
    b.add("column_comparison",
          categories=["[Segment/group]", "[Focus]", "[Segment/group]",
                      "[Segment/group]", "[Segment/group]", "[Segment/group]",
                      "[Segment/group]", "[Segment/group]", "[Segment/group]",
                      "[Segment/group]"],
          values=[670, 623, 580, 564, 514, 421, 376, 360, 342, 301],
          focus_index=1,
          takeaways=[
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "…",
          ])

    b.add("column_simple_growth",
          categories=[2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022],
          values=[1035, 1108, 1153, 1148, 1206, 1265, 1381, 1430, 1535],
          growth_pct="xx%",
          takeaways=[
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "…",
          ])

    b.add("column_split_growth",
          categories=[2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022],
          values=[1035, 1108, 1153, 1148, 1206, 1265, 1381, 1430, 1535],
          split_index=4,
          growth_pct_first="xx%", growth_pct_second="xx%",
          takeaways=[
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "…",
          ])

    b.add("column_historic_forecast",
          categories=[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026],
          values=[1035, 1108, 1153, 1148, 1206, 1265, 1381, 1430, 1535],
          forecast_from_index=5,
          historic_growth="xx%", forecast_growth="xx%",
          takeaways=[
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "[Short description and most important takeaways]",
              "…",
          ])

    # Trends / areas
    b.add("three_trends_icons",
          subtitle="[Insert subtitle]",
          trends=[
              {"label": "Trend/area 1", "icon": "♛",
               "bullets": [
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
               ]},
              {"label": "Trend/area 2", "icon": "$",
               "bullets": [
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
               ]},
              {"label": "Trend/area 3", "icon": "📣",
               "bullets": [
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
               ]},
          ])

    b.add("three_trends_table",
          trends=[
              {"name": f"Trend {i}",
               "description": [
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
               ],
               "examples": ["[Select examples]"]} for i in (1, 2, 3)
          ])

    b.add("three_trends_numbered",
          subtitle="[Insert subtitle]",
          trends=[
              {"label": f"Trend {i}",
               "bullets": [
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
                   "[Short description of trend in 3-4 bullet points]",
               ]} for i in (1, 2, 3)
          ])

    b.add("five_key_areas",
          subtitle="[Insert subtitle]",
          areas=[
              {"name": f"Area {i}",
               "description": "[Insert description, ideally 2-4 lines]"}
              for i in (1, 2, 3, 4, 5)
          ])

    b.save(output_path)
    return output_path


if __name__ == "__main__":
    out = build()
    print(f"wrote {out}")

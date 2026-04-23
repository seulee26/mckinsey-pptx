"""한국어 샘플 데크 — '국내 EV 배터리 사업의 글로벌 진출 전략 리뷰'.

McKinsey 스타일 그대로 유지하면서 한국어 폰트로 렌더링한다.
"""
from __future__ import annotations
from pathlib import Path
from dataclasses import replace

from mckinsey_pptx import PresentationBuilder, DEFAULT_THEME
from mckinsey_pptx.theme import Typography


# 한글 친화 테마 — macOS 기본 한글 폰트로 교체
KO_THEME = replace(
    DEFAULT_THEME,
    typography=replace(DEFAULT_THEME.typography, family="Apple SD Gothic Neo"),
    copyright_text="ⓒ 2026 mckinsey-AX",
)


def build(output_path: str = "output/demo_korean.pptx") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    b = PresentationBuilder(theme=KO_THEME, default_section_marker="진출 전략")

    # 1. Executive summary (key takeaway)
    b.add(
        "executive_summary_takeaways",
        title="Executive summary",
        sections=[
            {"takeaway": "글로벌 EV 배터리 시장은 2030년까지 연 22% 성장이 예상되며, 한국 3사의 합산 점유율은 23%까지 하락 위험",
             "bullets": [
                 "중국 CATL/BYD의 LFP 기반 가격 공세가 본격화되며 NCM 중심의 한국 3사 입지 축소",
                 "북미·유럽 OEM의 자체 배터리 생산 내재화(예: GM Ultium, Stellantis-Samsung JV) 가속",
                 "원소재(리튬·니켈) 가격 변동성이 마진을 5–7%p 압박",
             ]},
            {"takeaway": "원가·기술 양 측면에서 LFP 라인업 확장 + 차세대 전고체(SSB) R&D 가속이 필수",
             "bullets": [
                 "단기(2026–27): LFP 양산 라인을 폴란드/미시간 두 거점에 우선 증설",
                 "중기(2028–29): 46파이 원형셀 및 건식 전극 양산화 — 원가 12% 절감 목표",
                 "장기(2030+): 전고체 시제품을 2027년까지 OEM 인증 완료",
             ]},
            {"takeaway": "현지 파트너십과 IRA·CRMA 인센티브 활용을 결합한 거점 전략이 글로벌 입지 회복의 핵심",
             "bullets": [
                 "북미: GM·Ford 외 신규 OEM 1곳과 JV 추가 검토",
                 "유럽: 헝가리·폴란드 거점 가동률을 90% 이상으로 회복",
                 "ASEAN: 인도네시아 니켈-제련 + 셀 통합 모델 검증",
             ]},
        ],
        final_conclusion="결론: LFP 가속 + 전고체 R&D + 거점 재배치를 동시 추진하여 2030년 점유율 30% 회복을 목표로 한다.",
    )

    # 2. Five key areas — 진출 5대 우선 영역
    b.add(
        "five_key_areas",
        title="2030 K-배터리 5대 전략 영역",
        subtitle="단기 12개월 내 의사결정이 필요한 우선순위",
        areas=[
            {"name": "제품 포트폴리오",
             "description": "NCM 고에너지 라인업과 LFP 가성비 라인업을 동시 운영하여 OEM별 맞춤 대응"},
            {"name": "원소재 공급망",
             "description": "리튬·니켈 장기 계약 + 폐배터리 재활용 비중을 2027년까지 25%로 확대"},
            {"name": "거점 전략",
             "description": "북미 IRA 대응 거점 신설, 유럽 CRMA 대응 가동률 회복, ASEAN 통합 거점 검증"},
            {"name": "차세대 R&D",
             "description": "전고체·46파이 원형셀·건식 전극 3대 기술을 2030년 양산 가능 수준으로 확보"},
            {"name": "조직·역량",
             "description": "현지 인력 확보, 파일럿 라인 운영 인력 재교육, ESG·안전 거버넌스 강화"},
        ],
    )

    # 3. Three key trends (numbered) — 시장 3대 트렌드
    b.add(
        "three_trends_numbered",
        title="글로벌 EV 배터리 시장 3대 트렌드",
        subtitle="향후 5년간 산업 구조를 결정할 메가 트렌드",
        trends=[
            {"label": "LFP 대중화",
             "bullets": [
                 "Tesla·Ford·VW가 엔트리 모델에 LFP 채택을 빠르게 확대",
                 "kWh당 원가가 NCM 대비 약 30% 낮아 가격 경쟁의 핵심",
                 "한국 3사도 2025년부터 LFP 신규 라인 가동 본격화",
             ]},
            {"label": "공급망 재편",
             "bullets": [
                 "IRA·CRMA로 북미·유럽 현지화 비중이 의무화 수준에 도달",
                 "중국 의존도를 50% 미만으로 축소하기 위한 대체 공급원 확보 경쟁",
                 "폐배터리 재활용 시장이 2030년 약 300억 달러 규모로 확대",
             ]},
            {"label": "차세대 셀 기술",
             "bullets": [
                 "전고체(SSB)는 2027–2028년 시제품, 2030년 양산 진입 전망",
                 "46파이 원형셀이 차세대 표준으로 부상",
                 "건식 전극·실리콘 음극 등 공정 혁신이 원가 우위의 핵심",
             ]},
        ],
    )

    # 4. Three trends with icons
    b.add(
        "three_trends_icons",
        title="고객(OEM) 관점의 핵심 의사결정 요인",
        subtitle="신규 셀 공급사 선정 시 OEM이 가장 중시하는 3가지",
        trends=[
            {"label": "원가 경쟁력", "icon": "$",
             "bullets": [
                 "kWh당 단가와 BOM 구성의 투명성",
                 "원소재 가격 변동에 대한 헤지 능력",
                 "장기 공급 안정성을 위한 가격 인덱싱 구조",
             ]},
            {"label": "기술·품질", "icon": "♛",
             "bullets": [
                 "에너지 밀도·충전 속도·수명 등 핵심 사양",
                 "양산 수율 및 품질 관리 데이터의 신뢰성",
                 "차세대 셀 로드맵의 현실성",
             ]},
            {"label": "현지 대응력", "icon": "★",
             "bullets": [
                 "OEM 공장 인근 거점의 운영 능력",
                 "IRA·CRMA 등 규제 적격성 충족",
                 "ESG·안전 사고 대응 거버넌스",
             ]},
        ],
    )

    # 5. Column chart with historic + forecast — 시장 규모
    b.add(
        "column_historic_forecast",
        title="글로벌 EV 배터리 시장 규모 — 2018–2030 전망",
        data_label="시장 규모", data_unit="십억 달러",
        categories=[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025,
                    2026, 2027, 2028, 2029, 2030],
        values=[28, 36, 49, 78, 112, 145, 188, 235,
                292, 358, 432, 514, 605],
        forecast_from_index=5,
        historic_growth="+39%",
        forecast_growth="+22%",
        takeaways=[
            "2023년 이후 연평균 약 22%의 견조한 성장 지속 전망",
            "북미·유럽 비중이 2030년까지 합산 55%로 확대",
            "한국 3사 점유율은 23–30% 구간에서 정책·R&D 성과에 좌우",
            "…",
        ],
    )

    # 6. Column comparison — 2025년 시장 점유율
    b.add(
        "column_comparison",
        title="2025년 EV 배터리 제조사별 점유율",
        data_label="점유율", data_unit="%",
        categories=["CATL", "BYD", "LG에너지솔루션", "파나소닉",
                    "SK On", "삼성SDI", "CALB", "Gotion", "EVE", "기타"],
        values=[37, 17, 13, 8, 7, 5, 4, 3, 3, 3],
        focus_index=2,
        takeaways=[
            "CATL 단일 점유율이 한국 3사 합산(25%)을 상회",
            "한국 3사는 LG > SK > 삼성 순으로 차이 유지",
            "중국 2티어(CALB, Gotion, EVE)의 빠른 추격에 주의",
            "…",
        ],
    )

    # 7. Bubble chart with takeaways — 경쟁사 포지셔닝
    b.add(
        "bubble_chart_takeaways",
        title="EV 배터리 경쟁사 포지셔닝 — 원가 vs. 에너지 밀도",
        x_label="kWh당 원가", x_unit="달러/kWh",
        y_label="에너지 밀도", y_unit="Wh/kg",
        x_max=200, y_max=350,
        bubbles=[
            {"label": "CATL",  "x": 88,  "y": 260, "size": 4,   "group": "blue_dark"},
            {"label": "BYD",   "x": 82,  "y": 220, "size": 2.6, "group": "blue_dark"},
            {"label": "LGES",  "x": 105, "y": 295, "size": 2.4, "group": "blue_light"},
            {"label": "SK On", "x": 110, "y": 285, "size": 1.4, "group": "blue_light"},
            {"label": "삼성SDI", "x": 115, "y": 310, "size": 1.2, "group": "blue_light"},
            {"label": "Panasonic", "x": 118, "y": 305, "size": 1.6, "group": "mid_blue"},
            {"label": "CALB",  "x": 95,  "y": 230, "size": 0.8, "group": "blue_dark"},
            {"label": "EVE",   "x": 92,  "y": 215, "size": 0.6, "group": "blue_dark"},
        ],
        groups=(("blue_dark", "중국계"),
                ("blue_light", "한국 3사"),
                ("mid_blue", "일본/기타")),
        takeaways=[
            "한국 3사는 에너지 밀도(품질)에서 우위, 원가에서 열세",
            "CATL은 원가·물량 양쪽에서 압도적 — 신규 진입자에게 큰 위협",
            "BYD는 LFP 비중 확대로 원가 우위를 유지하며 추격 중",
            "차세대 셀(46파이·전고체)이 격차 전환의 핵심 변수",
        ],
    )

    # 8. Growth-share matrix — 사업부별 BCG
    b.add(
        "growth_share",
        title="사업부별 성장-점유율 매트릭스 (BCG)",
        bus=[
            {"name": "EV 셀(NCM)", "x": 28, "y": 38, "size": 4.5},
            {"name": "EV 셀(LFP)", "x": 12, "y": 46, "size": 3.0},
            {"name": "ESS",        "x": 18, "y": 34, "size": 1.8},
            {"name": "전고체 R&D", "x": 6,  "y": 48, "size": 1.4},
            {"name": "원형셀(46파이)", "x": 15, "y": 40, "size": 1.6},
            {"name": "소형 IT 셀", "x": 62, "y": 8,  "size": 1.2},
            {"name": "전동공구",   "x": 55, "y": 6,  "size": 0.8},
        ],
    )

    # 9. Prioritization matrix — 진출 옵션 우선순위
    b.add(
        "prioritization_matrix",
        title="진출 옵션 우선순위 매트릭스",
        items=[
            {"name": "북미 LFP 증설",  "x_band": 2, "y_band": 0, "ox": 0.7, "oy": 0.5, "status": "green"},
            {"name": "유럽 가동률 회복","x_band": 2, "y_band": 0, "ox": 0.25,"oy": 0.6, "status": "green"},
            {"name": "전고체 R&D 가속", "x_band": 2, "y_band": 2, "ox": 0.5, "oy": 0.4, "status": "amber"},
            {"name": "46파이 원형셀",  "x_band": 1, "y_band": 1, "ox": 0.6, "oy": 0.5, "status": "green"},
            {"name": "ASEAN 통합거점", "x_band": 1, "y_band": 2, "ox": 0.3, "oy": 0.5, "status": "amber"},
            {"name": "재활용 합작",     "x_band": 1, "y_band": 1, "ox": 0.2, "oy": 0.7, "status": "amber"},
            {"name": "원소재 장기계약", "x_band": 0, "y_band": 0, "ox": 0.6, "oy": 0.5, "status": "green"},
            {"name": "소형 IT 정리",    "x_band": 0, "y_band": 2, "ox": 0.5, "oy": 0.6, "status": "red"},
        ],
    )

    # 10. Assessment table — 진출 옵션 점검
    b.add(
        "assessment_table",
        title="2026 핵심 KPI 점검 — 사업부별 현황",
        categories=[
            {"name": "EV 사업부",
             "rows": [
                 {"kpi": "북미 가동률", "target": "85%", "actual": "78%",
                  "status_label": "근접", "status": "amber"},
                 {"kpi": "유럽 가동률", "target": "90%", "actual": "92%",
                  "status_label": "달성", "status": "green"},
                 {"kpi": "kWh 원가 절감", "target": "-10%", "actual": "-6%",
                  "status_label": "지연", "status": "red"},
                 {"kpi": "신규 OEM 수주", "target": "3건", "actual": "3건",
                  "status_label": "달성", "status": "green"},
             ]},
            {"name": "ESS 사업부",
             "rows": [
                 {"kpi": "북미 수주",      "target": "5GWh", "actual": "5.2GWh",
                  "status_label": "달성", "status": "green"},
                 {"kpi": "프로젝트 마진",  "target": "12%",  "actual": "9%",
                  "status_label": "지연", "status": "amber"},
                 {"kpi": "안전 사고",      "target": "0건",  "actual": "0건",
                  "status_label": "달성", "status": "green"},
             ]},
        ],
        source="2026년 1분기 사내 KPI 보고서",
        footnote="1. 가동률 = 실제 생산량 / 설계 능력",
    )

    # 11. Dark navy 임팩트 서머리 (장 구분용)
    b.add(
        "dark_navy_summary",
        body="[핵심 메시지]: 향후 5년이 K-배터리의 글로벌 우위를 가르는 결정적 구간이며, "
             "제품·공급망·거점·R&D·조직의 5대 영역을 동시에 가동해야 한다.",
        eyebrow="K-배터리 글로벌 진출 전략",
    )

    # 12. Issue tree — 핵심 이슈 분해
    b.add(
        "issue_tree",
        title="K-배터리 점유율 하락의 원인 분해",
        subtitle="2025년 23% 점유율 가정 하의 driver tree",
        root="K-배터리\n점유율 하락",
        main_drivers=[
            {"label": "원가 경쟁력 열위",
             "secondaries": [
                 {"label": "셀 원가 격차",
                  "underlying": ["LFP 라인업 부재", "원소재 장기계약 부족"]},
                 {"label": "공정 수율 부진",
                  "underlying": ["46파이 신라인 안정화 지연"]},
             ]},
            {"label": "포트폴리오 미스매치",
             "secondaries": [
                 {"label": "OEM 수요 변화",
                  "underlying": ["가성비 OEM 비중 확대", "전고체 전환 가속"]},
                 {"label": "신차 슬롯 미확보",
                  "underlying": ["북미 OEM JV 지연"]},
             ]},
        ],
    )

    # 13. Org chart — TF 조직도
    b.add(
        "org_chart",
        title="글로벌 진출 TF 조직 구성안",
        subtitle="CEO 직속 5개 본부 / 17개 팀",
        ceo="CEO 김OO",
        branches=[
            {"head": "전략기획본부",
             "reports": ["사업개발팀", "전략전담팀", "M&A팀"]},
            {"head": "북미 사업본부",
             "reports": ["IRA TF", "JV 운영팀", "공장 건설팀"]},
            {"head": "유럽 사업본부",
             "reports": ["CRMA TF", "OEM 영업팀"]},
            {"head": "R&D본부",
             "reports": ["전고체 R&D", "46파이 R&D", "건식 전극 R&D"]},
            {"head": "공급망본부",
             "reports": ["원소재 조달", "재활용 사업팀"]},
        ],
    )

    # 14. Project team circles — 핵심 인력 구성
    b.add(
        "project_team_circles",
        title="진출 TF 핵심 인력 구성",
        subtitle="리더 1명 + 5개 영역별 책임자",
        leader={"name": "프로젝트 리더",
                "description": "TF 총괄 / CEO 직보\n전사 의사결정 권한",
                "icon": "👥"},
        members=[
            {"name": "제품 책임",
             "description": "라인업·로드맵 의사결정", "icon": "🔧"},
            {"name": "공급망 책임",
             "description": "원소재·재활용 계약 통합", "icon": "📈"},
            {"name": "거점 책임",
             "description": "신규 공장 / JV 협상", "icon": "🚀"},
            {"name": "R&D 책임",
             "description": "차세대 기술 양산 검증", "icon": "☁"},
            {"name": "조직·HR",
             "description": "현지 채용·교육·거버넌스", "icon": "⚖"},
        ],
    )

    # 15. Team chart — Function별 인력 배치
    b.add(
        "team_chart",
        title="Function × Role 매트릭스",
        project_name="K-배터리\nTF",
        functions=[
            {"name": "Strategy",
             "description": "사업 의사결정·재무 모델",
             "roles": [{"name": "팀장", "kind": "filled"},
                       {"name": "전략 컨설턴트", "kind": "filled"},
                       {"name": "어소시에이트", "kind": "outline"}]},
            {"name": "북미",
             "description": "IRA 대응·JV 협상",
             "roles": [{"name": "본부장", "kind": "filled"},
                       {"name": "현지 PM", "kind": "filled"},
                       {"name": "법무 자문", "kind": "outline"}]},
            {"name": "유럽",
             "description": "CRMA 대응·OEM 영업",
             "roles": [{"name": "본부장", "kind": "filled"},
                       {"name": "현지 영업", "kind": "filled"}]},
            {"name": "R&D",
             "description": "차세대 셀·공정 개발",
             "roles": [{"name": "수석 연구원", "kind": "filled"},
                       {"name": "선임", "kind": "outline"},
                       {"name": "선임", "kind": "outline"}]},
            {"name": "공급망",
             "description": "원소재·재활용·물류",
             "roles": [{"name": "총괄", "kind": "filled"},
                       {"name": "구매", "kind": "outline"}]},
            {"name": "PMO",
             "description": "전체 실행 모니터링",
             "roles": [{"name": "PMO 리드", "kind": "filled"},
                       {"name": "분석가", "kind": "outline"}]},
        ],
    )

    # 16. Three phases chevron — 3단계 진행 계획
    b.add(
        "phases_chevron_3",
        title="진출 프로젝트는 3단계로 전개된다",
        phases=[
            {"label": "기회 발굴 및 가설 수립",
             "timeframe": "2026 Q2",
             "deliverables": ["북미·유럽 5대 OEM 인터뷰",
                              "시장 모델링 v1"],
             "people": ["전략실 4명", "외부 자문 2명"]},
            {"label": "상세 설계 및 MVP",
             "timeframe": "2026 Q3–Q4",
             "deliverables": ["라인업 합의안",
                              "JV term-sheet 초안"],
             "people": ["TF 12명", "법무·재무 협력"]},
            {"label": "실행 및 시장 진입",
             "timeframe": "2027 H1",
             "deliverables": ["북미 신규 라인 가동",
                              "OEM 수주 5건"],
             "people": ["TF 25명", "현지 채용"]},
        ],
    )

    # 17. Four phases (column-style)
    b.add(
        "phases_table_4",
        title="진출 프로젝트는 4개 phase로 진행된다",
        phases=[
            {"name": "Phase 1",
             "description": "현황 진단 및 시장·고객 deep dive를 통한 진출 영역 도출",
             "activities": ["북미 IRA 영향 분석", "유럽 CRMA 시뮬레이션",
                            "5대 OEM 의사결정 매핑"],
             "outcomes": ["시장 모델 v1", "후보 진출 영역 longlist"]},
            {"name": "Phase 2",
             "description": "후보 영역에 대한 사업성·운영성 정밀 평가",
             "activities": ["JV/M&A 타깃 스크리닝", "재무 모델링",
                            "현지 법무 검토"],
             "outcomes": ["진출 영역 shortlist (3개)", "투자 규모 추정치"]},
            {"name": "Phase 3",
             "description": "최종 의사결정 및 상세 실행 계획 수립",
             "activities": ["이사회 보고", "조직·인력 설계", "공급망 협상"],
             "outcomes": ["최종 진출안", "1기 실행 로드맵"]},
            {"name": "Phase 4",
             "description": "실행 및 모멘텀 확보",
             "activities": ["거점 신규 가동", "OEM 수주 클로징",
                            "거버넌스 정착"],
             "outcomes": ["MVP 라인 가동", "초도 매출 확보"]},
        ],
    )

    # 18. Four waves timeline
    b.add(
        "waves_timeline_4",
        title="진출은 4개 wave로 단계적으로 가동한다",
        subtitle="Wave별 책임 분리 및 결과물 명확화",
        waves=[
            {"name": "Wave 1", "headline": "기초 다지기",
             "timeframe": "2026 Q2",
             "activities": ["내부 정렬", "데이터 인프라 구축"],
             "deliverables": ["시장 모델", "내부 합의 메모"]},
            {"name": "Wave 2", "headline": "북미 거점 강화",
             "timeframe": "2026 Q3–Q4",
             "activities": ["IRA 대응 거점 신설", "OEM 5대 수주 협상"],
             "deliverables": ["JV 합의서", "수주 LOI 3건"]},
            {"name": "Wave 3", "headline": "유럽 회복",
             "timeframe": "2027 H1",
             "activities": ["가동률 회복", "CRMA 인증 통과"],
             "deliverables": ["가동률 92%", "CRMA 적합 통보"]},
            {"name": "Wave 4", "headline": "차세대 양산",
             "timeframe": "2028 H1",
             "activities": ["전고체 시양산", "46파이 양산 안정화"],
             "deliverables": ["GWh급 양산 라인"]},
        ],
    )

    # 19. Gantt timeline (주 단위)
    b.add(
        "gantt_timeline",
        title="향후 15주 상세 실행 계획",
        subtitle="2026년 9월~12월 (Week 37–51)",
        weeks=list(range(37, 52)),
        workstreams=[
            {"name": "01 시장·고객 분석",      "start_week": 37, "end_week": 41,
             "color": "blue_light"},
            {"name": "02 IRA 시나리오",        "start_week": 37, "end_week": 39,
             "color": "blue_light"},
            {"name": "03 OEM 인터뷰",          "start_week": 38, "end_week": 42,
             "color": "blue_light"},
            {"name": "01 JV 후보 스크리닝",    "start_week": 41, "end_week": 47,
             "color": "blue_dark"},
            {"name": "02 재무 모델링",         "start_week": 42, "end_week": 46,
             "color": "blue_dark"},
            {"name": "03 법무·세무 검토",      "start_week": 43, "end_week": 47,
             "color": "blue_dark"},
            {"name": "01 최종 의사결정 패키지", "start_week": 47, "end_week": 51,
             "color": "royal"},
        ],
        milestones=[
            {"week": 37, "label": "Kick-off"},
            {"week": 41, "label": "Stage-gate 1"},
            {"week": 47, "label": "Stage-gate 2"},
            {"week": 51, "label": "최종 보고"},
        ],
    )

    # 20. Overview of areas — 7개 진출 영역
    b.add(
        "overview_areas",
        title="진출 검토 7개 핵심 영역",
        subtitle="단기 의사결정 영역 vs. 중장기 투자 영역 구분",
        areas=[
            {"name": "[제품]",
             "bullets": ["LFP 라인업", "프리미엄 NCM", "ESS 셀"]},
            {"name": "[공급망]",
             "bullets": ["리튬·니켈 장기계약", "재활용 합작"]},
            {"name": "[북미]",
             "bullets": ["IRA 대응 거점", "OEM JV 확대"]},
            {"name": "[유럽]",
             "bullets": ["CRMA 적합 인증", "가동률 회복"]},
            {"name": "[R&D]",
             "bullets": ["전고체", "46파이", "건식 전극"]},
            {"name": "[조직]",
             "bullets": ["현지 채용", "ESG 거버넌스"]},
            {"name": "[디지털]",
             "bullets": ["품질 AI", "공급망 가시화"]},
        ],
        call_out="단기 12개월 결정 영역",
    )

    # 21. Process activities — Week별 활동/관리/산출물
    b.add(
        "process_activities",
        title="K-배터리 진출 검토는 12주에 걸쳐 진행된다",
        subtitle="High-level project plan and deliverables",
        steps=[
            {"name": "Week 1–4", "subtitle": "리서치 및 가설 수립",
             "activities": ["산업·경쟁 분석", "OEM 의사결정 매핑",
                            "1차 가설 도출"],
             "interaction": "Kick-off",
             "deliverable": "가설 개요서"},
            {"name": "Week 5–8", "subtitle": "분석 및 인터뷰",
             "activities": ["북미·유럽 OEM 인터뷰", "재무 모델링 v1",
                            "JV 후보 스크리닝"],
             "interaction": "1차 검토",
             "deliverable": "중간 보고서"},
            {"name": "Week 9–12", "subtitle": "결론 및 권고안",
             "activities": ["진출안 통합", "리스크·민감도 분석",
                            "이사회 보고 자료"],
             "interaction": "권고안 검토",
             "deliverable": "최종 보고서"},
        ],
    )

    b.save(output_path)
    return output_path


if __name__ == "__main__":
    out = build()
    print(f"wrote {out}")

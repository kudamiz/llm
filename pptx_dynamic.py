# 1. 템플릿 모드 데이터
{
    "type": "template",
    "layout_index": 1,
    "content_mapping": {"Title": "...", "Body": "..."}
}

# 2. 다이내믹 모드 데이터 (신규)
{
    "type": "dynamic",
    "layout_index": 9, # Dynamic_Base 레이아웃 번호
    "title": "시장 점유율 분석",
    "layout_plan": "Split_Left_Right", # 화면 분할 방식
    "components": [
        {"type": "chart", "position": "left", "data": ...},
        {"type": "text", "position": "right", "content": ...}
    ]
}

def planner_node(state: AgentState):
    # ... (기존 설정) ...
    
    system_prompt = """
    당신은 PPT 기획자입니다.
    
    [판단 로직]
    1. 사용자의 요청이 [템플릿 가이드]의 특정 레이아웃과 정확히 일치하면 -> **'template' 모드** 사용.
    2. 맞는 레이아웃이 없거나, 표/차트 등 복합 구성이 필요하면 -> **'dynamic' 모드** 사용.
    
    [Dynamic 모드 작성법]
    - layout_index: 'Dynamic_Base'의 인덱스를 사용.
    - components: 화면을 어떻게 채울지 정의 (chart, table, text, image).
    - layout_plan: 'Full', 'Split_Left_Right', 'Split_Top_Bottom' 중 택 1.

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

# --- [도구 함수들] ---
def draw_chart(slide, x, y, w, h, data):
    # data format: {"labels": ["A", "B"], "values": [10, 20]}
    chart_data = CategoryChartData()
    chart_data.categories = data.get('labels', [])
    chart_data.add_series('Series 1', data.get('values', []))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, chart_data)

def draw_table(slide, x, y, w, h, data_rows):
    rows = len(data_rows)
    cols = len(data_rows[0])
    graphic_frame = slide.shapes.add_table(rows, cols, x, y, w, h)
    table = graphic_frame.table
    for r in range(rows):
        for c in range(cols):
            table.cell(r, c).text = str(data_rows[r][c])

def draw_text(slide, x, y, w, h, text):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tb.text_frame.text = text
    # 스타일링 (Dynamic 모드는 폰트 지정 필요)
    for p in tb.text_frame.paragraphs:
        p.font.size = Pt(14)
        p.font.name = "AppleSDGothicNeo"

# --- [Renderer 메인 로직] ---
def renderer_node(state: AgentState):
    prs = Presentation(state["template_path"])
    
    # 캔버스 작업 영역 정의 (제목, 로고 제외한 빈 공간)
    canvas_x = Inches(0.5)
    canvas_y = Inches(1.5) # 제목 아래부터 시작
    canvas_w = Inches(9.0)
    canvas_h = Inches(5.0)

    for slide_plan in state["slide_data"]:
        # 1. 템플릿 모드 (기존 로직)
        if slide_plan["type"] == "template":
            # ... (기존 코드) ...
            pass
            
        # 2. 다이내믹 모드 (신규 로직)
        elif slide_plan["type"] == "dynamic":
            slide = prs.slides.add_slide(prs.slide_layouts[slide_plan["layout_index"]])
            
            # 제목은 Placeholder 사용 (일관성 유지)
            if slide.shapes.title:
                slide.shapes.title.text = slide_plan.get("title", "")
            
            # 레이아웃 계산 (간단한 그리드 시스템)
            plan = slide_plan["layout_plan"]
            comps = slide_plan["components"]
            
            # 영역 분할 로직
            regions = []
            if plan == "Split_Left_Right":
                regions = [
                    (canvas_x, canvas_y, canvas_w/2 - Inches(0.2), canvas_h), # 왼쪽
                    (canvas_x + canvas_w/2 + Inches(0.2), canvas_y, canvas_w/2 - Inches(0.2), canvas_h) # 오른쪽
                ]
            else: # Full
                regions = [(canvas_x, canvas_y, canvas_w, canvas_h)]
            
            # 컴포넌트 그리기
            for i, comp in enumerate(comps):
                if i >= len(regions): break # 영역보다 컴포넌트가 많으면 무시
                x, y, w, h = regions[i]
                
                if comp["type"] == "chart":
                    draw_chart(slide, x, y, w, h, comp["data"])
                elif comp["type"] == "table":
                    draw_table(slide, x, y, w, h, comp["data"])
                elif comp["type"] == "text":
                    draw_text(slide, x, y, w, h, comp["content"])
                elif comp["type"] == "image":
                     # 이미지 파일명으로 바이너리 찾아서 삽입 (이전 코드 활용)
                     pass

    prs.save(state["output_path"])

    [응답 형식 (JSON List)]
    [
      { "type": "template", "layout_index": 0, "content_mapping": {...} },
      { 
        "type": "dynamic", 
        "layout_index": 9, 
        "title": "매출 분석",
        "layout_plan": "Split_Left_Right",
        "components": [
           { "type": "table", "position": "left", "data": [["연도","매출"],["2024","100"]] },
           { "type": "text", "position": "right", "content": "매출이 급상승함" }
        ]
      }
    ]
    """
    # ... (LLM 호출 및 파싱 로직) ...


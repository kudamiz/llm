from pptx import Presentation

def get_template_guide(pptx_path):
    prs = Presentation(pptx_path)
    guide_text = "현재 사용 가능한 PPT 레이아웃 목록입니다:\n"
    
    # 모든 마스터 레이아웃을 순회
    for i, layout in enumerate(prs.slide_layouts):
        # 레이아웃 이름 (예: Comparison)
        layout_info = {
            "layout_index": i,
            "layout_name": layout.name,
            "placeholders": []
        }
        
        # 레이아웃 안의 구멍(Placeholder)들 이름 수집
        for shape in layout.placeholders:
            # PPT '선택 창'에서 지정한 이름을 그대로 가져옴
            p_info = f"{shape.name} (ID: {shape.placeholder_format.idx})"
            layout_info["placeholders"].append(p_info)
            
        guide_text += str(layout_info) + "\n"
        
    return guide_text

# 실행 결과 예시 (이 텍스트가 자동으로 생성됨)
# "{'layout_index': 1, 'layout_name': '2단비교', 'placeholders': ['Title (ID:0)', 'Body_Left (ID:1)', 'Body_Right (ID:2)']}"


# 시스템 프롬프트 템플릿
system_prompt = """
당신은 PPT 생성 전문가입니다. 
아래 제공된 [템플릿 가이드]를 보고, 사용자 입력에 가장 적합한 layout_index를 선택하고,
각 placeholder 이름에 맞는 내용을 JSON으로 생성하세요.

[템플릿 가이드]
{template_guide}  <-- 여기에 파이썬이 읽은 정보가 자동으로 들어감
"""

# 실행 시점
current_guide = get_template_guide("company_template_v2.pptx") # 파일만 바꾸면 됨
formatted_prompt = system_prompt.format(template_guide=current_guide)

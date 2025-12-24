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

from typing import List, Dict
from pydantic import BaseModel, Field
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate

# 1. LLM이 뱉어내야 할 최종 데이터 구조 정의 (Schema)
class SlideOutput(BaseModel):
    layout_index: int = Field(..., description="선택한 슬라이드 레이아웃의 인덱스 번호")
    # key: placeholder 이름, value: 들어갈 내용
    content_mapping: Dict[str, str] = Field(..., description="Placeholder 이름을 키(Key)로, 채울 내용을 값(Value)으로 하는 딕셔너리")
    reason: str = Field(..., description="이 레이아웃을 선택한 이유")

# 2. 에이전트 함수 정의
def generate_slide_json(user_input: str, template_guide: str):
    # 모델 설정 (JSON 모드 지원하는 모델 권장)
    llm = ChatOpenAI(model="gpt-4o", temperature=0)
    
    # 구조화된 출력을 하도록 설정
    structured_llm = llm.with_structured_output(SlideOutput)

    # 프롬프트 구성 (동적 템플릿 가이드 주입)
    system_prompt = """
    당신은 PPT 생성 전문가입니다. 
    사용자의 입력을 분석하고, 아래 [템플릿 가이드]를 참고하여 가장 적절한 레이아웃을 선택하세요.
    그리고 각 Placeholder의 'Name'에 맞춰 내용을 요약/배치하여 JSON으로 반환하세요.
    
    [템플릿 가이드]
    {guide}
    """
    
    prompt = ChatPromptTemplate.from_messages([
        ("system", system_prompt),
        ("human", "{input}")
    ])

    # 실행 체인
    chain = prompt | structured_llm
    
    # 결과 반환 (Pydantic 객체)
    return chain.invoke({"guide": template_guide, "input": user_input})

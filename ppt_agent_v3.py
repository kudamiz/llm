from pydantic import BaseModel, Field
from typing import Literal

class ReviewResult(BaseModel):
    status: Literal["PASS", "FAIL"] = Field(
        ..., description="검수 결과. 규칙을 모두 준수했으면 PASS, 위반사항이 있으면 FAIL"
    )
    feedback: str = Field(
        ..., description="FAIL일 경우, 구체적인 수정 지시사항. (PASS면 'Good' 입력)"
    )


def reviewer_node(state: AgentState):
    print("--- [Node: Reviewer] 데이터 검수 중 ---")
    
    slide_data = state["slide_data"]
    rules = state["template_details"]
    retry_count = state.get("retry_count", 0)
    
    # [안전장치] 3번 이상 빠꾸먹으면 그냥 통과시킴 (무한 루프 방지)
    if retry_count >= 3:
        print("   🚨 재시도 횟수 초과. 강제 통과합니다.")
        return {"review_status": "PASS", "review_feedback": "Max retries reached"}

    # LLM 설정
    llm = ChatOpenAI(model="gpt-4o", temperature=0)
    structured_llm = llm.with_structured_output(ReviewResult)
    
    system_prompt = f"""
    당신은 엄격한 PPT 품질 검수자(QA Auditor)입니다.
    현재 작성된 [슬라이드 데이터]가 [템플릿 규칙]을 완벽하게 준수하는지 검사하세요.

    [템플릿 규칙]
    {rules}

    [검사 항목]
    1. **제약 조건:** 글자 수 제한, 필수 포함 내용 등을 지켰는가?
    2. **데이터 누락:** 차트의 'values', 표의 'rows' 등이 비어있지 않은가?
    3. **스키마 준수:** Dynamic 컴포넌트의 데이터 구조가 올바른가?

    [작성된 슬라이드 데이터]
    {str(slide_data)}
    
    문제가 있다면 status="FAIL"과 함께 구체적인 피드백을 남기세요.
    """
    
    # 검수 실행
    result = structured_llm.invoke(system_prompt)
    
    print(f"   ⚖️ 판정: {result.status}")
    if result.status == "FAIL":
        print(f"   ❌ 지적사항: {result.feedback}")
        
    return {
        "review_status": result.status, 
        "review_feedback": result.feedback,
        "retry_count": retry_count + 1
    }


def content_node(state: AgentState):
    print("--- [Node: Content] 세부 내용 작성 중 ---")
    
    skeletons = state["skeleton_plan"]
    guide = state["template_details"]
    
    # [NEW] 피드백 확인
    feedback = state.get("review_feedback", "")
    current_data = state.get("slide_data", [])
    
    # 기본 프롬프트
    base_prompt = f"""
    당신은 PPT 콘텐츠 작가입니다. 
    기획안에 맞춰 내용을 작성하세요. 가이드의 제약조건을 반드시 지키세요.
    
    [기획안]
    {str(skeletons)}
    
    [가이드]
    {guide}
    """
    
    # [핵심] 재작성일 경우 프롬프트에 '수정 지시' 추가
    if feedback and feedback != "Good":
        print("   🔄 피드백 반영하여 수정 모드 진입")
        base_prompt += f"""
        
        !!! 긴급 수정 요청 !!!
        이전 작성 결과에 심각한 오류가 발견되었습니다.
        아래 피드백을 반영하여 데이터를 **처음부터 다시 올바르게 작성**하세요.
        
        [지적 사항]
        {feedback}
        
        [이전 작성 데이터 (참고용)]
        {str(current_data)}
        """

    # ... (LLM 호출 로직은 기존과 동일) ...
    # result = structured_llm.invoke(...)
    
    return {"slide_data": result.slides} # 수정된 데이터 반환


from langgraph.graph import StateGraph, END

# 라우팅 함수 (표지판 역할)
def route_after_review(state: AgentState):
    if state["review_status"] == "FAIL":
        return "content_node" # 다시 작성하러 돌아갓!
    else:
        return "renderer_node" # 합격! 인쇄하러 가자.

# 그래프 정의
workflow = StateGraph(AgentState)

# 노드 등록
workflow.add_node("scanner", scanner_node)
workflow.add_node("structure", structure_node)
workflow.add_node("content", content_node)
workflow.add_node("reviewer", reviewer_node) # NEW
workflow.add_node("renderer", renderer_node)

# 엣지 연결
workflow.set_entry_point("scanner")
workflow.add_edge("scanner", "structure")
workflow.add_edge("structure", "content")
workflow.add_edge("content", "reviewer") # 작성 후엔 무조건 검수

# [핵심] 조건부 엣지 (PASS냐 FAIL이냐)
workflow.add_conditional_edges(
    "reviewer",
    route_after_review,
    {
        "content_node": "content",   # FAIL이면 여기로
        "renderer_node": "renderer"  # PASS면 여기로
    }
)

workflow.add_edge("renderer", END)

app = workflow.compile()

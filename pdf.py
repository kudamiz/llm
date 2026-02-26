from typing import List, TypedDict, Literal
from langgraph.graph import StateGraph, END

# ==========================================
# 1. 상태(State) 정의
# ==========================================
class GraphState(TypedDict):
    question: str
    documents: List[str]
    generation: str
    needs_rewrite: bool

# ==========================================
# 2. 노드(Node) 함수 정의
# ==========================================
def retrieve(state: GraphState):
    """Vector DB에서 문서를 검색합니다."""
    print("▶ [NODE] RETRIEVE: Vector DB에서 문서 검색 중...")
    
    # TODO: 실제 구현 시 retriever.invoke(state["question"]) 사용
    documents = ["PDF에서 추출한 관련 텍스트 조각 A", "PDF에서 추출한 관련 텍스트 조각 B"] 
    
    return {"documents": documents}

def grade_documents(state: GraphState):
    """검색된 문서가 질문에 답하기 적절한지 평가합니다."""
    print("▶ [NODE] GRADE_DOCUMENTS: 검색된 문서의 유효성 평가 중...")
    
    # TODO: 실제 구현 시 LLM을 호출하여 문서 관련성 평가
    # 테스트를 위해 질문에 '재작성'이라는 단어가 있으면 부적절하다고 가정
    if "재작성" in state["question"]:
        print("   -> 문서가 질문과 관련 없음! 질문 재작성 필요.")
        return {"needs_rewrite": True} 
    else:
        print("   -> 문서가 질문과 관련 있음! 답변 생성 가능.")
        return {"needs_rewrite": False}

def generate(state: GraphState):
    """문서를 바탕으로 최종 답변을 생성합니다."""
    print("▶ [NODE] GENERATE: 최종 답변 생성 중...")
    
    # TODO: 실제 구현 시 LLM에 Prompt + Question + Documents를 넣고 답변 생성
    generation = f"'{state['documents'][0]}' 등을 참고하여 만든 최종 답변입니다."
    
    return {"generation": generation}

def rewrite_query(state: GraphState):
    """문서가 적절하지 않을 경우, 질문을 더 명확하게 수정합니다."""
    print("▶ [NODE] REWRITE_QUERY: 질문을 더 검색하기 좋게 수정 중...")
    
    # TODO: 실제 구현 시 LLM을 사용해 질문 수정
    # 테스트를 위해 '재작성' 단어를 빼고 키워드를 추가함
    better_question = state["question"].replace("재작성", "") + " (상세 키워드 추가됨)"
    
    return {"question": better_question}

# ==========================================
# 3. 조건부 라우팅 함수
# ==========================================
def decide_to_generate(state: GraphState) -> Literal["rewrite", "generate"]:
    """평가 결과에 따라 다음 노드를 결정합니다."""
    print("🔄 [ROUTING] 평가 결과 분석 중...")
    if state["needs_rewrite"]:
        return "rewrite"
    else:
        return "generate"

# ==========================================
# 4. 그래프 조립 및 컴파일
# ==========================================
def build_graph():
    workflow = StateGraph(GraphState)

    # 노드 추가
    workflow.add_node("retrieve", retrieve)
    workflow.add_node("grade_documents", grade_documents)
    workflow.add_node("generate", generate)
    workflow.add_node("rewrite_query", rewrite_query)

    # 기본 흐름 연결
    workflow.set_entry_point("retrieve")
    workflow.add_edge("retrieve", "grade_documents")

    # 조건부 흐름 연결
    workflow.add_conditional_edges(
        "grade_documents",
        decide_to_generate,
        {
            "rewrite": "rewrite_query",
            "generate": "generate",
        }
    )

    # 순환 및 종료 연결
    workflow.add_edge("rewrite_query", "retrieve") # 질문 수정 후 다시 검색
    workflow.add_edge("generate", END)             # 생성 완료 시 종료

    return workflow.compile()

# ==========================================
# 5. 실행 테스트
# ==========================================
if __name__ == "__main__":
    app = build_graph()

    print("\n=== 테스트 1: 정상적인 질문 (바로 답변 생성) ===")
    inputs_1 = {"question": "이 PDF의 핵심 요약은 뭐야?", "needs_rewrite": False}
    for output in app.stream(inputs_1):
        pass # 내부 print문 출력 확인용

    print("\n\n=== 테스트 2: 재작성이 필요한 질문 (순환 구조 테스트) ===")
    # '재작성' 이라는 단어를 넣어 고의로 fail을 유도 -> 재작성 -> 재검색 흐름 확인
    inputs_2 = {"question": "이 PDF 내용 좀 재작성 테스트해봐", "needs_rewrite": False}
    for output in app.stream(inputs_2):
        pass
    
    print("\n✅ 최종 완료!")

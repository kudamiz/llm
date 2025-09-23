from bs4 import BeautifulSoup
from langchain.text_splitter import RecursiveCharacterTextSplitter

def chunk_html_hybrid(
    html_content: str,
    chunk_size: int = 1000,
    chunk_overlap: int = 200,
):
    """
    HTML 콘텐츠를 구조와 의미를 고려하여 하이브리드 방식으로 청킹합니다.

    Args:
        html_content (str): 분석할 HTML 문서 문자열.
        chunk_size (int): 청크의 최대 크기 (글자 수 기준).
        chunk_overlap (int): 청크 간의 중복되는 글자 수.

    Returns:
        list[str]: 청킹된 텍스트 조각들의 리스트.
    """
    # 1. 구조적 청킹 (Structural Chunking)
    # --------------------------------------------------
    soup = BeautifulSoup(html_content, 'html.parser')

    # 스크립트, 스타일 등 불필요한 태그 제거
    for tag in soup(['script', 'style', 'nav', 'footer', 'aside']):
        tag.decompose()

    # 블록 레벨 태그를 기준으로 1차 청킹 후보군 생성
    # h 태그는 내용이 짧아도 중요한 의미를 가지므로 별도 처리 가능
    initial_chunks = []
    block_tags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'blockquote', 'div', 'table']
    
    for element in soup.find_all(True):
        if element.name in block_tags:
            text = element.get_text(separator=' ', strip=True)
            if len(text) > 10:  # 의미 없는 짧은 텍스트는 제외
                initial_chunks.append(text)

    if not initial_chunks: # 블록 태그가 전혀 없는 경우 전체 텍스트 사용
        initial_chunks.append(soup.get_text(separator=' ', strip=True))

    # 2. 재귀적/의미적 청킹 (Recursive Chunking)
    # --------------------------------------------------
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=chunk_size,
        chunk_overlap=chunk_overlap,
        length_function=len,
        is_separator_regex=False,
        separators=["\n\n", "\n", ". ", " ", ""], # 문단 -> 문장 -> 단어 순으로 분리
    )

    final_chunks = []
    for chunk in initial_chunks:
        # 1차 청킹된 덩어리가 너무 길면, 재귀적으로 다시 잘게 나눔
        if len(chunk) > chunk_size:
            smaller_chunks = text_splitter.split_text(chunk)
            final_chunks.extend(smaller_chunks)
        else:
            final_chunks.append(chunk)

    return final_chunks

# --- 사용 예시 ---

# 분석할 샘플 HTML (다양한 케이스 포함)
html_doc = """
<html>
<head><title>테스트 문서</title></head>
<body>
    <h1>HTML 청킹 테스트</h1>
    <p>이것은 첫 번째 문단입니다. BeautifulSoup와 LangChain을 활용한 하이브리드 청킹은 매우 효과적입니다. 구조적 정보를 먼저 활용하여 큰 의미 단위를 나누고, 너무 긴 덩어리는 재귀적으로 분할하여 최종 결과물의 품질을 높입니다. 이 문단은 일부러 길게 작성되었습니다.</p>
    
    <div>
        <h2>중요한 하위 주제</h2>
        이것은 div 태그 안에 있는 텍스트입니다. 줄바꿈은<br>br 태그를 사용했습니다. 
        규칙성이 없는 HTML도 잘 처리해야 합니다.
    </div>
    
    <p>이것은 짧은 두 번째 문단입니다.</p>
    
    <ul>
        <li>첫 번째 목록 항목: 구조를 유지하는 것이 중요합니다.</li>
        <li>두 번째 목록 항목: 각 항목은 별도의 청크로 인식되어야 합니다.</li>
    </ul>

    <table>
      <tr><th>이름</th><th>나이</th></tr>
      <tr><td>홍길동</td><td>30</td></tr>
      <tr><td>이순신</td><td>45</td></tr>
    </table>
    
    <script>console.log("이 부분은 제거되어야 합니다.");</script>
</body>
</html>
"""

# 함수 호출
chunks = chunk_html_hybrid(html_doc, chunk_size=200, chunk_overlap=40)

# 결과 출력
print(f"총 {len(chunks)}개의 청크로 분리되었습니다.\n")
for i, chunk in enumerate(chunks):
    print(f"--- 청크 {i+1} (길이: {len(chunk)}) ---")
    print(chunk)
    print()

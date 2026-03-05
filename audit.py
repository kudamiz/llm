import pandas as pd
import json
import openai

# API 키 설정 (본인의 OpenAI API 키 입력)
openai.api_key = "sk-your-api-key-here"

def analyze_excel_structure_with_llm(sample_csv_text):
    """
    LLM을 사용하여 엑셀 샘플 데이터에서 질문/답변 컬럼 위치와 데이터 시작 행을 찾습니다.
    """
    system_prompt = """
    너는 데이터 분석 전문가야. 
    사용자가 엑셀 파일의 상위 20행을 CSV 형식으로 제공할 거야.
    이 데이터를 분석해서 '질문(Question)'과 '답변(Answer)'에 해당하는 컬럼의 인덱스(0부터 시작)와,
    실제 데이터(헤더 제외)가 시작되는 행 인덱스(0부터 시작)를 찾아줘.
    
    반드시 아래와 같은 JSON 형식으로만 대답해:
    {
        "start_row_idx": 5, 
        "question_col_idx": 1, 
        "answer_col_idx": 2
    }
    """

    response = openai.chat.completions.create(
        model="gpt-4o", # 빠르고 저렴한 gpt-4o-mini를 사용해도 좋습니다.
        response_format={ "type": "json_object" },
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"다음은 엑셀 샘플 데이터야:\n\n{sample_csv_text}"}
        ],
        temperature=0.1 # 일관된 결과를 위해 온도를 낮춤
    )
    
    # LLM이 응답한 JSON 문자열을 파이썬 딕셔너리로 변환
    result_json = response.choices[0].message.content
    return json.loads(result_json)

def extract_qna_from_excel(file_path):
    """
    엑셀 파일을 읽어 질문과 답변 데이터만 추출합니다.
    """
    print(f"[{file_path}] 파일 분석 시작...")

    # 1. 샘플링: 컬럼명이나 형식을 모르므로 헤더 없이 상위 20행만 읽어옴
    try:
        sample_df = pd.read_excel(file_path, header=None, nrows=20)
    except Exception as e:
        return f"엑셀 파일을 읽는 중 오류 발생: {e}"

    # 결측치(NaN)를 빈 문자열로 처리하여 LLM이 헷갈리지 않게 함
    sample_csv = sample_df.fillna("").to_csv(index=False, header=False)

    # 2. LLM에게 구조 파악 요청
    print("LLM에게 구조 분석 요청 중...")
    try:
        structure = analyze_excel_structure_with_llm(sample_csv)
        start_row = structure.get("start_row_idx", 0)
        q_col = structure.get("question_col_idx", 0)
        a_col = structure.get("answer_col_idx", 1)
        print(f"✅ 분석 완료! 데이터 시작 행: {start_row}, 질문 컬럼: {q_col}, 답변 컬럼: {a_col}")
    except Exception as e:
        return f"LLM 분석 중 오류 발생: {e}"

    # 3. 파악된 구조를 바탕으로 전체 데이터 정확히 추출
    # header 매개변수를 사용하여 데이터가 시작되는 바로 윗줄을 헤더로 인식하게 하거나,
    # 인덱스 기반으로 슬라이싱할 수 있습니다. 여기서는 인덱스 슬라이싱 사용.
    
    full_df = pd.read_excel(file_path, header=None)
    
    # 실제 데이터가 있는 행부터 끝까지 자르기
    data_df = full_df.iloc[start_row:].copy()
    
    # 질문과 답변 컬럼만 선택
    try:
        qna_df = data_df.iloc[:, [q_col, a_col]].copy()
    except IndexError:
        return "LLM이 잘못된 컬럼 인덱스를 반환했습니다. 엑셀 양식을 확인해주세요."

    # 컬럼명 통일 (프론트엔드로 보내기 좋게)
    qna_df.columns = ["Question", "Answer"]
    
    # 질문이나 답변이 완전히 비어있는 행은 제거 (정제 과정)
    qna_df = qna_df.dropna(how='all')

    print("🎉 데이터 추출 성공!")
    return qna_df

# --- 실행 예시 ---
if __name__ == "__main__":
    # 테스트용 엑셀 파일 경로
    # test_file = "sample_qna.xlsx" 
    # extracted_data = extract_qna_from_excel(test_file)
    # print(extracted_data.head())
    
    # 프론트엔드로 전달할 때는 JSON 형태로 변환하면 좋습니다.
    # frontend_json = extracted_data.to_dict(orient="records")
    pass

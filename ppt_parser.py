from unstructured.partition.pptx import partition_pptx
import os

# 1. 경로 설정
output_image_dir = "./extracted_images"
os.makedirs(output_image_dir, exist_ok=True)

filename = "example_presentation.pptx"

# 2. PPT 파싱 (핵심 부분)
# extract_images_in_pdf=True 옵션은 PPTX에서도 작동하여 포함된 그림을 추출해줍니다.
elements = partition_pptx(
    filename=filename,
    extract_images_in_pdf=True,  # 이미지/차트 추출 활성화
    infer_table_structure=True,  # 표 구조(html) 추출 활성화
    image_output_dir_path=output_image_dir, # 추출된 이미지 저장 경로
)

# 3. 요소별 데이터 분류 (RAG용 데이터 전처리)
text_elements = []
table_elements = []
image_elements = []

for element in elements:
    # 요소의 타입 확인
    el_type = element.category
    
    if el_type == "Table":
        # 표는 HTML 메타데이터와 텍스트를 함께 저장
        table_elements.append({
            "text": element.text,
            "html": element.metadata.text_as_html,
            "page": element.metadata.page_number
        })
    
    elif el_type == "Image":
        # 이미지는 저장된 경로를 참조
        image_elements.append({
            "path": element.metadata.image_path,
            "page": element.metadata.page_number
        })
        
    elif el_type in ["Title", "NarrativeText", "ListItem"]:
        # 일반 텍스트
        text_elements.append({
            "text": element.text,
            "page": element.metadata.page_number
        })

print(f"텍스트 청크: {len(text_elements)}개")
print(f"추출된 표: {len(table_elements)}개")
print(f"추출된 이미지(차트 등): {len(image_elements)}개")

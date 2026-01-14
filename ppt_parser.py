8from unstructured.partition.pptx import partition_pptx
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


import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

def iter_shapes(shapes):
    """그룹 안에 숨은 도형까지 샅샅이 뒤지는 재귀 함수"""
    for shape in shapes:
        # 1. 그룹인 경우: 재귀적으로 내부 진입
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)
        else:
            yield shape

def extract_images_from_pptx(pptx_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    prs = Presentation(pptx_path)
    image_count = 0

    print(f"이미지 추출 시작: {pptx_path}")

    for i, slide in enumerate(prs.slides):
        # 슬라이드 내의 모든 도형(그룹 포함)을 순회
        for shape in iter_shapes(slide.shapes):
            
            # 2. 그림(Picture)인 경우
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image = shape.image
                    # 이미지 바이너리 데이터 가져오기
                    image_bytes = image.blob
                    # 확장자 결정 (jpg, png 등)
                    ext = image.ext
                    
                    filename = f"slide_{i+1}_img_{image_count}.{ext}"
                    filepath = os.path.join(output_dir, filename)
                    
                    with open(filepath, "wb") as f:
                        f.write(image_bytes)
                        
                    print(f"  [저장됨] {filename}")
                    image_count += 1
                except Exception as e:
                    print(f"  [에러] 이미지 저장 실패: {e}")

    print(f"총 {image_count}개의 이미지를 추출했습니다.")

# --- 실행 ---
extract_images_from_pptx("example.pptx", "./extracted_images")

import os
import subprocess
from pdf2image import convert_from_path

def ppt_to_images_via_libreoffice(ppt_path, output_dir):
    """
    LibreOffice를 이용해 PPT를 이미지로 변환하는 함수
    1. PPT -> PDF (LibreOffice headless 모드 사용)
    2. PDF -> Images (pdf2image 사용)
    """
    
    # 0. 경로 설정
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    ppt_dir = os.path.dirname(ppt_path)
    ppt_filename = os.path.basename(ppt_path)
    pdf_filename = os.path.splitext(ppt_filename)[0] + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_filename) # PDF도 일단 output_dir에 저장

    print(f"1. PDF 변환 시작: {ppt_filename}...")

    # 1. LibreOffice를 이용해 PDF로 변환 (터미널 명령어 실행)
    # --headless: 화면 없이 실행
    # --convert-to pdf: PDF로 변환
    # --outdir: 저장할 폴더
    command = [
        "libreoffice", 
        "--headless", 
        "--convert-to", "pdf", 
        "--outdir", output_dir, 
        ppt_path
    ]
    
    try:
        # subprocess로 리눅스 명령어 실행
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        # print(result.stdout.decode()) # 로그 보고 싶으면 주석 해제
    except subprocess.CalledProcessError as e:
        print(f"❌ LibreOffice 변환 실패: {e.stderr.decode()}")
        return []

    if not os.path.exists(pdf_path):
        print("❌ PDF 파일이 생성되지 않았습니다.")
        return []

    print("2. 이미지 변환 시작 (PDF -> Images)...")

    # 2. 변환된 PDF를 이미지로 쪼개기
    try:
        # dpi=300 : 고화질 설정 (OCR/VLM 인식률 높이려면 300 추천)
        images = convert_from_path(pdf_path, dpi=300)
        
        saved_image_paths = []
        for i, image in enumerate(images):
            # 슬라이드 번호는 1부터 시작
            image_filename = f"slide_{i+1}.jpg"
            save_path = os.path.join(output_dir, image_filename)
            
            image.save(save_path, "JPEG")
            saved_image_paths.append(save_path)
            print(f"  - 저장됨: {save_path}")
            
        print(f"✅ 변환 완료! 총 {len(saved_image_paths)}장")
        
        # (선택) 중간에 만든 PDF는 삭제하고 싶다면:
        # os.remove(pdf_path)
        
        return saved_image_paths

    except Exception as e:
        print(f"❌ 이미지 변환 실패: {e}")
        return []

# --- 실행 예시 ---
# ppt_file = "./data/defect_report.pptx"
# output_folder = "./extracted_images/report_01"

# images = ppt_to_images_via_libreoffice(ppt_file, output_folder)

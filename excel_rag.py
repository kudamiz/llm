import pandas as pd
import subprocess
import os
from pdf2image import convert_from_path
from PIL import Image

def extract_text_as_markdown(excel_path):
    """Track B: pandas를 사용해 엑셀의 표 데이터를 마크다운으로 추출"""
    try:
        # 모든 시트를 읽어옵니다
        excel_data = pd.read_excel(excel_path, sheet_name=None)
        markdown_text = ""
        
        for sheet_name, df in excel_data.items():
            markdown_text += f"### Sheet: {sheet_name}\n"
            # 빈 값이 있는 경우 처리 및 마크다운 변환
            markdown_text += df.fillna("").to_markdown(index=False)
            markdown_text += "\n\n"
            
        return markdown_text
    except Exception as e:
        print(f"텍스트 추출 오류: {e}")
        return ""

def convert_excel_to_images(excel_path, output_dir):
    """Track A: LibreOffice를 사용해 PDF 변환 후 이미지로 분할"""
    excel_path = os.path.abspath(excel_path)
    output_dir = os.path.abspath(output_dir)
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    pdf_filename = os.path.splitext(os.path.basename(excel_path))[0] + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)
    
    try:
        # 1. LibreOffice를 사용하여 백그라운드에서 PDF로 강제 변환
        command = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            excel_path
        ]
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("PDF 변환 완료.")
        
        # 2. 생성된 PDF를 이미지 리스트로 변환 (DPI 200 설정으로 가독성 확보)
        images = convert_from_path(pdf_path, dpi=200)
        image_paths = []
        
        # 3. 각 페이지를 PNG 이미지로 저장
        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, f"page_{i + 1}.png")
            image.save(image_path, "PNG")
            image_paths.append(image_path)
            
        print(f"총 {len(image_paths)}장의 이미지 분할 완료.")
        
        # 임시 PDF 파일 삭제 (용량 관리)
        os.remove(pdf_path)
        
        return image_paths
    except Exception as e:
        print(f"이미지 변환 오류: {e}")
        return []

def prepare_vlm_payload(excel_path, output_dir="./output"):
    """투트랙 데이터를 병합하여 VLM에 보낼 준비를 하는 메인 함수"""
    print(f"[{os.path.basename(excel_path)}] 하이브리드 파이프라인 처리 시작...")
    
    # 1. 텍스트 데이터 추출
    markdown_context = extract_text_as_markdown(excel_path)
    
    # 2. 이미지 데이터 추출
    image_paths = convert_excel_to_images(excel_path, output_dir)
    
    # 3. LLM에 전달할 최종 페이로드(Payload) 구성
    vlm_payloads = []
    
    # 각 이미지(페이지)마다 전체 텍스트 컨텍스트를 묶어서 전달
    for img_path in image_paths:
        payload = {
            "image_path": img_path,
            "system_prompt": (
                "너는 엑셀 문서를 RAG 데이터베이스용 구조화된 마크다운으로 변환하는 전문가야.\n"
                "첨부된 이미지는 문서의 시각적 형태이고, 아래 텍스트는 원본 데이터야.\n"
                "이미지 안의 레이아웃과 차트를 설명하고, 표를 그릴 때는 반드시 아래 텍스트의 숫자를 우선 참조해.\n\n"
                f"### 원본 텍스트 데이터:\n{markdown_context}"
            )
        }
        vlm_payloads.append(payload)
        
    print("VLM 페이로드 구성 완료.")
    return vlm_payloads

# 실행 예시
# payloads = prepare_vlm_payload("sample_report.xlsx")
# for idx, data in enumerate(payloads):
#     print(f"--- Payload {idx+1} ---")
#     print(f"Image: {data['image_path']}")
#     print(f"Prompt Length: {len(data['system_prompt'])} characters\n")


import os
import nest_asyncio
from llama_parse import LlamaParse

# Jupyter 환경 등에서 비동기 실행을 위한 설정
nest_asyncio.apply()

def parse_excel_with_multimodal_ai(excel_path, api_key):
    """
    최신 AI 파서인 LlamaParse를 이용해 엑셀의 시각적 맥락(순서도, 차트 등)을
    포함하여 문서를 통째로 분석하는 실습 코드입니다.
    """
    os.environ["LLAMA_CLOUD_API_KEY"] = api_key

    print(f"[{excel_path}] 멀티모달 AI 파싱을 시작합니다. (시각적 레이아웃 분석 중...)")

    # LlamaParse 초기화 
    # premium_mode를 켜면 내부적으로 VLM을 사용하여 차트와 순서도의 의미와 배치를 해석합니다.
    parser = LlamaParse(
        result_type="markdown",  # 최종 RAG 적재용 포맷
        premium_mode=True,       # VLM 기반의 복잡한 시각적 객체 해석 활성화
        verbose=True
    )

    # 엑셀 파일 파싱 실행 (클라우드 엔진이 엑셀을 통째로 렌더링하여 맥락을 파악함)
    documents = parser.load_data(excel_path)

    # 결과 출력
    for i, doc in enumerate(documents):
        print(f"\n--- [시각적 맥락이 반영된 파싱 결과 {i+1}] ---")
        # 내용이 길 수 있으므로 앞부분만 출력
        print(doc.text[:800] + "\n\n... [중략] ...")
        
    return documents

# 실행 예시
# api_key = "llx-your-api-key-here"
# parsed_docs = parse_excel_with_multimodal_ai("complex_sample.xlsx", api_key)


import openpyxl
import subprocess
import os
from pdf2image import convert_from_path

def convert_excel_without_clipping(excel_path, output_dir="./output"):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    temp_excel_path = os.path.join(output_dir, f"{base_name}_temp_scaled.xlsx")
    
    print("1단계: openpyxl을 사용해 잘림 방지 인쇄 설정 주입 중...")
    
    # 1. 엑셀 파일 로드
    wb = openpyxl.load_workbook(excel_path)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # 핵심 설정 1: '자동 맞춤' 활성화
        ws.page_setup.fitToPage = True
        
        # 핵심 설정 2: 가로(Width)는 무조건 1페이지 안에 다 들어오도록 압축
        ws.page_setup.fitToWidth = 1
        
        # 핵심 설정 3: 세로(Height)는 데이터 길이에 따라 자연스럽게 여러 장으로 나뉘도록 제한 해제 (0 설정)
        ws.page_setup.fitToHeight = 0 
        
        # 여백을 최소화하여 공간 확보 (단위: 인치)
        ws.page_margins.left = 0.1
        ws.page_margins.right = 0.1
        ws.page_margins.top = 0.1
        ws.page_margins.bottom = 0.1

    # 조작된 설정을 적용하여 임시 파일로 저장
    wb.save(temp_excel_path)
    wb.close()
    
    print("2단계: 주입된 임시 파일을 LibreOffice로 PDF 변환 중...")
    
    # 2. LibreOffice 실행 (이제 가로 너비가 무조건 1페이지에 맞춰져서 나옵니다)
    command = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        temp_excel_path
    ]
    subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    pdf_path = os.path.join(output_dir, f"{base_name}_temp_scaled.pdf")
    
    print("3단계: 변환된 PDF를 고화질 이미지로 추출 중...")
    
    # 3. PDF를 다시 이미지로 변환 (필요시)
    images = convert_from_path(pdf_path, dpi=300) # 글씨가 작아질 수 있으므로 고해상도(DPI 300) 권장
    
    image_paths = []
    for i, image in enumerate(images):
        img_path = os.path.join(output_dir, f"{base_name}_page_{i+1}.png")
        image.save(img_path, "PNG")
        image_paths.append(img_path)
        
    # 흔적 지우기 (임시 엑셀 파일 및 PDF 삭제)
    os.remove(temp_excel_path)
    os.remove(pdf_path)
    
    print(f"✅ 완료! 우측 잘림 없이 총 {len(image_paths)}장의 이미지가 생성되었습니다.")
    return image_paths

# 실행 예시
# safe_images = convert_excel_without_clipping("my_wide_excel.xlsx")

import zipfile
import os
import re
import subprocess
from pdf2image import convert_from_path

def safe_convert_without_clipping(excel_path, output_dir="./output"):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    temp_excel_path = os.path.join(output_dir, f"{base_name}_patched.xlsx")
    
    print("1단계: 미디어 원본 보존을 위한 XML 인젝션(Injection) 진행 중...")
    
    # 1. 원본 엑셀(ZIP) 파일을 열어서 읽으면서, 동시에 새로운 파일로 복사합니다.
    with zipfile.ZipFile(excel_path, 'r') as zin:
        with zipfile.ZipFile(temp_excel_path, 'w') as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                
                # 시트 설정을 담당하는 XML 파일만 타겟으로 잡아 수정합니다.
                if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
                    xml_str = content.decode('utf-8')
                    
                    # '1페이지에 가로 너비 맞춤'을 강제하는 XML 태그 주입
                    setup_tag = '<pageSetup fitToPage="1" fitToWidth="1" fitToHeight="0" orientation="landscape"/>'
                    
                    # 기존에 pageSetup 태그가 있으면 덮어쓰고, 없으면 적절한 위치에 끼워 넣습니다.
                    if '<pageSetup' in xml_str:
                        xml_str = re.sub(r'<pageSetup[^>]*>', setup_tag, xml_str)
                    else:
                        # 통상적으로 <pageMargins> 태그 바로 앞에 위치해야 에러가 나지 않습니다.
                        xml_str = xml_str.replace('<pageMargins', f'{setup_tag}<pageMargins', 1)
                        
                    # 엑셀이 인쇄 설정을 인식하도록 sheetPr 속성도 활성화해 줍니다.
                    if '<sheetPr' not in xml_str:
                        xml_str = re.sub(r'(<worksheet[^>]*>)', r'\1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', xml_str, 1)
                    elif 'fitToPage=' not in xml_str:
                        xml_str = re.sub(r'(<sheetPr[^>]*>)', r'\1<pageSetUpPr fitToPage="1"/>', xml_str, 1)
                        
                    content = xml_str.encode('utf-8')
                    
                # 수정된 XML(또는 원본 미디어 파일)을 새 압축 파일에 그대로 씁니다.
                zout.writestr(item, content)

    print("2단계: 이미지 증발이 없는 안전한 파일로 LibreOffice PDF 변환 중...")
    
    # 2. 이제 이 patched 파일을 LibreOffice에 넘깁니다. (이미지 100% 보존됨)
    command = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        temp_excel_path
    ]
    subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    pdf_path = os.path.join(output_dir, f"{base_name}_patched.pdf")
    
    print("3단계: 변환된 PDF를 고화질 이미지로 추출 중...")
    
    # 3. PDF를 고해상도(DPI 300) 이미지로 변환
    images = convert_from_path(pdf_path, dpi=300)
    
    image_paths = []
    for i, image in enumerate(images):
        img_path = os.path.join(output_dir, f"{base_name}_page_{i+1}.png")
        image.save(img_path, "PNG")
        image_paths.append(img_path)
        
    # 흔적 지우기
    os.remove(temp_excel_path)
    os.remove(pdf_path)
    
    print(f"✅ 완료! 우측 잘림 현상과 이미지 증발 없이 완벽히 캡처되었습니다.")
    return image_paths

# 실행 예시
# final_images = safe_convert_without_clipping("my_complex_excel.xlsx"


import os
from pdf2image import convert_from_path
from PIL import Image

def split_pdf_with_overlap(pdf_path, output_dir="./overlap_slices", window_height=1200, overlap=300):
    """
    PDF를 하나의 거대한 이미지로 이어 붙인 후, 위아래가 겹치도록 분할하는 실습 코드.
    - window_height: 잘라낼 각 조각의 세로 길이(픽셀)
    - overlap: 조각끼리 겹치게 할 세로 길이(픽셀)
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    print(f"[{base_name}] 1단계: PDF를 이미지로 변환 중...")
    # PDF의 모든 페이지를 고해상도(DPI 300) 이미지로 변환
    pages = convert_from_path(pdf_path, dpi=300)
    
    if not pages:
        print("PDF에서 이미지를 추출하지 못했습니다.")
        return []

    print(f"[{base_name}] 2단계: {len(pages)}장의 페이지를 세로로 이어 붙이는 중...")
    # 전체 캔버스 크기 계산 (가로는 첫 페이지 기준, 세로는 모든 페이지 높이의 합)
    total_width = pages[0].width
    total_height = sum(page.height for page in pages)
    
    # 거대한 빈 캔버스 생성
    stitched_image = Image.new('RGB', (total_width, total_height))
    
    # 페이지들을 차례대로 이어 붙이기
    y_offset = 0
    for page in pages:
        stitched_image.paste(page, (0, y_offset))
        y_offset += page.height

    print(f"[{base_name}] 3단계: 오버랩 슬라이싱(겹쳐 자르기) 진행 중...")
    
    current_y = 0
    slice_count = 0
    slice_paths = []

    # 슬라이딩 윈도우 방식으로 이미지 자르기
    while current_y < total_height:
        # 자를 영역의 끝 지점 계산
        end_y = min(current_y + window_height, total_height)
        
        # (좌, 상, 우, 하) 좌표로 크롭
        crop_region = (0, current_y, total_width, end_y)
        slice_img = stitched_image.crop(crop_region)
        
        # 조각 이미지 저장
        slice_filename = f"{base_name}_slice_{slice_count + 1}.png"
        slice_path = os.path.join(output_dir, slice_filename)
        slice_img.save(slice_path, "PNG")
        slice_paths.append(slice_path)
        
        # 다음 시작점 계산: 겹치는 부분(overlap)만큼 뒤로 무른다
        if end_y == total_height:
            break # 끝까지 다 잘랐으면 반복문 종료
            
        current_y = end_y - overlap
        slice_count += 1

    print(f"✅ 전처리 완료! 총 {len(slice_paths)}개의 오버랩 이미지 조각이 생성되었습니다.")
    return slice_paths

# 실행 예시 (가로 전체 유지, 세로 1200px씩 자르되 300px씩 겹침)
# result_images = split_pdf_with_overlap("sample_document.pdf", window_height=1200, overlap=300)


import os
from pdf2image import convert_from_path
from PIL import Image, ImageOps

def remove_white_margins(img, padding=30):
    """
    이미지에서 위아래의 텅 빈 하얀색 여백을 자동으로 계산하여 잘라내는 함수입니다.
    가로 너비는 유지하여 나중에 이어 붙일 때 어긋나지 않게 합니다.
    """
    # 1. 이미지를 흑백으로 변환하고 색상을 반전시킵니다. 
    # (흰색 배경은 검은색(0)이 되고, 글씨나 차트는 밝은색이 됩니다.)
    gray_img = img.convert("L")
    inverted_img = ImageOps.invert(gray_img)
    
    # 2. 내용이 존재하는 영역(0이 아닌 픽셀들)의 경계 상자(Bounding Box)를 찾습니다.
    bbox = inverted_img.getbbox()
    
    if bbox:
        # bbox = (좌, 상, 우, 하)
        _, upper, _, lower = bbox
        
        # 3. 너무 바짝 자르면 답답하므로 약간의 패딩(Padding)을 줍니다.
        upper = max(0, upper - padding)
        lower = min(img.height, lower + padding)
        
        # 4. 가로(Width)는 원본 그대로 두고, 세로(Height) 여백만 잘라냅니다.
        return img.crop((0, upper, img.width, lower))
    
    # 만약 페이지 전체가 완전 백지라면 원본을 그대로 반환 (또는 예외 처리 가능)
    return img

def split_pdf_with_smart_stitching(pdf_path, output_dir="./smart_slices", window_height=1200, overlap=300):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    print(f"[{base_name}] 1단계: PDF를 이미지로 변환 중...")
    raw_pages = convert_from_path(pdf_path, dpi=300)
    
    if not raw_pages:
        return []

    print(f"[{base_name}] 2단계: 각 페이지의 상하 빈 공간(여백) 자동 크롭 중...")
    # 핵심 실습: 각 페이지를 이어붙이기 전에 여백부터 제거합니다.
    cropped_pages = [remove_white_margins(page) for page in raw_pages]
    
    print(f"[{base_name}] 3단계: 여백이 제거된 알맹이들만 세로로 이어 붙이는 중...")
    total_width = cropped_pages[0].width
    total_height = sum(page.height for page in cropped_pages)
    
    stitched_image = Image.new('RGB', (total_width, total_height), color='white')
    
    y_offset = 0
    for page in cropped_pages:
        stitched_image.paste(page, (0, y_offset))
        y_offset += page.height

    print(f"[{base_name}] 4단계: 밀도 높은 이미지를 오버랩 슬라이싱 진행 중...")
    current_y = 0
    slice_count = 0
    slice_paths = []

    while current_y < total_height:
        end_y = min(current_y + window_height, total_height)
        crop_region = (0, current_y, total_width, end_y)
        slice_img = stitched_image.crop(crop_region)
        
        slice_filename = f"{base_name}_dense_slice_{slice_count + 1}.png"
        slice_path = os.path.join(output_dir, slice_filename)
        slice_img.save(slice_path, "PNG")
        slice_paths.append(slice_path)
        
        if end_y == total_height:
            break
            
        current_y = end_y - overlap
        slice_count += 1

    print(f"✅ 전처리 완료! 공백이 제거된 알찬 {len(slice_paths)}개의 오버랩 조각이 생성되었습니다.")
    return slice_paths

# 실행 예시
# result = split_pdf_with_smart_stitching("presentation.pdf", window_height=1000, overlap=250)


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

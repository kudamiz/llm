import win32com.client
import os

def remove_ppt_protection(input_path, output_path, password=""):
    # PowerPoint 애플리케이션 시작
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    
    # 백그라운드 실행을 원할 경우 (일부 DRM 모듈은 창이 보여야 작동할 수 있음)
    # ppt_app.Visible = True 
    
    presentation = None
    
    try:
        # 파일 열기
        # 암호가 있는 경우 Open 메서드에서 지정 가능
        # WithWindow=msoFalse로 하면 창을 띄우지 않고 엽니다 (오류 가능성 있음)
        presentation = ppt_app.Presentations.Open(
            FileName=input_path, 
            ReadOnly=False, 
            WithWindow=True, 
            Password=password # 파일을 여는 암호가 있다면 입력
        )
        
        # (옵션) 쓰기 암호 등 내부 보안 설정 제거
        presentation.Password = ""
        presentation.WritePassword = ""
        
        # 다른 이름으로 저장 (보안 해제된 상태)
        presentation.SaveAs(output_path)
        print(f"성공: {output_path}에 저장되었습니다.")
        
    except Exception as e:
        print(f"오류 발생: {e}")
        
    finally:
        if presentation:
            presentation.Close()
        # PowerPoint 종료 (필요 시)
        # ppt_app.Quit()

# 사용 예시
input_file = os.path.abspath("protected.pptx")
output_file = os.path.abspath("unprotected.pptx")

# 파일이 열리는 암호를 알고 있다면 인자로 전달, DRM이 로그인 연동이라면 빈 문자열
remove_ppt_protection(input_file, output_file, password="known_password")


import win32com.client
import os

def extract_ppt_content(input_path, output_txt_path):
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True  # DRM 파일은 창이 보여야 내용이 로드되는 경우가 많음
    
    presentation = None
    extracted_text = []

    try:
        # 1. 파일 열기 (ReadOnly=True로 열어서 편집 권한 문제 회피)
        presentation = ppt_app.Presentations.Open(
            FileName=input_path, 
            ReadOnly=True, 
            WithWindow=True
        )
        
        print(f"[{input_path}] 파일 열기 성공. 데이터 추출 시작...")

        # 2. 슬라이드 순회
        for i, slide in enumerate(presentation.Slides):
            slide_content = []
            slide_title = f"=== Slide {i+1} ==="
            slide_content.append(slide_title)
            
            # 3. 슬라이드 내의 도형(Shape) 순회
            for shape in slide.Shapes:
                # 3-1. 텍스트 프레임(글상자)인 경우
                if shape.HasTextFrame:
                    if shape.TextFrame.HasText:
                        text = shape.TextFrame.TextRange.Text
                        slide_content.append(f"[Text] {text.strip()}")
                
                # 3-2. 표(Table)인 경우 (표 안의 텍스트도 추출)
                if shape.HasTable:
                    table = shape.Table
                    slide_content.append(f"[Table] ({table.Rows.Count}x{table.Columns.Count})")
                    for r in range(1, table.Rows.Count + 1):
                        row_text = []
                        for c in range(1, table.Columns.Count + 1):
                            try:
                                cell_text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                                row_text.append(cell_text.strip().replace('\r', ' '))
                            except:
                                pass
                        slide_content.append(f"  Row {r}: {' | '.join(row_text)}")

            # 슬라이드 하나의 내용 합치기
            extracted_text.append("\n".join(slide_content))

        # 4. 추출된 내용을 새로운 텍스트 파일로 저장 (Python으로 직접 쓰기)
        with open(output_txt_path, 'w', encoding='utf-8') as f:
            f.write("\n\n".join(extracted_text))
            
        print(f"추출 완료! 결과 파일: {output_txt_path}")

    except Exception as e:
        print(f"오류 발생: {e}")
        
    finally:
        if presentation:
            presentation.Close()
        # ppt_app.Quit() # 필요시 주석 해제

# 사용 예시
input_file = os.path.abspath("drm_protected.pptx")
output_txt = os.path.abspath("extracted_content.txt")

extract_ppt_content(input_file, output_txt)

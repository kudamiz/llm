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

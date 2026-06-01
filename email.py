import logging
from exchangelib import Credentials, Account, DELEGATE, Configuration
from exchangelib.errors import AutoDiscoverFailed, UnauthorizedError, TransportError

# 상세 에러 원인 분석을 위해 로깅 활성화
logging.basicConfig(level=logging.INFO)

def verify_exchange_connection(email, username, password, server_domain=None):
    """
    exchangelib의 사용 가능 여부를 검증하는 함수
    """
    # 1. 인증 정보 설정 (사내망의 경우 도메인\ID 형식 요구 가능)
    credentials = Credentials(username=username, password=password)
    
    try:
        if server_domain:
            # 자동 탐색(Autodiscover)이 막혀있을 경우를 대비한 수동 설정
            config = Configuration(server=server_domain, credentials=credentials)
            account = Account(
                primary_smtp_address=email, 
                config=config, 
                autodiscover=False, 
                access_type=DELEGATE
            )
        else:
            # 일반적인 자동 탐색 시도
            account = Account(
                primary_smtp_address=email, 
                credentials=credentials, 
                autodiscover=True, 
                access_type=DELEGATE
            )
        
        # 2. 실제 서버 데이터 호출 테스트 (수신함 메일 수 조회)
        inbox_count = account.inbox.total_count
        print(f"\n[성공] exchangelib 사용 가능 확인.")
        print(f"[정보] 연결 계정: {account.primary_smtp_address}")
        print(f"[정보] 현재 수신함 메일 수: {inbox_count}개")
        return True

    except AutoDiscoverFailed:
        print("\n[실패] Autodiscover 자동 탐색에 실패했습니다.")
        print("[대책] 사내 웹메일 URL 주소 등을 'server_domain' 인자에 명시하여 수동 접속해야 합니다.")
        return False
        
    except UnauthorizedError:
        print("\n[실패] 인증에 실패했거나 권한이 없습니다.")
        print("[대책] 계정/비밀번호를 확인하십시오. 사내 보안 정책상 EWS API 호출이 차단되었을 수 있습니다.")
        return False
        
    except TransportError as te:
        print(f"\n[실패] 네트워크 또는 프로토콜 오류 발생: {te}")
        print("[대책] 방화벽에 의해 포트(443)가 막혀있거나, Exchange 서버가 아닐 확률이 높습니다.")
        return False
        
    except Exception as e:
        print(f"\n[실패] 예외 오류 발생: {e}")
        return False

# --------------------------------------------------
# 실행 예시 (실제 사내 계정 정보 입력)
# --------------------------------------------------
USER_EMAIL = "your_email@company.com"
USER_ID = "company_domain\\your_id"  # 도메인 입력이 필요 없는 경우 'your_id'만 입력
USER_PW = "your_password"
EXCHANGE_SERVER = "mail.company.com" # 사내 웹메일 도메인 (필요시 입력, 모를 경우 None)

verify_exchange_connection(USER_EMAIL, USER_ID, USER_PW, server_domain=EXCHANGE_SERVER)

import asyncio
from mcp.client.sse import sse_client
from mcp import ClientSession

async def run():
    # 1. [귀 열기] SSE로 서버에 접속 (주소 확인 필수!)
    # 주의: Swagger에서 봤던 그 주소여야 합니다. (예: /mcp/sse 또는 /api/mcp/sse)
    url = "http://localhost:5000/mcp/sse" 

    print(f"Connecting to {url}...")
    
    async with sse_client(url) as (read, write):
        async with ClientSession(read, write) as session:
            # 2. [인사] 초기화 (Handshake)
            await session.initialize()
            
            # 3. [확인] 도구 목록 가져오기
            tools = await session.list_tools()
            print(f"\n✅ 연결 성공! 발견된 도구: {[t.name for t in tools.tools]}")
            
            # 4. [말하기] 도구 사용 요청 (여기를 님의 함수 이름으로 바꾸세요!)
            # 예: 만약 함수 이름이 'get_user'라면 아래처럼
            # result = await session.call_tool("get_user", arguments={"user_id": 1})
            
            # 일단 목록만 봐도 성공입니다!

if __name__ == "__main__":
    try:
        asyncio.run(run())
    except Exception as e:
        print(f"❌ 에러 발생: {e}")

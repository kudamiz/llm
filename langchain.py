# pip install langchain langchain-core chromadb pandas matplotlib
from langchain.agents import AgentExecutor, Tool
from langchain_core.prompts import ChatPromptTemplate
from langchain.chat_models import ChatOpenAI  # OpenAI 호환 사내 LLM 엔드포인트로 대체 가능
import pandas as pd
import matplotlib.pyplot as plt
import os, glob

# --- 1) 도구 정의 ---
def list_new_files(folder="data/inbox"):
    return glob.glob(os.path.join(folder, "*.csv"))

def profile_csv(path):
    df = pd.read_csv(path)
    info = {
        "columns": df.columns.tolist(),
        "n_rows": len(df),
        "numeric": [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])],
    }
    return {"path": path, **info}

def pick_columns(info_dict):
    cols = info_dict["columns"]
    # 간단 규칙 + LLM 판단으로 대체 가능: 여기선 규칙만 예시
    x = next((c for c in cols if "time" in c.lower() or "date" in c.lower()), cols[0])
    # numeric에서 y 하나 선택
    nums = info_dict["numeric"] or cols[1:2]
    y = nums[0]
    return {"x": x, "y": y}

def analyze_and_report(path, x, y, out="output"):
    os.makedirs(out, exist_ok=True)
    df = pd.read_csv(path)
    df[x] = pd.to_datetime(df[x], errors="coerce")
    df = df.dropna(subset=[x, y]).sort_values(x)
    # 아주 단순한 통계 + 그림 (여기에 Logic1/Logic2 등 고도화 가능)
    fig = plt.figure()
    plt.plot(df[x], df[y])
    plt.title(f"{os.path.basename(path)} — {y} over {x}")
    img = os.path.join(out, f"{os.path.basename(path)}.png")
    plt.savefig(img, bbox_inches="tight"); plt.close(fig)
    md = os.path.join(out, f"{os.path.basename(path)}.md")
    with open(md, "w", encoding="utf-8") as f:
        f.write(f"# Report for {os.path.basename(path)}\n\n")
        f.write(f"- rows: {len(df)}\n- x: `{x}`\n- y: `{y}`\n\n")
        f.write(f"![plot]({img})\n")
    return {"report": md, "image": img}

tools = [
    Tool(name="list_new_files", func=lambda : list_new_files(), description="새 CSV 목록을 가져온다"),
    Tool(name="profile_csv", func=profile_csv, description="CSV의 열/행/수치형 정보를 요약한다"),
    Tool(name="pick_columns", func=pick_columns, description="프로파일을 받아 x(시간)/y(값) 열을 선택한다"),
    Tool(name="analyze_and_report", func=lambda args: analyze_and_report(**args),
         description="선택된 x/y로 분석하고 리포트를 만든다"),
]

# --- 2) 에이전트(계획 + 실행 루프) ---
prompt = ChatPromptTemplate.from_template("""
너는 데이터 처리 에이전트다.
목표: 새 CSV를 찾고 → 구조 파악 → x/y 선택 → 리포트 생성.
필요할 때만 도구를 호출하고, 각 단계의 결과로 다음 행동을 정하라.
최종으로 생성된 report 경로를 알려줘.
""")

llm = ChatOpenAI(base_url="http://INTRANET_LLM:port/v1", api_key="DUMMY", model="internal-llm")  # 사내 LLM

agent = AgentExecutor.from_agent_and_tools(llm=llm, tools=tools, verbose=True)

result = agent.invoke({"input": "작업을 시작해."})
print(result)
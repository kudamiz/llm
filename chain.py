from langchain_openai import ChatOpenAI
from langchain_core.runnables import RunnableLambda, RunnableParallel, RunnablePassthrough
from langchain_core.documents import Document
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import PromptTemplate
from operator import itemgetter
import os, requests, json
from dotenv import load_dotenv

# OPENSEARCJ
host = 'com'
port = ...
auth = ()
index_name = '모델 번호'
top_n = 3
embedding_model_id = "TOKEN"

def retriever(query, embedding_model_id, top_n):
    url = f"https://{host}/{index_name}/_search"
    search = {
        "_source": {
            "excludes": ["passage_chunk_embedding"]
        },
        "query": {
            "nested": {
                "score_mode": "max",
                "path": "passage_chunk_embedding",
                "query": {
                    "neural": {
                        "passage_chunk_embedding.knn": {
                            "query_text": query,
                            "model_id": embedding_model_id
                        }
                    }
                }
            }
        }
    }

llm = ChatOpenAI(
    model="",
    openai_api_key="",
    openai_api_base="",
    temperature=0.5,
)

classification_prompt = PromptTemplate.from_template(
    """You are an assistant trained to categorize given user queries into three types : "Question", "Request", and 'Other'
Analyze the following user query and determine its category.
Just name the category, without any explanation.

Query : "{query}"
"""
)

classification_chain = (
    classification_prompt
    | llm
    | StrOutputParser()
)

def classification_route(info):
    query = info["query"]
    # question인 경우
    if "question" in info["topic"].lower():
        # summarize
        summarize_prompt = PromptTemplate.from_template(
            """You are an assistant trained to summarize given user query.
User query is a kind of question. Your mission is to clarify the question so that it is easier to understand and answer.
Don't answer, just summarize the question.
Summarize the question in a clear and concise manner.
Don't miss numbers and english words.
Answer only in one sentence.
Answer in KOREAN.

Query : "{query}"
"""
        )
        summarize_chain = (
            summarize_prompt
            | llm
            | StrOutputParser()
        )
        question = summarize_chain.invoke({"query": query})

        # answer
        retrieved_data = [
            {
                'retrieved_title': doc['_id'],
                'retrieved_answer': doc['_source']['text'],
                'score': doc['_score'],
                'url': doc['_source']['url']
            }
            for doc in retriever(question, embedding_model_id, top_n)
        ]

        answer_prompt = """
You are an assistant for question-answering tasks.
Before the answer, explain what was the question.
examples
PROMPT

Question: "{question}"
Retrieved_data: "{retrieved_data}"
"""
        prompt_template = PromptTemplate(
            template=answer_prompt,
            input_variables=["question", "retrieved_data"]
        )

        def render_prompt(inputs):
            return prompt_template.format(**inputs)

        question_chain = (
            {"question": lambda x: question, "retrieved_data": lambda x: retrieved_data}
            | RunnableLambda(render_prompt)
            | llm
        )

        return question_chain

    # request인 경우
    elif "request" in info["topic"].lower():
        request_chain = (
            PromptTemplate.from_template(
                """You are designed to answer only questions.
Answer "죄송합니다. 저는 질문에만 답변할 수 있습니다."
"""
            )
            | llm
        )
        return request_chain

    # 그 외의 경우
    else:
        other_chain = (
            PromptTemplate.from_template(
                """You are designed to answer only questions.
Answer "죄송합니다. 저는 질문에만 답변할 수 없습니다. 질문이 올바른지 확인 부탁드립니다."
"""
            )
            | llm
        )
        return other_chain

full_chain = (
    {"topic": classification_chain, "query": itemgetter("query")}
    | RunnableLambda(classification_route)
    | StrOutputParser()
)

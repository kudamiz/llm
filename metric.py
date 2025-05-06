# 기준 정답이 있다면 정답 chunk 포함 여부 확인
def evaluate_retrieval(retrieved_docs, ground_truth_chunks):
    retrieved_ids = [doc['_id'] for doc in retrieved_docs]
    true_positives = [doc_id for doc_id in retrieved_ids if doc_id in ground_truth_chunks]
    precision = len(true_positives) / len(retrieved_docs) if retrieved_docs else 0
    recall = len(true_positives) / len(ground_truth_chunks) if ground_truth_chunks else 0
    return {"precision": precision, "recall": recall}

from sentence_transformers import SentenceTransformer, util

# 사전 학습된 임베딩 모델 (로컬 또는 huggingface)
similarity_model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2')

def evaluate_similarity(answer, reference_answer):
    emb1 = similarity_model.encode(answer, convert_to_tensor=True)
    emb2 = similarity_model.encode(reference_answer, convert_to_tensor=True)
    similarity = util.cos_sim(emb1, emb2).item()
    return {"cosine_similarity": similarity}


evaluation_prompt = PromptTemplate.from_template("""
You are a helpful evaluator. Given the original question and the answer, evaluate the answer based on:
- Relevance to the question
- Factual correctness
- Clarity and conciseness

Rate the answer on a scale of 1 to 5.
Respond in JSON: {{"score": <score>, "reason": "<reason>"}}

Question: {question}
Answer: {answer}
""")
evaluation_chain = evaluation_prompt | llm | StrOutputParser()


paraphrase_prompt = PromptTemplate.from_template("""
Paraphrase the following question into a different phrasing that preserves its meaning:

"{question}"
""")
paraphrase_chain = paraphrase_prompt | llm | StrOutputParser()

def evaluate_robustness(original_question, answer_fn):
    paraphrased_question = paraphrase_chain.invoke({"question": original_question})
    paraphrased_answer = answer_fn(paraphrased_question)
    return {
        "paraphrased_question": paraphrased_question,
        "paraphrased_answer": paraphrased_answer
    }

# 필요한 라이브러리 설치 필요:
# pip install sentence-transformers scikit-learn pandas

import re
from itertools import combinations
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.cluster import AgglomerativeClustering
from sklearn.feature_extraction.text import TfidfVectorizer
import pandas as pd
import numpy as np

# 1. 예시 문서 chunk (RAG에서 쪼갠 상태라고 가정)
chunks = [
    "최근 rg 라인의 수율이 향상되었다. 기존 대비 공정 안정성이 좋아졌다.",
    "rigel 공정은 고속 라인에서 사용되며, rg와 동일 계열이다.",
    "canopus 제품은 cp로도 불리며, 주로 소비자용 라인에서 쓰인다.",
    "캐노는 cp 제품군 중 고성능 모델을 지칭한다.",
    "홍길동 부장이 보고서를 검토했다.",
    "김과장은 DRAM 라인에 대해 리뷰했다.",
    "강부장은 리지드 공정 이상에 대해 보고했다.",
    "리겔은 고성능 rg 제품이다. 클레임 이슈가 있었다."
]

# 2. TF-IDF로 주요 단어 자동 추출
vectorizer = TfidfVectorizer(ngram_range=(1, 2), max_features=30, token_pattern=r'\b\w+\b')
X = vectorizer.fit_transform(chunks)
auto_terms = vectorizer.get_feature_names_out()

# 3. 단어별 문맥 수집
term_contexts = {term: [] for term in auto_terms}
for chunk in chunks:
    for term in auto_terms:
        if re.search(rf"\b{re.escape(term)}\b", chunk):
            term_contexts[term].append(chunk)

# 4. 문맥 임베딩
model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
term_representations = {}
for term, contexts in term_contexts.items():
    if contexts:
        joined_context = " ".join(contexts)
        embedding = model.encode(joined_context)
        term_representations[term] = embedding

# 5. 유사한 단어쌍 추출
terms = list(term_representations.keys())
vectors = np.array([term_representations[t] for t in terms])
similarity_matrix = cosine_similarity(vectors)

similar_pairs = []
for i, j in combinations(range(len(terms)), 2):
    sim = similarity_matrix[i][j]
    if sim > 0.8:
        similar_pairs.append({
            "Entity A": terms[i],
            "Entity B": terms[j],
            "Cosine Similarity": round(sim, 4)
        })
df_similar = pd.DataFrame(similar_pairs)

# 6. 클러스터링
distance_matrix = 1 - similarity_matrix
clustering = AgglomerativeClustering(
    affinity='precomputed',
    linkage='average',
    distance_threshold=0.25,
    n_clusters=None
)
labels = clustering.fit_predict(distance_matrix)

cluster_df = pd.DataFrame({
    "Term": terms,
    "Cluster ID": labels
}).sort_values("Cluster ID")

# 7. NER 학습용 구조 생성
ner_data = []
for chunk in chunks:
    for term, cluster in zip(cluster_df["Term"], cluster_df["Cluster ID"]):
        for match in re.finditer(rf"\b{re.escape(term)}\b", chunk):
            ner_data.append({
                "Text": chunk,
                "Entity": term,
                "Start": match.start(),
                "End": match.end(),
                "Label": f"CLUSTER_{cluster}"
            })

df_ner = pd.DataFrame(ner_data)

# 8. 결과 출력 예시 (원하는 경우 저장도 가능)
print("\n[유사한 단어쌍]")
print(df_similar)

print("\n[클러스터링 결과]")
print(cluster_df)

print("\n[NER 학습용 구조]")
print(df_ner.head())

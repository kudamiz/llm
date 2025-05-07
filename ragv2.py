import requests import pandas as pd import numpy as np from sklearn.metrics.pairwise import cosine_similarity from sklearn.cluster import AgglomerativeClustering

1. OpenSearch에서 미리 임베딩된 벡터 불러오기

host = "https://your-opensearch-domain" index_name = "your-index-name" auth = ("your-username", "your-password")

2. 쿼리를 통해 임베딩된 청크 벡터 가져오기 (이미 저장된 벡터)

def get_embeddings_from_opensearch(): url = f"{host}/{index_name}/_search" query = { "size": 1000,  # 가져올 벡터 수 "query": { "match_all": {} }, "_source": ["text", "passage_chunk_embedding"] } response = requests.post(url, json=query, auth=auth).json() docs = response['hits']['hits']

texts = [doc['_source']['text'] for doc in docs]
embeddings = [doc['_source']['passage_chunk_embedding'] for doc in docs]
return texts, np.array(embeddings)

3. 임베딩 및 텍스트 로드

texts, vectors = get_embeddings_from_opensearch()

4. 유사도 계산 (cosine similarity)

similarity_matrix = cosine_similarity(vectors)

5. 유사한 단어쌍 추출

similar_pairs = [] for i in range(len(texts)): for j in range(i + 1, len(texts)): sim = similarity_matrix[i][j] if sim > 0.8:  # 유사도 기준 조정 가능 similar_pairs.append({ "Text A": texts[i], "Text B": texts[j], "Cosine Similarity": round(sim, 4) })

6. 클러스터링 (계층적 군집화)

distance_matrix = 1 - similarity_matrix clustering = AgglomerativeClustering( affinity='precomputed', linkage='average', distance_threshold=0.25, n_clusters=None ) labels = clustering.fit_predict(distance_matrix)

7. 결과 정리

cluster_df = pd.DataFrame({ "Text": texts, "Cluster ID": labels }).sort_values("Cluster ID")

8. 출력

print("[유사한 단어쌍]") print(pd.DataFrame(similar_pairs))

print("[클러스터링 결과]") print(cluster_df)


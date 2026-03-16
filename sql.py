-- 1. 대화 세션 테이블
CREATE TABLE IF NOT EXISTS Sessions (
    session_id TEXT PRIMARY KEY,       -- 고유 세션 ID (UUID)
    user_id TEXT,                      -- 사용자 식별자
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    status TEXT DEFAULT 'ACTIVE',      -- ACTIVE, CLOSED 등
    summary_q TEXT,                    -- 추후 LLM이 요약할 질문
    summary_a TEXT                     -- 추후 LLM이 요약할 답변
);

-- 2. 개별 메시지 테이블
CREATE TABLE IF NOT EXISTS Messages (
    message_id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT,                   -- 소속된 세션 ID (FK)
    role TEXT,                         -- 'user' 또는 'assistant'
    content TEXT,                      -- 질문 또는 답변 내용
    context_used TEXT,                 -- 챗봇이 참고한 RAG 문서 이름/ID (JSON 등)
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(session_id) REFERENCES Sessions(session_id)
);

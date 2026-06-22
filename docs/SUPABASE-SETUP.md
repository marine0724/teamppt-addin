# Supabase 셋업 인계 체크리스트 (A-1b/c 선행 작업)

> 사용자(관리자)가 Supabase 대시보드에서 1회 수행. 끝나면 키 3개를 로컬 파일에 넣는다(커밋 금지).
> 코드 쪽 상세는 [A-1b/c plan](superpowers/plans/2026-06-22-a1bc-supabase-ingest-upload.md) Task 0~1.

---

## 0. 프로젝트 생성
- [ ] supabase.com → New project. 이름 자유(예: `teamppt`), 비밀번호·리전(가까운 곳, 예: Tokyo) 설정.
- [ ] 생성 완료까지 1~2분 대기.

## 1. pgvector 확장 켜기
- [ ] 좌측 **SQL Editor** → New query → 아래 실행:
```sql
create extension if not exists vector;
```

## 2. assets 테이블 + 벡터 인덱스
- [ ] SQL Editor에서 실행:
```sql
create table assets (
  id          uuid primary key default gen_random_uuid(),
  file        text not null,          -- Storage 경로: "pptx/표지_01.pptx"
  thumb       text not null,          -- Storage 경로: "thumb/표지_01.png"
  name        text not null,
  category    text not null,          -- 섹션명
  kind        text not null,          -- layout / component
  scope       text,
  tags        text[],
  use_when    text,
  content_fit text[],
  metadata    jsonb,                  -- colors, fonts, slots + 미래 속성
  embed_text  text,
  embedding   vector(768),            -- gemini text-embedding-004
  source_deck text,
  created_at  timestamptz default now()
);
create index on assets using hnsw (embedding vector_cosine_ops);
```

## 3. 벡터 검색 함수(RPC)
- [ ] SQL Editor에서 실행:
```sql
create or replace function match_assets(query_embedding vector(768), match_count int)
returns table (
  id uuid, file text, thumb text, name text, category text, kind text,
  scope text, tags text[], use_when text, content_fit text[], metadata jsonb,
  similarity float
) language sql stable as $$
  select a.id, a.file, a.thumb, a.name, a.category, a.kind,
         a.scope, a.tags, a.use_when, a.content_fit, a.metadata,
         1 - (a.embedding <=> query_embedding) as similarity
  from assets a
  order by a.embedding <=> query_embedding
  limit match_count;
$$;
```

## 4. 보안 정책(RLS) — 일반 사용자는 읽기만
- [ ] SQL Editor에서 실행:
```sql
alter table assets enable row level security;
create policy "anon_read_assets" on assets for select to anon using (true);
-- 쓰기 정책 없음 → service_role 키(RLS 우회)만 추가/수정/삭제 가능
```

## 5. Storage 버킷 2개 (public read)
- [ ] 좌측 **Storage** → New bucket → 이름 `pptx`, **Public bucket 체크** → 생성.
- [ ] 같은 방식으로 `thumb` 버킷 생성(Public 체크).
  - (에셋·썸네일은 비밀이 아니라 public 읽기로 단순화. 추후 서명 URL로 좁힐 수 있음.)

## 6. 키/URL 확보 (커밋 절대 금지)
- [ ] 좌측 **Project Settings → API**에서 복사:
  - **Project URL** (예: `https://abcd1234.supabase.co`)
  - **anon public** 키
  - **service_role** 키 (비밀! 관리자 PC에만)

## 7. 로컬 파일에 넣기 (이 둘 다 gitignore — 커밋 안 됨)
- [ ] `src/TeampptAddin/Assets/api-keys.json` — 기존 `gemini` 값은 **그대로 두고** 두 줄만 추가:
```json
{
  "gemini": "<기존 값 유지 — 손대지 말 것>",
  "supabaseUrl": "<Project URL>",
  "supabaseAnonKey": "<anon public 키>"
}
```
- [ ] `%LOCALAPPDATA%\TeampptAddin\admin.json` 새로 생성(이 파일이 있어야 인제스트 버튼이 켜짐 = 관리자 PC):
```json
{
  "supabaseServiceKey": "<service_role 키>",
  "geminiKey": "<api-keys.json의 gemini와 동일 값>"
}
```

## 완료 확인
- [ ] Table Editor에 `assets` 테이블 보임.
- [ ] Database → Functions에 `match_assets` 보임.
- [ ] Storage에 `pptx`, `thumb` 버킷 보임(둘 다 Public).
- [ ] 위 두 로컬 파일 작성됨(커밋 안 함).

→ 여기까지 되면 A-1b/c(적재) / A-1d(읽기) 개발 실행 가능. **A-1a(이해 어댑터)는 이 셋업 없이도 먼저 진행 가능**(Gemini 키만 필요).

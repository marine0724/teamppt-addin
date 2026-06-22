# Phase F — Supabase 벡터 저장 + 대화 메모리 최적화 설계 (진행중)

> 작성: 2026-06-19 · 상태: **아키텍처 확정, 세부 7개 미결 / 코드 미착수**
> 선행: 구조화 출력(responseSchema) = [plans/2026-06-19-structured-output-responseschema.md](../plans/2026-06-19-structured-output-responseschema.md) (코드 완료, PowerPoint 수동검증 대기)
> 상위: [AI 리디자인 마스터](2026-06-19-teamppt-ai-redesign-design.md)

---

## 0. 이 문서의 목적
새 세션에서 이 작업을 **끊김 없이 재개**하기 위한 결정·근거·미결 항목 보존. 대화에서 확정한 내용을 잃지 않기 위함.

## 1. 풀려는 문제 (대화 출발점)
1. **에셋 추천이 느림** — 전체 카탈로그(JSON)를 매 프롬프트에 통째로 싣고 thinkingBudget까지 써서 응답 지연. 에셋이 늘면 토큰 폭발.
2. **대화 기록 토큰 증가** — 턴이 쌓이면 PPT 완성 전 토큰이 너무 커짐.
3. 데이터가 핵심. "결국 데이터 저장·출력 싸움."

## 2. 환경 제약 (결정에 직접 영향)
- **COM Shared Add-in** (`IDTExtensibility2`), **.NET Framework 4.8**, PowerPoint 프로세스에 인프로세스 로드.
- 의존성 최소(Newtonsoft.Json + Office Interop). **무거운 프레임워크 금지** — 로드한 DLL이 Office와 AppDomain 공유 → 버전 충돌 위험.
- HTTP는 `System.Net.Http.HttpClient` 직접 사용 (기존 [GeminiAiService](../../../src/TeampptAddin/Services/GeminiAiService.cs) 패턴).

## 3. 확정된 아키텍처 결정

### 3.1 벡터 저장 = Supabase (Postgres + pgvector)
- REST/RPC를 HttpClient로 **직접 호출**. `supabase-csharp` SDK 통째 도입 금지(net48/COM 의존성 지옥).
- 벡터 검색은 Postgres 함수(RPC) 하나 만들어 HTTP POST.
- **진짜 이득은 속도가 아니라 중앙 집중식 에셋 관리** — 행 추가하면 애드인 재배포 없이 전원 반영.

### 3.2 임베딩 = Gemini `text-embedding-004` (768차원)
- 단일 벤더(API 키 하나, 기존 HTTP 플러밍 재사용).
- **무엇을 임베딩하나가 정확도의 핵심**: 에셋당 "의미 문서" = name + use_when + content_fit + tags + **예시 의도 문장 여러 개**("투자 유치 IR 표지", "회사 소개 첫 장" ...).

### 3.3 추천 흐름
사용자 의도 → 임베딩 → pgvector로 **상위 N개만** 검색 → 그 N개만 Flash에 전달(전체 카탈로그 X). LLM 역할이 "검색"→"좋은 후보 중 선택·설명"으로 바뀜 → 빠르고 정확.

### 3.4 대화 메모리 = 구조화된 상태 객체 + 작은 윈도우 (프레임워크 없이 직접)
- **Semantic Kernel / LangChain 둘 다 배제** (net48/COM 부적합. SK는 일반론으론 C# 1순위지만 이 환경엔 안 맞음. LangChain.NET은 미성숙).
- **윈도우를 늘리는 게 아님 — 줄임(2~3턴).** 장기 맥락은 transcript가 아니라 **구조화 상태**로:
  ```json
  { "concept": "...", "palette": "navy-formal", "font": "Pretendard",
    "selectedAssets": ["header_1","header_3"], "openQuestions": ["..."] }
  ```
- 이 상태는 **이미 파싱하는 [AiRecommendation](../../../src/TeampptAddin/Models/AiRecommendation.cs)에서 결정적으로 갱신** → 추가 LLM 호출 0, 비용 0, 토큰 평탄.
- (대안 B = LLM 롤링 요약, 추가 호출 필요·비결정적. 우선순위 낮음.)

### 3.5 에셋 파일 저장: pptx-per-asset 유지 + 메타/바이너리 분리
- 한 에셋 = 한 pptx **유지** (pptx가 네이티브 충실도 포맷. JSON 변환 시 충실도 손실+막대한 비용).
- 효율은 "합치기"가 아니라 **저장 위치 분리**:
  | 무엇 | 어디 |
  |---|---|
  | 메타데이터 + 임베딩 (작음, 자주 쿼리) | Postgres 테이블 |
  | 실제 .pptx 바이너리 (큼, 가끔) | Supabase Storage(오브젝트) |
  | 다운로드한 pptx | 로컬 캐시 (재다운로드 X) |
  | 썸네일 (가볍게 먼저 표시) | 기존 ThumbnailGenerator |
- **지연 로딩**: 추천 시 썸네일만, .pptx는 삽입 순간에만 다운로드+캐시.

### 3.6 제안 스키마 (초안)
```sql
create table assets (
  id          uuid primary key default gen_random_uuid(),
  file        text not null,          -- header_1.pptx
  name        text not null,
  category    text,
  scope       text,                   -- deck / slide
  tags        text[],
  metadata    jsonb,                  -- colors, fonts, slots (구조 데이터)
  embed_text  text,                   -- 임베딩한 원문 (재현용)
  embedding   vector(768)             -- gemini text-embedding-004
);
```
원칙: **임베딩 텍스트(`embed_text`, 검색용)와 구조 메타데이터(`metadata`, 삽입용)를 분리.**

## 4. 미결 결정 7개 — ✅ **해소됨 (2026-06-22)**

> 7개 전부 확정: [vibe-designing-a-first §11](2026-06-22-vibe-designing-a-first-design.md). 요약: ①anon+RLS/admin.json ②캐시+번들 폴백 ③애드인 관리자 버튼 ④인제스트 경로 재사용+LLM 재생성 ⑤embedding-004(768)+hnsw ⑥top8·하드컷없음 ⑦세션 메모리. + 유료화 진화경로(Auth/role/Edge Function) & `IAccessPolicy` seam.

(원문 — 결정 근거 기록용)
1. **인증** — anon 키를 api-keys.json에 + RLS 읽기전용 정책? (권장: 예)
2. **오프라인/네트워크 실패** — Supabase 불통 시 번들 로컬 JSON 폴백 / 마지막 카탈로그 캐시?
3. **임베딩 생성 파이프라인** — 에셋 추가 시 누가 임베딩 생성·업로드? 별도 관리자 스크립트? (사실상 별도 워크플로우)
4. **기존 7개 에셋 마이그레이션** — assets.json → Supabase 행 + 임베딩 (일회성 스크립트)
5. **임베딩 모델/차원/인덱스** — text-embedding-004(768) + pgvector hnsw 확정?
6. **top-N + 유사도 임계값** — 후보 개수, 무관 에셋 컷 점수
7. **대화 상태 객체** — 필드 확정, 세션 메모리 vs Supabase deck별 저장

## 5. 현재 상태 / 블로커
- ✅ **구조화 출력(A) 코드 완료** — `BuildResponseSchema()`, 프롬프트 형식블록 제거, responseSchema 연결. 단위테스트 19개 PASS. PowerPoint 수동검증만 남음.
- ✅ **503/429/500 재시도(백오프)** 추가.
- ✅ 유출 키 제거 + [[feedback-no-keys-in-docs]] 기록.
- 🚫 **블로커: Gemini API 키** — AI Studio에서 받은 키가 `AQ.`로 시작하는 53자(임시/OAuth 토큰, 만료됨→400 API_KEY_INVALID). **정상 키는 `AIza`로 시작 ~39자.** 올바른 "API key"를 받아 [api-keys.json](../../../src/TeampptAddin/Assets/api-keys.json)에 넣고 **관리자 재빌드** 필요. (PowerPoint 검증과 B 진행 모두 이것 해결 후 가능.)

## 6. 다음 액션 (재개 시)
1. (사용자) 올바른 `AIza...` 키 확보 → api-keys.json 교체 → 관리자 빌드 → 구조화 출력 PowerPoint 검증.
2. 미결 7개 결정 (각 추천안+이유).
3. 확정되면 Phase F 코드 계획(plans/) 작성 후 실행.

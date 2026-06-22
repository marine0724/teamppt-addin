# 🗺️ TEAMPPT 개발 진행 보드

> 이 파일 하나만 열어두면 "지금 어디서 뭘 하는지" 보입니다. Claude가 매 세션 함께 유지.
> **기록용 아카이브가 아니라 "지금 여기" 작업 보드.** 끝난 잎(Task)은 지우고 교체, 숲·나무 단위는 끝날 때까지 유지. (규칙: CLAUDE.md)
> 계층: **나라 > 대지 > 숲 > 나무 > 잎** (2026-06-22 재정립)
> 최종 갱신: 2026-06-22 · 현재 작업: **A-1b/c Supabase 적재경로 ✅ 코드 완료(42테스트 GREEN) → 수동 검증(마이그레이션) 대기.**

---

## 🌍 나라 — 제품 정체성: **바이브 디자이닝**

> 커서가 "바이브 코딩"이듯, TEAMPPT는 PPT 디자인을 그렇게. **전문 역량 → 노멀 역량이 되어도 같은 결과물.**
> 설계: [vibe-designing & A-first](docs/superpowers/specs/2026-06-22-vibe-designing-a-first-design.md) · 발표 자산: [docs/PITCH.md](docs/PITCH.md)

## 🏔️ 대지 — 고객 사용 루트 (A/B/C)

```
[A 에셋 조립·커서식]      [B 리디자인·위임식]       [C 기획+에셋]
   ▲ 지금 여기                 (다음)                  (추후)
 곁에서 추천·삽입         통째로 한방 변환          백지부터 구성
 "어떤 에셋?"             "어떤 컨셉?"              "어떤 구성?"
```

> A·B의 유일한 차이 = "전체를 한 번에 갈아치우는 실행엔진"(=D)의 유무. 둘 다 같은 '읽기 엔진' 공유.
> **지금은 토대인 A를 온전하게.** A만 써도 감동이 나와야 B·C가 이상적으로 얹힌다. (D 한방엔진은 B 루트로 분리·후순위)

---

## 🌲 숲 — 지금 로드맵: **"온전한 A"** (데이터 → 공유 → 추천·진단 UX)

```
[A-1 데이터 토대]──▶[A-2 실시간 공유엔진]──▶[A-3 추천·진단 UX]      ┊  (분리) D 한방 실행엔진
      ⬜ 다음              ⬜                      ⬜               ┊   → B 루트·후순위
  인제스트→Supabase     슬라이드 읽기+공유표시     에셋추천+일관성진단
  벡터추천 (해자)        [슬라이드 N 공유중]
```

목표: **A만 써도 감동**. "수작업 json"이 아니라 "LLM 인제스트→Supabase 벡터추천"이 도는 토대 = 투자자 비전이 하늘과 땅.

- **A-1(다음 깃발):** 인제스트 LLM 이해 어댑터(**타입별 판단·분류**) → 임베딩 → Supabase 업로드 + LLM이 Supabase에서 불러와 추천(벡터검색 읽기경로). **대상 = 로컬 7개(헤더) + 인제스트 실험 35개(15섹션, 레이아웃~그래프).** R&D가 에셋 계속 붓는 레일. (기존 B 인제스트 나머지 + F 벡터 흡수)
- **A-2:** 슬라이드 리더(현재/전체 온디맨드) + `[슬라이드 N 공유중]` indicator. (기존 C 확장)
- **A-3:** ② 텍스트 질의 추천("장점 3가지 추천")=벡터DB만, *먼저·완성도 우선* → ① 현재슬라이드 추천(2~3개)=공유엔진+벡터 + 전체덱 *일관성* 진단(헤더·색 컨셉·구성). 넓은 진단은 C로 미룸.
- 5대 결정·세 루트 전부 유지. 인제스트 로컬코어 ✅·Phase A/E ✅.

---

## 🌳 나무 — 지금 plan: **A-1 데이터 토대** (3개 plan, 순차)

> 결정 7개 확정 + 유료화 진화경로 spec §11. 실행계획 작성 완료(아래 3개), **착수 대기**.

| plan | 무엇 | 선행 | 핵심 깃발 |
|---|---|---|---|
| [A-1a LLM 이해 어댑터](docs/superpowers/plans/2026-06-22-a1a-llm-understanding-adapter.md) ✅ | PNG+섹션명 → Gemini 멀티모달 이해 → 구조화 레코드(kind 분류)+embed_text | 로컬 인제스트 코어 ✅ | **완료** (5커밋, 37테스트 GREEN) |
| [A-1b/c Supabase 적재경로](docs/superpowers/plans/2026-06-22-a1bc-supabase-ingest-upload.md) ✅ 코드 | 인프라(테이블/RPC/RLS/버킷)+임베딩+업로드+`IAccessPolicy` 게이트 | A-1a + Supabase 프로젝트 | 코드 완료(7커밋, 42테스트 GREEN), 수동 검증(마이그레이션) 대기 |
| [A-1d 벡터검색 읽기경로](docs/superpowers/plans/2026-06-22-a1d-vector-read-path.md) | 질의→임베딩→match_assets(top8)→Gemini 선택→AI탭 추천 + 오프라인 폴백 | A-1b/c | 텍스트 질의 추천 동작(A 첫 시연) |

---

## 🍃 잎 — Task 현황

> **A-1b/c ✅ 코드 완료** (2026-06-22): AdminCredentials(admin.json 로더) · IAccessPolicy(관리자 게이트 seam) · AssetRowBuilder(Postgres row+pgvector 문자열) · EmbeddingService(text-embedding-004 768) · SupabaseClient(REST insert/RPC/Storage) · AssetIngestUploader(오케스트레이터). 신규 5 단위테스트, 전체 42개 GREEN.
> **현재 위치:** A-1b/c 코드 완료. **수동 검증(마이그레이션) 대기** — admin.json+split 산출물 준비 후 AssetIngestUploader로 7+35 에셋 Supabase 적재 필요.
> ⚠️ **수동 검증 미완:** EmbeddingService·SupabaseClient·AssetIngestUploader의 실제 API 호출 + Supabase 적재 검증은 admin.json 존재 환경에서 수동 확인 필요.
> **다음 = A-1d 벡터검색 읽기경로** (사용자 승인 후 착수).
> **빌드:** `MSBuild TeampptAddin.csproj /t:Build /p:Configuration=Debug /p:Platform=AnyCPU /p:RegisterForComInterop=false`. COM 재등록 불필요.

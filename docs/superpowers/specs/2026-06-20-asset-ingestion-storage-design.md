# 에셋 인제스트 & 저장 규약 설계 (Phase B 본체)

> 작성: 2026-06-20 · 상태: **설계 확정, 코드 미착수**
> 선행: [Phase F Supabase 벡터 설계](2026-06-19-phase-f-supabase-vector-design.md) (저장/검색 인프라)
> 상위: [AI 리디자인 마스터](2026-06-19-teamppt-ai-redesign-design.md)

---

## 0. 이 문서의 목적

디자이너가 만든 **묶음 pptx**(레이아웃 여러 개가 한 파일)를 받아 → 개별 에셋으로 쪼개고 → LLM이 이해하고 → Supabase에 저장하는 **인제스트 파이프라인과 저장 규약**을 확정한다. 이게 정해져야 Supabase 저장·출력을 한 번에 설계할 수 있다(로드맵 Phase B = 인제스트 자동화 = 해자).

## 1. 출발점 (대화에서 드러난 문제)

- 현재: 에셋 1개 = pptx 1개(header_1~7), `assets.json` 메타데이터를 사람이 수기 작성.
- 새 현실: 디자이너는 **1개 pptx에 레이아웃 여러 개**를 넣고, **PowerPoint 섹션(Section)**으로 분류해둠 (예: `레이아웃(표지)`, `레이아웃(목차)`).
- 배달 포맷은 앞으로 바뀔 수 있음 → 물리적 규약에 과의존하면 깨짐. **거친 분류는 섹션(결정적), 세밀한 이해는 LLM**으로 역할 분담.

## 2. 확정된 핵심 결정

### 2.1 에셋의 단위 = 슬라이드 1장
- **섹션 = category** (섹션명을 그대로 category로). 코드가 `SectionProperties`로 결정적으로 읽음.
- **섹션 안 슬라이드 1장 = 에셋 1개** (그 카테고리의 한 variant). 한 에셋이 여러 슬라이드에 걸치는 경우는 없음(확정).

### 2.2 데이터 흐름 (split 먼저, LLM은 이해만)
```
디자이너 묶음 pptx (PowerPoint 섹션으로 분류됨)
   │
   ① 코드: SectionProperties 읽기 → (섹션명=category, 슬라이드 범위)
   │
   ② 코드: 슬라이드 1장 = 에셋 1개 단위로 split → 개별 pptx N개 + PNG N개
   │
   ③ LLM(Gemini 멀티모달): 각 PNG + 섹션명 힌트 → 구조화 에셋 레코드
   │     (name, use_when, content_fit, tags, 예시의도문장들, slots, colors, fonts)
   │
   ④ 코드: embed_text 조립 → Gemini text-embedding-004 → 768벡터
   │
   ⑤ 코드: Supabase 업로드 (메타+벡터는 Postgres, pptx+PNG는 Storage)
```
- 사용자가 제안한 순서(`LLM 판단 → split → 저장`)에서 **split을 먼저** 하도록 조정: 섹션이 이미 분류를 결정적으로 주므로 LLM이 거친 분류를 다시 할 필요 없음. LLM은 각 조각의 *이해*(의도·슬롯·색·폰트)에만 집중.

### 2.3 디자이너 의도 = LLM 시각 추론
- 디자이너는 **섹션 분류만** 한다. 슬라이드별 세부 의도("3단 강점, 임팩트 줄 때")는 **적지 않음**.
- Gemini 멀티모달이 슬라이드 PNG + 섹션명 힌트를 보고 `use_when`·`content_fit`·예시의도문장을 생성. → 디자이너 부담 0, "이해는 인제스트 타임에 펑펑" 철학과 일치.

### 2.4 슬롯도 LLM/휴리스틱 추론
- **슬롯(slot)** = 에셋 안 "사용자 글이 들어갈 이름 붙은 빈자리"(title/subtitle/body/image…). Route B 리디자인에서 사용자 초안 텍스트를 매핑할 대상.
- 디자이너가 shape 이름(`slot.title`)을 직접 달지 **않음**. 인제스트 때 LLM이 위치·크기·폰트 단서로 "이 텍스트박스는 제목 자리"라고 추론.
- 트레이드오프: 디자이너 부담 0 ↔ LLM이 가끔 틀릴 수 있음(인제스트 타임 토큰으로 정확도 보강).

### 2.5 인제스트 실행 위치 = 애드인 안 관리자 모드
- 별도 콘솔 도구 대신 **애드인 안에 숨은 "에셋 인제스트" 버튼**. PowerPoint가 이미 떠 있으니 섹션 읽기·슬라이드 split·PNG 렌더를 전부 기존 Interop으로 처리.
- 일반 사용자에겐 버튼이 안 보임.

### 2.6 관리자 게이트 = 로컬 자격증명 파일 (로그인 서버 불필요)
- 인제스트는 Supabase **쓰기**가 필요 → **service-role 키**(쓰기) 필요. 일반 사용자는 **anon 키**(읽기 전용)만.
- **"관리자 자격증명 파일이 로컬에 있느냐"가 곧 관리자 게이트**:
  ```
  %LOCALAPPDATA%\TeampptAddin\admin.json   ← 관리자 PC에만 존재
     { "supabaseServiceKey": "...", "geminiKey": "..." }
  ```
  - 있으면 → 인제스트 버튼 노출 + 쓰기 키 사용 (관리자 모드)
  - 없으면 → 일반 사용자 모드 (anon 읽기, 버튼 숨김)
- 파일은 **.gitignore + 배포본 미포함**. 게이트와 자격증명을 한 파일로 통합.
- **권한 범위**: 이 파일 소유자는 추가뿐 아니라 수정·삭제까지 전부 가능(service-role은 RLS 우회 마스터 키).
- **완화책**: ① 유출 시 Supabase에서 키 회전(rotate)하면 즉시 무력화 → 복구 가능한 사고. ② 디자이너 여럿이 인제스트하게 되면 service-role 통째 대신 "assets 쓰기만, 삭제 불가" 좁은 전용 DB 역할 키로 교체. (초기엔 service-role로 단순하게.)

## 3. 저장 스키마 (Supabase)

"자주 쿼리하는 작은 것"과 "가끔 받는 큰 것"을 분리.

### 3.1 Postgres 테이블 `assets` (메타+벡터, 작음, 매번 쿼리)
```sql
create table assets (
  id          uuid primary key default gen_random_uuid(),
  file        text not null,        -- Storage 경로: "pptx/표지_01.pptx"
  thumb       text not null,        -- Storage 경로: "thumb/표지_01.png"
  name        text not null,        -- LLM 생성: "우측정렬 연도강조 표지"
  category    text not null,        -- 섹션명 그대로: "표지"
  scope       text,                 -- deck / slide
  tags        text[],               -- LLM 생성
  use_when    text,                 -- LLM 생성
  content_fit text[],               -- LLM 생성
  metadata    jsonb,                -- colors, fonts, slots + 미래 실험 속성 (탈출구)
  embed_text  text,                 -- 임베딩에 넣은 원문 (재현/디버깅)
  embedding   vector(768),          -- gemini text-embedding-004
  source_deck text,                 -- 원본 묶음 pptx 이름 (추적)
  created_at  timestamptz default now()
);
create index on assets using hnsw (embedding vector_cosine_ops);
```
- **id/파일명 규칙**: 사람이 읽는 이름 `{category}_{순번}`(예: `표지_01`) + uuid 병행.
- **원칙**: `embed_text`(검색용 의미문서) ↔ `metadata`(삽입용 구조데이터) 분리. 검색은 벡터로, 삽입은 metadata로.

### 3.2 Supabase Storage (바이너리, 큼, 가끔)
| 버킷 | 내용 | 언제 받나 |
|---|---|---|
| `pptx/` | split된 개별 에셋 pptx (1슬라이드 1파일) | **삽입하는 순간에만** 다운로드 |
| `thumb/` | 슬라이드 PNG 썸네일 | 추천 카드 표시할 때 (가벼움, 먼저) |

### 3.3 로컬 캐시 (재다운로드 방지)
```
%LOCALAPPDATA%\TeampptAddin\cache\
   pptx\표지_01.pptx     ← 한 번 받으면 재사용
   thumb\표지_01.png
```

### 3.4 추천→삽입 종합 흐름
```
사용자 의도 → 임베딩 → Postgres RPC(벡터검색) → 상위 N개 행
   → 카드에 thumb만 표시 (가벼움)
   → 사용자가 하나 클릭 → 그때 pptx/ 다운로드 → 캐시 → 슬라이드에 삽입
```
Postgres엔 "고르는 데 필요한 작은 정보 전부", Storage엔 "고른 뒤에만 필요한 무거운 pptx" → 추천은 항상 빠름.

## 4. 유동성 전략 (최우선 — 글로벌/현지화 대비)

목표는 글로벌. 나라별 현지화 에셋이 추가될 수 있음. **지금 컬럼을 늘리지 않되, 늘리기 쉬운 형태**로 둔다.

1. **`metadata jsonb`를 탈출구로** — 정체가 불확실/실험적 속성은 전부 jsonb. 키 추가는 마이그레이션 불필요.
2. **타입 컬럼은 거르고/정렬/검색하는 것만** — category, tags, embedding 등 안정적 골격만. 나중 컬럼 추가도 `ALTER TABLE ADD COLUMN`으로 기존 행 안 깨고 가능.
3. **현지화 대비, 지금은 비움** — 나라별 에셋이 오면 그때 `locale` 컬럼 + RPC에 `where locale = ?` 한 줄. 지금은 아무것도 안 하되 **RPC·코드가 "단일 국가"를 가정하지 않게** 둠.
4. **코드는 모르는 필드를 버리지 않음** — `[JsonExtensionData]` 패턴(이미 `AssetSlot`에 사용), `schemaVersion` 유지로 포맷 변화 추적.

원칙 한 줄: **"고정 = 검색에 필요한 최소 골격, 나머지 전부 jsonb."**

## 4.5 PNG 렌더 해상도 정책 (전역 원칙 — 토큰 효율)

LLM이 화면을 PNG로 읽는 **모든 경로**에 적용. 원칙: **정확한 미션 수행이 되는 최대한을 쓰되, 그 중 최소 화질.**

Gemini 이미지 과금은 픽셀당이 아니라 **타일 단위**: 양변 ≤384px = 1타일(258토큰), 그보다 크면 768×768 타일로 쪼개 타일당 258토큰.

| 렌더 (16:9) | 타일 | 토큰 | 비고 |
|---|---|---|---|
| 384×216 | 1 | 258 | 잔글씨 뭉개짐 |
| **768×432** | **1** | **258** | **같은 토큰, 훨씬 선명 — 최저 구간 내 최대 화질** |
| 1280×720 | 2 | ~516 | 과사용(2배) |

- **기본 = 긴 변 768px** (16:9면 768×432). 1타일/258토큰.
- 인제스트에선 슬라이드 **텍스트는 pptx XML(Interop)에서 정확히 읽음** → PNG는 OCR용이 아니라 레이아웃·스타일·색감용. 768px면 충분.
- **상향은 미션이 요구할 때만** (XML 없이 렌더된 잔글씨를 꼭 읽어야 하는 등) → 1536px(2×2 타일) 등 단계적.
- **단일 설정 상수**(`LlmImageLongEdgePx = 768`)로 모든 화면읽기 경로가 공유 → 한 곳만 조정. 토큰을 아껴 정말 필요한 곳에 더 쓴다.

### Claude(인제스트 본채택 모델) 비전 토큰 — 하이브리드 대비

인제스트는 향후 Claude(Opus 4.8), 사용자 출력은 저렴한 Gemini로 가는 하이브리드. Claude의 비전 과금은 **타일이 아니라 면적 비례**라 구조가 다름:
- 공식: 토큰 ≈ (가로 × 세로) / 750. **계단 없음, 선형.** Gemini식 "공짜 구간" 없음.
- 최대 해상도: 긴 변 2576px(~3.75MP)까지. 그 이상 자동 다운스케일. 이미지당 상한 ~4784토큰.

| 렌더(16:9) | Gemini | Claude(W×H/750) |
|---|---|---|
| 768×432 | 258(1타일) | ~442 |
| 1280×720 | ~516(2타일) | ~1229 |

- **768px 단일 상수가 양쪽에서 합리적**: Gemini는 최저 타일 구간, Claude는 ~442토큰(텍스트는 XML에서 읽으므로 충분).
- 모델 어댑터가 같은 px를 받아 각자 인코딩(Claude=면적, Gemini=타일)만 처리. Claude는 계단이 없어 미션이 디테일 요구 시 768→2576px로 부드럽게 상향 가능.

## 5. 디자이너 규약 (인제스트 성립 최소 약속)

| # | 규약 | 이유 |
|---|---|---|
| 1 | 레이아웃을 **PowerPoint 섹션**으로 묶는다 | 섹션명=category, 코드가 결정적으로 읽음 |
| 2 | 섹션명은 사람이 읽을 **카테고리**로 (`표지`, `목차`, `간지`, `3단강점`…) | 그대로 category + LLM 힌트 |
| 3 | 1 섹션 = 같은 종류 변형들, **1 슬라이드 = 1 에셋** | split 단위 고정 |
| 4 | 텍스트 자리는 **실제 텍스트박스**로 (이미지 위 글자 X) | 슬롯 추론·재디자인 가능. 단 shape 이름은 안 달아도 됨 |
| 5 | 의도 설명은 **안 적어도 됨** | LLM 시각 추론 |

## 6. 기존 코드와의 연결

- `AssetSchema.cs`(colors/fonts/slots 모델) → `metadata jsonb` 안 구조와 일치하게 재사용.
- `CatalogEntry.cs` → Supabase 행에서 카드 표시용으로 매핑.
- `ThumbnailGenerator.cs` → 인제스트 ② 단계 PNG 렌더에 재사용.
- `GeminiAiService.cs`의 HTTP 패턴 → 임베딩 호출·Supabase REST 호출에 재사용.
- 설계 원칙 유지: Core/Connect.cs/Globals.cs 수정 금지, Newtonsoft.Json, IAiService 고정.

## 7. 미결/실험 영역 (지금 정하지 않음 — 유동성으로 흡수)

- 추가 컬럼(저작권/디자이너/버전/활성플래그/locale) → 필요해지면 jsonb 또는 ADD COLUMN.
- top-N 개수·유사도 임계값 → Phase F 미결 6번에서 실험으로 결정.
- 임베딩 텍스트(`embed_text`) 조립 공식 → 인제스트 구현 시 튜닝.
- 오프라인 폴백·기존 7개 에셋 마이그레이션 → Phase F 미결 2·4번.

## 8. 다음 액션

1. 이 spec 사용자 리뷰.
2. writing-plans로 인제스트 파이프라인 구현 계획 작성.
3. 순서: 섹션 읽기·split(코드) → PNG 렌더 재사용 → LLM 이해(responseSchema 확장) → 임베딩 → Supabase 업로드 → 관리자 게이트 → (별도) 추천·삽입 읽기 경로.

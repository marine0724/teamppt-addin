# kind 3분류 (slide/layout/component) — 설계

날짜: 2026-06-23
관련: [인제스트 파이프라인](../../../CLAUDE.md), `UnderstandingSchema.cs`

## 목적

인제스트 시 에셋의 `kind`를 현재 2분류(`layout`/`component`)에서 3분류(`slide`/`layout`/`component`)로 확장한다. 표지처럼 슬라이드 통째가 하나의 단위인 에셋을 별도로 분류하기 위함.

## 배경

- 현재 `kind`는 Gemini 멀티모달 이해(`AssetUnderstandingService`)가 슬라이드 이미지 1장 + 섹션명 힌트로 판정한다.
- `kind`는 **삽입 동작을 분기시키지 않는다.** DB `assets.kind` 컬럼 적재 + AssetPanel 진행 표시용 메타데이터/검색 라벨일 뿐이다.
- 그림삽입 기능은 슬라이드를 통째로 들고 올 수 있어, 표지 같은 "통째 단위" 에셋을 별도 종류로 구분하면 분류가 자연스러워진다.

## 규약 (LLM 판정 규칙)

판정 우선순위 순서대로:

1. **`slide`** — 발표를 여닫는 **표지 슬라이드**(오픈표지·엔드표지). 페이지 통째가 하나의 완결 단위. (미래에 통째로 삽입할 대상.)
2. **`layout`** — 텍스트·콘텐츠 슬롯을 채워 재사용하는 본문 페이지 틀 (예: 3단 가로 레이아웃).
3. **`component`** — 틀 위에 얹는 부품 (그래프·표·다이어그램 등).

규칙: 오픈/엔드 표지에 적합한 통째 슬라이드면 `slide` → 채워 쓰는 본문 페이지 틀이면 `layout` → 부품이면 `component`.

## 범위 (이번 작업 = 분류 라벨만)

slide의 "슬라이드 통째 삽입" 동작은 디자이너와 추후 대대적으로 설계한다. **이번 작업은 분류 체계/스키마/프롬프트 확장에 한정**하며 삽입 코드는 손대지 않는다.

### 변경

| 파일 | 변경 |
|---|---|
| `Services/UnderstandingSchema.cs` (enum) | `["layout","component"]` → `["slide","layout","component"]` |
| `Services/UnderstandingSchema.cs` (BuildSystemPrompt) | `kind` 판단 규칙 한 줄을 위 3분류 규약으로 교체 |
| `Tests/UnderstandingSchemaTest.cs` | `kind` enum 단언을 3개 값으로 갱신 |

### 손대지 않는 것

- 삽입 코드 (slide의 통째 삽입 동작은 보류).
- `UnderstandingParser` 기본값: 누락 시 `"component"` 유지.
- DB 스키마: `assets.kind`는 자유 text 컬럼 → 마이그레이션 불필요.
- `AssetSchemaMigrator` / `SupabaseAssetMapper` 기본값(`"component"`) 유지.
- AssetPanel 진행 표시: `kind` 문자열을 그대로 출력하므로 `slide`도 자동 표시됨.

## 메모 (작업 아님)

- **기존 데이터:** 기존 DB의 표지류 에셋은 `kind="layout"`으로 적재돼 있다. 새 규약을 반영하려면 **재인제스트**가 필요(파일 UNIQUE + upsert라 중복 없음). 이번 범위엔 미포함. 수동 검증 시 표지 번들을 재인제스트해 `slide`로 분류되는지 확인.
- **`scope` 컬럼과 이름 겹침:** 별도 `scope` 컬럼에도 `"slide"` 값이 존재(다른 축 — 적용 범위). `kind="slide"`와 문자열이 겹치지만 **컬럼이 달라 충돌 없음**. 코드/쿼리에서 두 축을 혼동하지 말 것.

## 검증

1. 단위 테스트: `UnderstandingSchemaTest`가 3개 enum 값을 단언, 통과.
2. 빌드: 관리자 권한 MSBuild → DLL 타임스탬프 갱신 + 로그 오류 0건 (CLAUDE.md 절차).
3. 수동(선택): 표지 포함 번들 재인제스트 → 표지가 `slide`로 분류되는지 진행 표시/DB에서 확인.

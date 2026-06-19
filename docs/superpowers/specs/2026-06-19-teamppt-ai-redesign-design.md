# TEAMPPT AI 리디자인 — 설계 문서 (Master Spec)

> 작성: 2026-06-19 (Opus) · 실행: Sonnet · 상태: 설계 확정, Phase A 착수 대기
> 관련 계획: [Phase A 실행계획](../plans/2026-06-19-phase-a-asset-schema.md)

---

## Extract (한눈에)

**무엇을 만드나:** PowerPoint Task Pane 안에서, 사용자가 만든 초안 슬라이드를 AI가 **한방에 리디자인**해주는 디자인 어시스턴트. 핵심은 "에셋을 얼마나 잘 저장했느냐"로 품질이 결정된다 — 그래서 **데이터 스키마**가 제품의 레버리지다.

**핵심 통찰 5개:**
1. **두 개의 토큰 예산** — *이해*는 인제스트 타임(에셋당 1회, 펑펑 써도 됨)에 몰아넣고, *매칭*은 런타임(매 요청, 극도로 인색)에서 싸구려로. → 사용자는 마음껏 쓰고 우리 마진은 산다.
2. **색·폰트를 "절대값"이 아니라 "역할(role)"로 저장** — `{role, value, locked}`. 원본 값은 보존(=as-designed 기본 컨셉, 의도 손실 0), 역할은 *선택적* 재테마를 가능케 함. → **적은 에셋이 많게 느껴진다**(에셋 1개 × 컨셉 N = 효과적 N개).
3. **슬롯(slot)** — 에셋이 `title/subtitle/body` 같은 *이름 붙은 자리*를 선언. 인제스트 = 슬롯 탐지, 리디자인 = 초안 텍스트를 슬롯에 매핑. 스키마·리더·리디자인을 잇는 다리.
4. **컨셉 레이어** — 덱 단위로 {팔레트(역할→hex) + 폰트(역할→family) + 스타일태그}를 1회 락. 컨셉이 카탈로그를 일관된 부분집합으로 *사전 필터* → 슬라이드별 조합이 구조적으로 일관됨 + 런타임 토큰 평탄.
5. **비파괴 리디자인** — 절대 초안 위에 덮지 않는다. `Slide.Duplicate()`로 복제 후 복제본만 손댐. 비포애프터는 PowerPoint 썸네일 레일에 N/N+1로 공짜로 나옴.

**누구에게:** 내용은 있는데 디자인을 못 하는 사람 — 발표 잦은 직장인·컨설턴트·대학(원)생. "초안은 있다, 예쁘게만 해달라."

**개발 순서:** Phase A(데이터 스키마) → E(Claude API 가볍게) → C(슬라이드 리더) → B(인제스트 자동화) → D(Route B 리디자인). **지금 착수 = Phase A.**

---

## 1. 제품 비전

사용자가 PowerPoint Task Pane(우측 패널) 안에서 AI와 함께 슬라이드를 디자인한다. 제품 가치의 정점은 **"이미 만든 초안을 한방에 리디자인"**이다. 콘텐츠가 이미 존재하므로 AI가 내용을 지어낼 필요가 없고(할루시네이션 0), "내용 먼저 → 디자인 절망"이라는 실제 워크플로와 정확히 맞물린다.

## 2. 세 가지 루트 (제품 스테이징 = 진입 분기)

시작 시 "지금 어떤 상황이세요?"로 루트를 고르게 한다. 세 루트는 같은 엔진(카탈로그·슬롯·컨셉·역할치환)을 공유하고 입구만 다르다.

| 루트 | 모드 | 상황 | 단계 |
|------|------|------|------|
| **A 조립** | 에셋을 하나씩 추천 → 드래그/클릭 삽입 | "처음부터 만들래요" | Stage 1 (거의 구현됨) |
| **B 리디자인** | 초안을 읽어 한방에 변환(복제본 생성) + 추후 에셋 갈아끼우기 | "초안 있는데 못생겼어요" | Stage 2 (킬러 피처) |
| **C 기획+시안** | 구성 제안 → 시안(네모박스) → 박스 클릭 시 에셋 추천 | "백지인데 구성부터" | Stage 3 (추후) |

> Route A와 B는 **별개 기능**이다. A는 "조립", B는 "변환". 섞지 않는다. 비포애프터 개념은 B에서만 존재한다.

## 3. 핵심 아키텍처 결정

### 3.1 두 개의 토큰 예산 (대원칙)
- **인제스트 타임** (오프라인, 에셋당 1회, R&D팀): Vision·분해·풍부한 설명에 토큰을 아끼지 않는다. 비용은 그 에셋의 모든 미래 사용에 분산된다.
- **런타임** (온라인, 사용자 액션마다): 에셋을 *절대 재분석하지 않는다*. 미리 만든 **컴팩트 카탈로그** + 현재 슬라이드 구조만 전달해 "매칭"만 시킨다. 저렴한 모델로도 충분.
- 한 줄: **"이해는 인제스트에, 매칭은 런타임에서 싸게."**

### 3.2 역할 기반 색/폰트
- 저장 단위: 색 `{role, value, locked}`, 폰트 `{role, family, fallback, weight, source}`.
- **원본 값 보존** = 기본 컨셉("as-designed") = 의도 손실 0. 역할은 그 위의 *선택적* 치환 레이어.
- `locked: true`(로고/브랜드 색)는 어떤 컨셉에서도 안 바뀐다.
- 용도 의도는 색이 아니라 `use_when`/`provenance`에 따로 저장 — 색-역할과 직교.
- 리스크는 "역할 라벨 오태깅" → 인제스트 시 디자이너가 확인/수정(B4), 항상 원본 폴백 존재.

### 3.3 슬롯
- 에셋은 `slots: [{name, type, perSlide}]`를 선언. `perSlide: true`인 슬롯(예: `subtitle`)은 슬라이드마다 텍스트가 바뀌고, 나머지는 고정.
- **슬롯 식별 방식: 텍스트 박스 + shape 이름 규약 (placeholder 아님).**
  - Placeholder는 사용자 직접 입력엔 편하지만 썸네일에 샘플 텍스트가 안 보여 R&D·사용자 모두 불편. 우리 서비스는 AI가 슬롯을 채우므로 placeholder의 장점이 거의 안 쓰임.
  - 텍스트 박스에 shape 이름을 `slot.title` / `slot.subtitle` / `slot.body` / `slot.image1` 식으로 지정 → COM `shape.Name`으로 인식, 화면·썸네일에 영향 0.
  - 이미지 자리(점선 박스)도 `slot.image1` (type=`image`)로 동일 규약.
  - 슬롯 텍스트 박스는 **autofit/shrink 켜기** — 초안 텍스트가 샘플보다 길 때 깨지지 않게.
- 슬롯 type 어휘: `text | image | chart | table`.

### 3.3.1 에셋 2-tier 구조 (kind)
- **`kind: layout`** — 슬라이드 페이지 전체를 감싸는 템플릿. 슬롯이 박혀 있고, Route B에서 복제해서 채우는 대상.
  - 카테고리: 표지/목차/간지/연혁/3단가로/4단가로/5분할/6분할/좌텍스트우이미지/좌이미지우텍스트/마무리
- **`kind: component`** — 레이아웃 위에 붙이는 부품(그래프/다이어그램/표 등). Route A(조립)에서 하나씩 추천·삽입.
  - 기존 header_N.pptx는 전부 `component`.
- 대표 R&D 산출물이 이 2-tier를 이미 자연스럽게 따르고 있음(2026-06-19 확인).

### 3.4 컨셉 레이어 + 덱 레벨 vs 슬라이드 레벨
- **컨셉** = {팔레트(역할→hex) + 폰트 페어링(역할→family) + 스타일태그}. 기존 `StylePalette`/`StyleConfig`의 확장.
- **덱 레벨 (1회 락):** 컨셉 + **헤더/푸터 등 반복 요소**. 헤더는 슬라이드별로 추천하지 않고 덱 단위로 1개를 골라 모든 본문 슬라이드에 동일 적용(소제목만 슬라이드별 슬롯으로 변동). → 일관성이 설계로 보장됨.
- **슬라이드 레벨 (매번):** 본문 레이아웃 + 그 슬라이드 고유 콘텐츠. 표지/엔드는 예외 슬라이드 타입.
- 스키마 반영: 에셋에 `scope: deck | slide`, 슬롯에 `perSlide`.
- 컨셉 제안 UX: AI가 초안 덱을 읽어 ① "현재 느낌을 더 깔끔하게"(원본 존중) + ② 다른 방향 1~2개, 총 **2~3개 컨셉 카드**를 리디자인 탭 상단에 제시(StylePanel UI 재사용). 클릭 시 덱 컨셉 락.

### 3.5 비파괴 리디자인 + 비포애프터
- `Slide.Duplicate()`로 원본(N) 뒤에 복제본(N+1) 생성 → 복제본만 리디자인. 원본 무손상.
- 비포애프터: PowerPoint 슬라이드 썸네일 레일에 N/N+1이 나란히(공짜). 추가로 패널 내 WPF 비포/애프터 슬라이더(데모 와우, 기존 `ThumbnailService` PNG 재사용).

### 3.6 폰트 전략 (3단 방어)
- **캡처(인제스트, 토큰 0):** COM에서 `Font.Name/Size/Bold`를 읽는다. 폰트 이름은 .pptx XML에 저장돼 있어 *미설치 상태에서도 이름은 정확히* 읽힌다(렌더만 대체됨 → 그래서 Vision은 보조).
- **큐레이션 + 번들:** R&D팀은 **재배포 가능한(오픈소스) 폰트 팔레트**만 사용(Pretendard, Noto 등). 제품에 동봉.
- **보장(런타임):** 삽입 직전 필요한 폰트 설치 여부 확인 → 없으면 `%LOCALAPPDATA%\Microsoft\Windows\Fonts`에 **사용자 권한으로 자동 설치**(Win10 1809+ 관리자 불필요) → 못 구하면 `fallback` 체인 + "이 에셋은 X 폰트에서 제일 예뻐요" 고지.
- 참고: PPT "파일에 폰트 포함" 임베딩은 *열 때*만 유효, shape 복사엔 안 따라옴 → 번들+자동설치가 더 확실.

## 4. 컴포넌트 분해

```
A 데이터 계층 ──┬──> B 인제스트 (A에 기록)
                ├──> C 슬라이드 리더 (A의 슬롯 개념 공유)
                └──> D 리디자인 엔진 (A 카탈로그 + C 소비)
E LLM 토대 ────────> B, D 가 호출
F 기획 모드(Route C) ──> (미래) C·D의 슬롯 재사용
```

- **A. 데이터 계층** 〔제일 중요·토대〕 — A1 저장 스키마 v2(역할 색/폰트, 슬롯, scope, provenance, schemaVersion), A2 스키마 진화(저장모델↔런타임모델 분리 + 마이그레이터 + JsonExtensionData 관대수용), A3 카탈로그 빌더(저장→컴팩트 런타임), + 컨셉 역할치환 리졸버.
- **B. 인제스트 파이프라인** 〔R&D 도구·해자〕 — B1 신규 .pptx 감지(FileSystemWatcher)+COM 구조 추출(텍스트/위치/색/폰트/슬롯후보), B2 Vision 보조(썸네일로 미적 판단), B3 LLM 메타데이터 생성→스키마 반영→assets.json 업데이트→핫리로드, B4 인간 입력 머지(provenance/역할확인), B0 R&D팀 제작 규약(shape 네이밍으로 슬롯 선언, 폴더 규약).
- **C. 슬라이드 리더** 〔재사용 핵심 모듈〕 — C1 현재 슬라이드 COM 직렬화(텍스트/슬롯/색/폰트/구조), C2 슬라이드 PNG→Vision(옵션), C3 컴팩트 "현재 슬라이드 표현".
- **D. 리디자인 엔진(Route B)** 〔킬러 피처〕 — D0 컨셉 수립/락, D-헤더 덱레벨 락, D1 매칭(로컬 사전필터→LLM 후보 랭킹, 저토큰), D2 슬롯 매핑(초안 텍스트→에셋 슬롯), D3 적용(복제+역할치환+텍스트 흘려넣기), D4 비포애프터.
- **E. LLM 토대** 〔가볍게〕 — E1 MockAiService→ClaudeAiService(IAiService 고정), E2 API 키 관리(`%AppData%\TEAMPPT\settings.json`), E3 토큰 예산 원칙 코드화.
- **F. 기획 모드(Route C)** 〔추후〕 — 시안 생성 + `Shape.Tags`로 슬롯 도장 + `WindowSelectionChange`로 클릭 감지 → 추천. 지금은 자리만 비워둠.

## 5. 데이터 스키마 v2 (Phase A가 만드는 것)

기존 `assets.json`(schemaVersion 1, 색이 객체형)을 v2로 확장. **마이그레이터가 v1→v2를 흡수**하므로 구버전 파일도 안전.

```jsonc
{
  "schemaVersion": 2,
  "file": "header_3.pptx",
  "name": "장점 나열",
  "kind": "component",                   // layout=전체 슬라이드 틀 | component=레이아웃 위 부품
  "category": "헤더",
  "scope": "deck",                       // deck=반복요소(헤더/푸터) | slide=본문
  "content_fit": ["장점 나열", "특징 비교"],
  "use_when": "3~4가지 핵심 장점을 아이콘과 함께 나열할 때",
  "tags": ["장점", "나열", "아이콘"],
  "provenance": "삼성 IR 덱에서 사용",     // optional, 인간 입력, 신뢰도용
  "grid_columns": 1,
  "colors": [                            // v1의 {main,sub1,sub2,text} 객체 → 역할 배열로 마이그레이션
    { "role": "main", "value": "#2563EB", "locked": false },
    { "role": "sub1", "value": "#3B82F6", "locked": false },
    { "role": "sub2", "value": "#93C5FD", "locked": false },
    { "role": "text", "value": "#1E293B", "locked": false }
  ],
  "fonts": [
    { "role": "heading", "family": "Pretendard", "fallback": "맑은 고딕", "weight": "Bold", "source": "bundled" },
    { "role": "body",    "family": "Pretendard", "fallback": "맑은 고딕", "weight": "Regular", "source": "bundled" }
  ],
  "slots": [                            // shape 이름(slot.xxx) 기반 식별, placeholder 아님
    { "name": "title",    "type": "text",  "perSlide": true },
    { "name": "subtitle", "type": "text",  "perSlide": true },
    { "name": "body",     "type": "text",  "perSlide": true },
    { "name": "image1",   "type": "image", "perSlide": true }  // 이미지 슬롯 예시
  ]
}
```

**컨셉 모델 (DesignConcept):** `{ id, name, styleTags[], colors: {role→hex}, fonts: {role→family} }`.

**역할 치환 (ConceptResolver, 순수 함수, Phase A의 키스톤):** `Resolve(AssetRecord, DesignConcept)` → 각 색/폰트 역할에 대해 `locked==false`면 컨셉 값으로 치환, `locked==true`거나 컨셉에 해당 역할이 없으면 원본 유지. 이 함수가 D(리디자인)·E(스타일적용)의 심장이며 COM 없이 단위 테스트 가능.

**런타임 카탈로그 (CatalogEntry, CatalogBuilder):** 저장 레코드 → LLM 매칭용 컴팩트 투영. 필드 = `file, name, kind, category, scope, tags, useWhen, slotNames[], colorRoles[], fontRoles[]`. *색 hex·폰트 family 같은 무거운 값은 제외*(토큰 절감).

### R&D 제작 규약 (대표 동기화 사항, 2026-06-19)
1. **슬롯 도형 이름**: `slot.title` / `slot.subtitle` / `slot.body` / `slot.image1` 등 `slot.` 접두사 + 역할명.
2. **autofit/shrink**: 슬롯 텍스트 박스에 "넘치면 텍스트 줄이기" 켜기.
3. **색 역할 + locked**: 팔레트(main/sub/accent/text)와 브랜드/로고 색(locked) 구분 지정.
4. **폰트**: 재배포 가능(오픈소스) 폰트 사용 권장.
5. **종류 구분**: 레이아웃(.pptx)과 컴포넌트(.pptx)를 폴더 또는 파일명 규약으로 분리.

### Phase 경계 & 열린 항목 (괴리 방지 — 반드시 읽을 것)
- **Phase A는 대표 산출물과 무관하다.** 기존 `header_N.pptx` + `assets.json`만 다루는 순수 데이터 계층 작업이라, 이 단계에서는 대표 의도와 괴리가 발생할 수 없다. Sonnet은 Phase A에서 대표의 레이아웃 .pptx를 찾거나 인제스트하려 하지 말 것.
- **대표 합의가 실제로 필요한 시점 = Phase B(인제스트) / D(리디자인).** 그 전에 결정해야 할 열린 항목:
  - **(중요) 대표 템플릿은 현재 "이름 없는 일반 텍스트 박스"** 상태 — `slot.xxx` 네이밍 규약이 아직 미적용. Phase B 착수 전 둘 중 택1: **(a)** R&D가 슬롯 도형에 `slot.xxx` 이름 부여, 또는 **(b)** Phase B에 *이름 없는 텍스트박스용 휴리스틱 슬롯 추론*(샘플텍스트 "제목/소제목/내용" 패턴 + 위치/폰트크기 기반) 폴백 구현. → 현실적으로 (a)+(b) 병행 권장.
  - 색 역할/locked 매핑, 폰트 재배포 가능 여부, 차트/표/SmartArt 타입(리컬러 충실도) — 대표와 확정.
  - 레이아웃 .pptx를 컴포넌트와 어떻게 폴더/명명 분리할지.

## 6. 보류 항목 (Deferred)

- **캔버스 웹UI / 컨트롤 슬라이드** — "슬라이드 위 카드를 눌러 진행"하는 아이디어. 편집 모드에선 클릭=선택이라 진짜 버튼이 아니고 폴리싱 비용이 큼. **보류.** 모든 의사결정·진행은 **우측 패널에서 더 상세하게** 처리한다. (in-canvas 상호작용 원시기능은 미래 Route C에서만 검토.)

## 7. 개발 로드맵

| Phase | 내용 | 의존 | COM/LLM 필요? | 비고 |
|-------|------|------|----------------|------|
| **A** | 데이터 스키마 v2 + 마이그레이터 + 카탈로그 빌더 + 컨셉 리졸버 | — | ❌ 순수 로직 | **지금 착수**, 단위테스트 가능 |
| E | Claude API 연동(가볍게) | A | LLM | MockAiService→ClaudeAiService |
| C | 슬라이드 리더 | A | COM | 영구 재사용 모듈 |
| B | 인제스트 자동화 | A,E | COM+LLM+Vision | R&D 도구·해자 |
| D | Route B 리디자인 | A,C,E | COM+LLM | 킬러 피처 |
| F | Route C 기획 모드 | A,C,D | COM+LLM | 추후 |

각 Phase는 자체 spec→plan→구현 사이클을 갖는다(writing-plans 원칙: 1 plan = 1 subsystem).

## 8. 첫 Deliverable

**Phase A** — COM/PowerPoint/LLM 없이 순수 C#/JSON 로직만으로 완결·단위테스트 가능한 데이터 계층. 모든 후속 Phase의 토대.
→ 상세 TDD 계획: [docs/superpowers/plans/2026-06-19-phase-a-asset-schema.md](../plans/2026-06-19-phase-a-asset-schema.md)

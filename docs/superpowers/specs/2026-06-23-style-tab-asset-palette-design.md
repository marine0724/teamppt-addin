# 스타일 탭 — 에셋 기반 팔레트 & 슬라이드 적용 설계

날짜: 2026-06-23
상태: 설계 확정 (구현 대기 — Sonnet 실행 예정)

## 1. 문제 정의

현재 스타일 탭은 동작하지만 "거의 안 먹히는" 수준이다.

- 팔레트는 `styles.json`의 **하드코딩 5개 프리셋**만 보여준다 — 삽입한 에셋과 무관.
- 적용 시 **텍스트 색·폰트만** 바뀐다. 팔레트의 Main/Sub1/Sub2는 UI에만 보이고 슬라이드에 반영 안 됨. 도형 배경·선·슬라이드 배경은 전혀 안 건드림.
- 인제스트 파이프라인은 이미 에셋마다 `colors`(role+hex+locked), `fonts`(role+family)를 LLM으로 추출해 Supabase `metadata`에 저장하고, `SupabaseAssetMapper`로 `HeaderAsset.Colors`/`Fonts`까지 로드한다. **데이터는 다 있는데 스타일 탭이 안 읽는 게 핵심 문제.**

목표: 에셋의 실제 색을 앵커로 삼아 **컨셉이 통째로 바뀌는** 팔레트를 생성하고, 슬라이드 전체(배경·도형·텍스트)에 적용한다.

## 2. 핵심 결정 (확정된 의사결정)

1. **적용 시점 = 수동 선택.** 에셋을 삽입하면 그 에셋의 색이 스타일 탭에 뜨고, 사용자가 팔레트를 클릭해야 적용된다. (자동 적용 아님)
2. **role 기반 지능 매핑.** 색을 도형 종류·역할에 따라 매핑한다 (배경/fill/outline/텍스트). 단순 일괄 아님.
3. **locked(브랜드 고정색) 무시.** 지금은 고객 브랜드색을 모르므로 모든 색 변경 가능.
4. **에셋 팔레트가 메인, 공식은 변주 생성기.** 단일 색에서 5종 하모니를 기계적으로 다 뽑지 않는다. 에셋의 실제 팔레트(디자이너가 만든 검증된 색)를 앵커로 두고, 안전한 변환만 파생한다.
5. **두 메인 축 = 원본 + 반전.** (아래 예시 이미지로 검증됨)

### 2.1 "원본 + 반전" 공식 (검증된 성공 사례)

실제 목업 두 장(다크 메인 → 라이트 반전)에서 확인된 변환:

| 요소 | 원본(다크) | 반전(라이트) | 규칙 |
|---|---|---|---|
| 슬라이드 배경 | 다크 네이비 | 화이트 | 명도(L) 반전 |
| 본문 텍스트 | 화이트 | 다크 네이비 | 명도 반전 |
| 제목 강조 | 라이트 시안 | 로열 블루 | **Hue·S 유지, 새 배경 대비 맞춰 L만 조정** |
| 로고 칩(SAP) | 흰 칩 유지 | 흰 칩 유지 | 자체 배경 가진 요소는 미변경 |

핵심 원리: **Hue를 잠그고 명도/대비로 컨셉을 바꾼다** = 단색(Monochromatic) + 대비관리. 삼각·보색으로 Hue를 돌렸으면 강조색이 주황·마젠타로 튀어 촌스러워진다.

### 2.2 두 부류의 팔레트

- **에셋 기반 팔레트** (원본·반전·안전한 변주): Hue 잠금, 브랜드 정체성 유지.
- **대체 컨셉 팔레트** (기존 하드코딩 프리셋): Hue 자유. "완전히 다른 분위기를 원할 때"용. 빨강/노랑·크림 같은 것. 그대로 유지.

## 3. 아키텍처

LLM 호출 없이 **순수 C# 색상 수학**으로 처리한다. 새 컴포넌트는 작고 단일 책임으로 쪼갠다.

```
에셋 삽입 (클릭/드래그)
  → HeaderAsset.Colors / .Fonts 추출
  → PaletteRoleMapper: AssetColor[] → 정규화된 5-role 팔레트 (빈 role은 휴리스틱 보강)
  → PaletteGenerator: 원본 + 반전 + 안전 변주 N개 생성
  → (+ styles.json 대체 컨셉 프리셋)
  → 스타일 탭 UI 갱신 (PopulateStylePanel)
  → 사용자 클릭 → SlideStyleApplier: 슬라이드 배경·도형·텍스트에 적용
```

### 3.1 새/변경 컴포넌트

| 컴포넌트 | 책임 | 종류 |
|---|---|---|
| `ColorHsl` (신규) | RGB↔HSL 변환, WCAG 대비비 계산, L 조정 헬퍼 | 순수 함수, 테스트 용이 |
| `PaletteRoleMapper` (신규) | `List<AssetColor>` → 5-role 정규 팔레트. 빈 role 휴리스틱 보강 | 순수 함수 |
| `PaletteGenerator` (신규) | 정규 팔레트 → `List<StylePalette>` (원본·반전·변주) | 순수 함수 |
| `SlideStyleApplier` (신규) | `StylePalette`+`StyleFont` → 활성 슬라이드 적용 (배경·도형·텍스트) | PPT Interop |
| `PaletteColors` (확장) | `Background` 필드 추가 | 모델 |
| `AssetPanel.PopulateStylePanel` (변경) | 에셋 팔레트 반영하도록 갱신 | UI |
| `TaskPaneHost.OnStyleApply` (변경) | `SlideStyleApplier` 호출로 위임 | UI 글루 |

`OnStyleApply`의 기존 텍스트 적용 로직은 `SlideStyleApplier`로 옮겨 확장한다 (TaskPaneHost가 비대해지지 않게).

### 3.2 데이터 모델 변경

`PaletteColors`에 `Background` 추가 (5-role 구조):

```csharp
public class PaletteColors
{
    public string Background { get; set; } // 슬라이드 배경 (신규)
    public string Main { get; set; }       // 도형 fill·강조
    public string Sub1 { get; set; }       // 도형 outline·보조
    public string Sub2 { get; set; }       // 추가 액센트
    public string Text { get; set; }       // 텍스트
}
```

`Background`는 nullable — 기존 `styles.json` 프리셋은 이 필드가 없으므로 적용 시 null이면 슬라이드 배경을 건드리지 않는다 (하위 호환).

## 4. 색상 로직 상세

### 4.1 role 매핑 (PaletteRoleMapper)

`AssetColor.Role`은 LLM이 생성한 자유 문자열(`main`/`sub1`/`sub2`/`text`/`accent`/`background` 등). 정규화 규칙:

- **main**: role이 `main`/`primary`면 그것. 없으면 채도 가장 높은 색.
- **text**: role이 `text`면 그것. 없으면 명도 극단값(배경과 대비 큰 쪽).
- **background**: role이 `background`면 그것. 없으면 가장 밝은 색, 그것도 없으면 main의 아주 밝은 틴트(L≈95%).
- **sub1/sub2**: 남은 색 순서대로. 부족하면 main의 명도 변주로 채움.

원칙: **있는 실제 색을 최우선 사용, 빈 role만 수학으로 보강.**

### 4.2 반전 알고리즘 (PaletteGenerator)

원본 정규 팔레트 → 반전 팔레트:

1. Hue 고정 (브랜드 정체성).
2. `Background.L ↔ Text.L` 스왑 (다크 ↔ 라이트).
3. 강조색(Main/Sub): H·S 유지, **L을 새 배경 대비 WCAG ≥ 4.5:1 되도록 자동 조정**. 라이트 배경 → 강조 진하게, 다크 배경 → 강조 밝게.
4. 대비 미달 시 L을 단계적으로 조정해 임계 통과시킴.

### 4.3 안전 변주 (선택, 우선순위 낮음)

원본에서 추가로:
- **단색(Monochromatic)**: Hue 고정, S/L 변주 → 톤 정리된 버전.
- **유사색(Analogous)**: Hue ±30° → 약간의 변화.
- **삼각·분할보색은 기본 제외.** 넣더라도 채도 클램프(S 상한)로 강제.

> 변주는 v1 필수가 아니다. **원본+반전이 핵심.** 변주는 시간 남으면.

## 5. 슬라이드 적용 (SlideStyleApplier)

활성 슬라이드 1장 대상. role 기반 매핑:

| 대상 | 매핑 색 | 비고 |
|---|---|---|
| 슬라이드 배경 | `Background` | null이면 미변경 |
| 도형 fill (solid·opaque) | `Main` | 투명/채우기 없음 도형은 미변경 (의도된 빈자리) |
| 도형 outline (선 있음) | `Sub1` | 선 없으면 미변경 |
| 텍스트 | `Text` + 선택 폰트 | 단락별 적용 (기존 로직 계승) |

**제외 대상 (절대 안 건드림):**
- Picture/이미지 shape, 로고, 자체 배경 가진 요소 — 색 입히면 깨짐 (반전 예시에서 SAP 칩이 살아남은 이유).
- `MsoShapeType` 검사로 `msoPicture`·`msoPlaceholder`(이미지) 스킵.

**대비 안전장치:** 텍스트 색 적용 후 해당 도형 배경과 대비가 임계 미달이면 텍스트 L을 보정 (선택, 여력 시).

## 6. 엣지 케이스

| 상황 | 처리 |
|---|---|
| 에셋에 colors 없음(빈 배열) | 에셋 팔레트 생성 스킵, 하드코딩 프리셋만 표시 |
| 색이 1개뿐 | 단색 변주로 5-role 채움 |
| 반전 대비 미달 | L 단계 조정으로 WCAG 통과시킴 |
| `styles.json` 프리셋(Background 없음) | 적용 시 배경 미변경 (하위 호환) |
| 슬라이드에 도형 다수·이미지 혼재 | 이미지·로고 스킵, 나머지만 |

## 7. 테스트

**단위 테스트 (순수 함수 — 우선):**
- `ColorHsl`: RGB↔HSL 라운드트립, WCAG 대비비 계산 정확도.
- `PaletteRoleMapper`: role 누락 시 휴리스틱 보강 (main/text/background 추론).
- `PaletteGenerator` 반전: Background/Text 명도 스왑 확인, 강조색 대비 ≥ 4.5:1 보장.

**수동 검증:**
- 에셋 삽입 → 스타일 탭에 원본+반전 팔레트 표시.
- 원본 클릭 → 슬라이드 다크 컨셉, 반전 클릭 → 라이트 컨셉 (목업 1·2번처럼).
- 이미지·로고가 안 깨지는지 확인.

## 8. 범위 밖 (YAGNI)

- 사용자 커스텀 팔레트 저장/편집.
- 슬라이드 전체(덱) 일괄 적용 — v1은 활성 슬라이드 1장.
- 차트·SmartArt·그라데이션 정밀 스타일링.
- locked 색상 보호 UI (브랜드색 미확보로 보류).
- 삼각/분할보색 등 위험 하모니의 기본 제공.

## 9. 기존 코드 연결

- `HeaderAsset.Colors`/`Fonts` (이미 매핑됨) → 팔레트 생성 입력.
- `AssetPanel.PopulateStylePanel` / `_styleConfig` → 생성 팔레트 + 프리셋 병합 표시.
- `TaskPaneHost.OnStyleApply` → `SlideStyleApplier`로 위임.
- 설계 원칙 유지: Core/Connect.cs/Globals.cs 구조 존중, Newtonsoft.Json, CoordinateConverter에 폴백 로직 추가 금지.

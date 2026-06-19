# TEAMPPT Add-in 개발 인계문서

> 최종 업데이트: 2026-06-19 (Phase E 완료, 구조화 출력 완료, **Phase F 설계 확정·API키 블로커**)  
> 프로젝트 경로: `C:\Projects\teamppt-addin\src\TeampptAddin`

---

## ⚡ 재개 지점 (READ FIRST — 새 세션은 여기부터)

**브랜치:** `phase-e-gemini-flash`

**최근 완료:**
- 구조화 출력(Gemini `responseSchema`): `BuildResponseSchema()` 추가, 프롬프트 형식블록 제거, 요청에 연결. 단위테스트 19개 PASS. → [plans/2026-06-19-structured-output-responseschema.md](docs/superpowers/plans/2026-06-19-structured-output-responseschema.md)
- 503/429/500 일시오류 지수백오프 재시도 추가 ([GeminiAiService.cs](src/TeampptAddin/Services/GeminiAiService.cs)).
- 유출된 평문 API 키를 계획문서에서 제거(키는 Google이 비활성화됨).

**🚫 현재 블로커 — Gemini API 키:**
AI Studio에서 받은 키가 `AQ.`로 시작하는 53자(임시/OAuth 토큰, 곧 만료 → 400 `API_KEY_INVALID`). **정상 Gemini API 키는 `AIza`로 시작 ~39자.** 올바른 "API key"를 받아 `src/TeampptAddin/Assets/api-keys.json`에 넣고 **관리자 재빌드**(아래 4.2/Start-Process RunAs) 필요. ⚠️ repo가 **공개**라 키를 문서/커밋에 절대 넣지 말 것.

**빌드 메모(이 PC):** MSBuild = `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`. 단위테스트는 관리자 불필요: 본프로젝트 `/p:RegisterForComInterop=false` 빌드 → 테스트프로젝트 `/p:BuildProjectReferences=false` 빌드 → `dotnet test --no-build --no-restore`. (`dotnet test` 직접 실행은 레거시 csproj 복원 깨짐.)

**다음 작업 — Phase F (Supabase 벡터 + 대화 메모리):** 설계 확정, 미결 7개. → [specs/2026-06-19-phase-f-supabase-vector-design.md](docs/superpowers/specs/2026-06-19-phase-f-supabase-vector-design.md)
순서: ① (사용자) AIza 키 확보 → 구조화 출력 PowerPoint 검증 → ② 미결 7개 결정 → ③ Phase F 코드 계획 작성·실행.

---

---

## 1. 프로젝트 개요

PowerPoint COM Add-in (.NET Framework 4.8). 오른쪽 Task Pane에 헤더 에셋 카드 목록을 표시하고, 클릭 또는 드래그앤드롭으로 현재 슬라이드에 Shape을 삽입하는 도구.

### 핵심 기능
- **카드 리스트**: `assets/headers/header_1.pptx` ~ `header_7.pptx`를 읽어서 카드로 표시
- **썸네일**: Shape-only PNG export (투명 배경) → 캐시 → 카드/고스트에 사용
- **클릭 삽입**: 카드 클릭 → `ShapeInserter.InsertToActiveSlide` → 슬라이드 중앙에 삽입
- **드래그앤드롭**: 마우스 캡처 방식 (OLE DragDrop 아님) → 고스트 윈도우 → 드롭 위치에 삽입

---

## 2. 파일 구조 및 역할

> 2026-06-18 모듈화 완료. Core/UI/Models/Interop 분리.  
> Phase 3: WPF UI 추가 (UI/Wpf/). 기존 WinForms UI는 폴백으로 보존.

### 루트

| 파일 | 역할 | 상태 |
|------|------|------|
| `Connect.cs` | COM 진입점, `IDTExtensibility2` + `ICustomTaskPaneConsumer` | 완성 |
| `Globals.cs` | 전역 `Application` 참조, `AssetsDir`, `ThumbnailDir` 경로 | 완성 |
| `install.bat` / `uninstall.bat` | 수동 설치/제거 스크립트 | 완성 |

### Core/ — UI 무관, WinForms 참조 없음 (VSTO 전환 시 그대로 재사용)

| 파일 | 역할 |
|------|------|
| `Logger.cs` | 디버그 로깅 (`%LocalAppData%\TeampptAddin\debug.log`) |
| `ShapeInserter.cs` | Shape 복사/붙여넣기 (CopyShapesToClipboard, InsertToActiveSlide) |
| `ThumbnailGenerator.cs` | COM 기반 shape-only PNG export + slide export 폴백 |
| `CoordinateConverter.cs` | PointsToScreenPixels 역변환, 슬라이드 좌표 계산, ShapeRange 위치 지정 |

### UI/ — TaskPaneHost + WinForms 폴백

| 파일 | 역할 |
|------|------|
| `TaskPaneHost.cs` | COM 호스팅 컨테이너. WPF 초기화 → 실패 시 WinForms 폴백. WPF 드래그 처리 |
| `CardControl.cs` | WinForms 카드 렌더링 (폴백용, OnPaint 기반) |
| `DragHandler.cs` | WinForms 드래그 상태 관리 (폴백용) |
| `GhostWindow.cs` | WS_EX_LAYERED 투명 윈도우 (WPF/WinForms 공통 사용) |

### UI/Wpf/ — WPF UI (현재 활성)

| 파일 | 역할 |
|------|------|
| `AssetCard.cs` | WPF 카드 컨트롤 (호버 스케일 애니메이션, 호버 Popup 프리뷰, 클릭/드래그 이벤트) |
| `AssetPanel.cs` | WPF 메인 패널 (AI/에셋/스타일 3탭, AI 채팅+애니메이션, 카테고리 필터, StylePanel) |
| `ThemeResources.cs` | Theme.xaml 런타임 미러 (정적 Brush/CornerRadius/FontFamily) |

### Models/

| 파일 | 역할 |
|------|------|
| `HeaderAsset.cs` | 에셋 저장 모델 v2 (SchemaVersion/Kind/Scope/Provenance/Colors→List\<AssetColor\>/Fonts/Slots) |
| `AssetSchema.cs` | 값 타입: AssetColor{Role,Value,Locked}, AssetFont{Role,Family,Fallback,Weight,Source}, AssetSlot{Name,Type,PerSlide} |
| `DesignConcept.cs` | 컨셉 모델 (Id/Name/StyleTags/Colors:역할→hex/Fonts:역할→family) |
| `CatalogEntry.cs` | 런타임 컴팩트 카탈로그 항목 (hex/family 제외, 역할·슬롯 이름만 투영) |
| `StylePalette.cs` | 스타일 데이터 모델 (PaletteColors, StylePalette, StyleFont, StyleConfig) |
| `AiRecommendation.cs` | AI 추천 응답 모델 (AssetSuggestion, StyleSuggestion, AiRecommendation) |

### Services/

| 파일 | 역할 |
|------|------|
| `AssetLoader.cs` | assets.json 파싱 (JArray→마이그레이터→바인딩) + 폴더 스캔 폴백 + 파일 존재 검증 |
| `AssetSchemaMigrator.cs` | v1→v2 정규화 (객체형 colors→역할 배열, scope/kind 기본값) |
| `CatalogBuilder.cs` | HeaderAsset 리스트 → CatalogEntry 리스트 (무거운 값 제외, 역할/슬롯 이름만) |
| `ConceptResolver.cs` | 역할 치환 순수 함수 (ResolveColors/ResolveFonts, locked·missing 존중) |
| `StyleLoader.cs` | styles.json 파싱 + 하드코딩 기본값 폴백 |
| `IAiService.cs` | IAiService 인터페이스 + MockAiService stub |
| `ThumbnailService.cs` | LoadThumbnail, LoadImageNoLock (TaskPaneHost에서 추출) |

### Interop/

| 파일 | 역할 |
|------|------|
| `IObjectSafety.cs` | COM IObjectSafety 인터페이스 정의 |

---

## 3. 아키텍처 핵심 결정사항

### 3.1 WPF 호스팅 전략
- **Phase 2 발견**: 생성자에서 ElementHost 초기화 → COM 충돌. `OnSizeChanged(Width > 0)`에서 지연 생성 → 정상 동작
- **Phase 3 구조**: TaskPaneHost(WinForms UserControl, COM 필수) → ElementHost → AssetPanel(WPF)
- **폴백**: WPF 초기화 실패 시 기존 WinForms UI 자동 로드
- **드래그**: WPF 카드에서 이벤트 발생 → TaskPaneHost에서 Win32 Capture 기반 처리 (GhostWindow 공유)

### 3.2 COM 등록 요구사항
`TaskPaneHost`가 PowerPoint의 `CreateCTP`에서 ActiveX로 호스팅되려면:
1. `[ComVisible(true)]`, `[Guid]`, `[ProgId]` 어트리뷰트
2. `IObjectSafety` 인터페이스 구현 (보안 검증용)
3. **레지스트리에 Control 카테고리 수동 등록** 필수:
   ```
   HKCR\CLSID\{2D4E6F8A-...}\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}
   ```
   `RegAsm`만으로는 이 카테고리가 안 들어감. 수동으로 추가해야 함.

### 3.3 드래그앤드롭: Win32 마우스 캡처 방식
- WPF `DragDrop.DoDragDrop()` → PowerPoint 슬라이드 위 드롭 시 실패
- **현재 방식** (PowerMockup 스타일):
  1. `MouseDown` → 드래그 시작 위치 기록
  2. `MouseMove` → 임계값 초과 시 `CopyShapesToClipboard` + `Capture = true` + 고스트 윈도우 표시
  3. `MouseUp` → Task Pane 밖이면 `slide.Shapes.Paste()` + 좌표 변환으로 위치 지정

### 3.4 고스트 윈도우: WS_EX_LAYERED + UpdateLayeredWindow
- `Form` 상속, `WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW`
- `UpdateLayeredWindow` P/Invoke로 per-pixel alpha 렌더링
- `SourceConstantAlpha = 180` (~70% 전체 불투명도) + shape PNG의 자체 투명도
- 거의 실제 사이즈 (화면 85%까지 스케일), 커서 중앙 정렬

### 3.5 썸네일: Shape-only export
```csharp
var range = slide.Shapes.Range(indices);
var group = range.Group();
group.Export(path, PpShapeFormat.ppShapeFormatPNG);
group.Ungroup();
```
- 슬라이드 배경 없이 shape만 투명 배경 PNG로 추출
- 단일 shape이면 `shape.Export()` 직접 호출
- 실패 시 `slide.Export("PNG")` 폴백 (배경 포함)
- 캐시: `%LocalAppData%\TeampptAddin\thumbnails\header_N.png`
- pptx 수정일 vs 캐시 수정일 비교하여 재생성

### 3.6 드롭 위치 좌표 변환
```
PointsToScreenPixelsX(0), PointsToScreenPixelsY(0) → 슬라이드 원점의 스크린 좌표
PointsToScreenPixelsX(500), PointsToScreenPixelsY(500) → 스케일 계산용 참조점
scaleX = (refX - originX) / 500
slideX = (mouseScreenX - originX) / scaleX
```
- 슬라이드 범위(0~slideW, 0~slideH)로 클램핑
- 붙여넣은 ShapeRange의 중심을 드롭 좌표로 이동

---

## 4. 빌드 및 배포 절차

### 4.1 테스트 실행
```powershell
# 1) MSBuild로 솔루션 빌드 (dotnet은 COM 프로젝트 빌드 불가)
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  src\TeampptAddin\TeampptAddin.sln /p:RegisterForComInterop=false /v:quiet
# 2) vstest.console.exe로 테스트 DLL 실행
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  src\TeampptAddin.Tests\bin\Debug\net48\TeampptAddin.Tests.dll
```
> `dotnet test`는 legacy COM 프로젝트 빌드 실패 → MSBuild+vstest 폴백 확정 (Phase A Task 1)

### 4.2 빌드
```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "C:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.sln" /p:Configuration=Debug /v:minimal
```
> **주의**: PowerPoint가 DLL을 잠그므로 빌드 전 반드시 PPT 종료

### 4.3 COM 등록 (관리자 권한 필요)
```powershell
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" `
  "C:\Projects\teamppt-addin\src\TeampptAddin\bin\Debug\TeampptAddin.dll" /codebase /tlb

# Control 카테고리 수동 등록 (필수!)
New-Item -Path "Registry::HKEY_CLASSES_ROOT\CLSID\{2D4E6F8A-1B3C-5D7E-9F0A-4C6E8D2B1A3F}\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}" -Force
```

### 4.4 PowerPoint Add-in 레지스트리
```powershell
$regPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect"
New-Item -Path $regPath -Force
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "TEAMPPT"
Set-ItemProperty -Path $regPath -Name "Description" -Value "TEAMPPT Header Assets Add-in"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Type DWord -Value 3
```

### 4.5 썸네일 캐시 클리어 (shape-only 재생성 시)
```powershell
Get-ChildItem "$env:LOCALAPPDATA\TeampptAddin\thumbnails" -File |
  ForEach-Object { Remove-Item $_.FullName -Force }
```

### 4.6 전체 원커맨드 배포
```powershell
Stop-Process -Name "POWERPNT" -Force -ErrorAction SilentlyContinue
Start-Sleep 2
# Build → RegAsm → Control Category → Clear cache → Launch PPT
```

---

## 5. 디버깅

- **로그 위치**: `%LocalAppData%\TeampptAddin\debug.log`
- 타임스탬프 + 메시지 형식: `[HH:mm:ss.fff] message`
- 주요 로그 포인트:
  - Constructor STA 확인
  - 썸네일 생성 성공/실패
  - BeginDrag / EndDrag
  - 좌표 변환 (Coord, Raw slide, Shape count, Center, delta)
  - PositionShapes 성공/실패

---

## 6. 현재 동작 상태 (2026-06-18 Step 9 완료 후)

| 기능 | 상태 | 비고 |
|------|------|------|
| Task Pane 표시 | **동작** | COM 호스팅 → ElementHost → WPF |
| 3탭 네비게이션 | **동작** | AI / 에셋 / 스타일 탭 전환 |
| AI 채팅 | **동작** | MockAiService stub, 채팅 버블, 에셋 추천 카드 |
| AI 애니메이션 | **동작** | 빈 상태 아이콘 펄스, 로딩 점 바운스, 타이핑 효과, 카드 스태거 슬라이드인, 크로스페이드 |
| 에셋 스키마 v2 | **동작** | 역할 기반 색/폰트 + 슬롯 + scope + kind, AssetSchemaMigrator v1→v2 자동 변환 |
| 카탈로그 빌더 | **동작** | HeaderAsset→CatalogEntry 컴팩트 투영 (런타임 토큰 절감) |
| 컨셉 리졸버 | **동작** | DesignConcept으로 역할 치환 (locked/missing 존중) |
| 단위 테스트 | **동작** | xUnit 12개 PASS (src/TeampptAddin.Tests, MSBuild+vstest 러너) |
| 카드 리스트 렌더링 | **동작** | WPF AssetPanel + AssetCard, 라이트 테마 |
| 카테고리 필터 | **동작** | 전체/헤더/섹션/레이아웃/마무리 |
| 호버 Popup 프리뷰 | **동작** | 300ms 딜레이, 카드 간 즉시 전환, fade+slide 애니메이션 |
| StylePanel | **동작** | 팔레트 카드 5개 + 폰트 칩 4개 + 적용 버튼 |
| 스타일 적용 | **동작** | 현재 슬라이드 텍스트 shape에 폰트/색상 일괄 적용 (COM API) |
| 썸네일 (shape-only PNG) | **동작** | Group→Export→Ungroup, 캐시 |
| 클릭 삽입 | **동작** | InsertToActiveSlide |
| 드래그앤드롭 | **동작** | TaskPaneHost에서 Win32 Capture 기반 처리 |
| 고스트 윈도우 | **동작** | UpdateLayeredWindow, WPF/WinForms 공통 |
| 호버 애니메이션 | **동작** | WPF ScaleTransform 1.02x + 배경/테두리 색상 전환 (150ms ease) |
| WinForms 폴백 | **대기** | WPF 초기화 실패 시 기존 WinForms UI 자동 로드 |

---

## 7. 다음 실행사항 (TODO)

### 🎯 2026-06-19 방향 재정립 — AI 리디자인 제품 설계 확정

> 전체 설계: [docs/superpowers/specs/2026-06-19-teamppt-ai-redesign-design.md](docs/superpowers/specs/2026-06-19-teamppt-ai-redesign-design.md)
> 첫 실행계획(착수 대기): [docs/superpowers/plans/2026-06-19-phase-a-asset-schema.md](docs/superpowers/plans/2026-06-19-phase-a-asset-schema.md)

**제품 핵심:** 사용자가 만든 초안 슬라이드를 AI가 **한방에 리디자인**(비파괴 — `Slide.Duplicate()` 복제본에 적용, 비포애프터는 썸네일 레일로 공짜). 품질의 레버리지는 **데이터 스키마**.

**확정된 5대 결정:**
1. **두 토큰 예산** — 이해는 인제스트 타임(에셋당 1회, 펑펑)에, 매칭은 런타임(매 요청, 저렴)에.
2. **역할 기반 색/폰트** — `{role,value,locked}`. 원본 보존 + 선택적 재테마 → 적은 에셋이 많게 느껴짐.
3. **슬롯(slot)** — 에셋이 `title/subtitle/body` 자리 선언. 스키마·리더·리디자인을 잇는 다리.
4. **컨셉 레이어** — 덱 단위로 {팔레트+폰트+스타일태그} 1회 락. 헤더는 **덱 레벨 상수**(소제목만 슬라이드별 슬롯).
5. **비파괴 리디자인** — 복제 후 복제본만 수정.

**세 루트(스테이징):** A 조립(현재) → B 리디자인(킬러, 다음) → C 기획+시안(추후).

**개발 순서:** **Phase A(데이터 스키마, 지금)** → E(Claude API 가볍게) → C(슬라이드 리더) → B(인제스트 자동화) → D(Route B 리디자인).

**폰트 전략:** 캡처(COM, 토큰0) + 큐레이션·번들(오픈소스 폰트) + 런타임 사용자권한 자동설치 + fallback 체인.

**에셋 2-tier (대표 동기화 2026-06-19):**
- **layout** = 슬라이드 전체 틀 (표지/목차/간지/연혁/3단가로/4단가로/5분할/6분할/좌텍스트우이미지/좌이미지우텍스트/마무리)
- **component** = 레이아웃 위에 붙이는 부품 (그래프/다이어그램/표). 기존 header_N은 component.
- 스키마에 `kind: layout | component` 필드 추가.
- **슬롯 식별: 텍스트 박스 + shape 이름 규약** (`slot.title`/`slot.image1` 등). Placeholder 아님(썸네일에 샘플텍스트 보여야 하므로). autofit/shrink 켜기.
- R&D 제작 규약: slot 접두사 네이밍 + 색 역할/locked 지정 + 오픈소스 폰트 사용 + layout/component 폴더 분리.

**Phase 경계 (괴리 방지):** Phase A는 **대표 산출물과 무관** — 기존 header_N + assets.json만 다루는 순수 데이터 작업이라 이 단계에선 괴리 불가. 대표 합의가 필요한 건 Phase B(인제스트)/D(리디자인). **열린 항목**: 대표 템플릿은 현재 "이름 없는 텍스트 박스"라 `slot.xxx` 규약 미적용 → Phase B 전에 (a) R&D가 네이밍 적용 또는 (b) 휴리스틱 슬롯 추론(샘플텍스트+위치) 폴백 결정 필요. 상세는 설계 문서 §5 "Phase 경계 & 열린 항목".

**보류:** "캔버스 위 카드 클릭(웹UI)" 아이디어 — 편집모드 클릭=선택이라 폴리싱 비용 큼. **모든 진행은 우측 패널에서 상세히 처리.**

**Phase A (완료 2026-06-19):** 스키마 v2 + AssetSchemaMigrator + CatalogBuilder + ConceptResolver + xUnit 12개 PASS. `phase-a-asset-schema` 브랜치 → main 머지 완료.

**Phase E (다음):** Gemini Flash API 연동. MockAiService → GeminiAiService. API 키는 JSON 파일(gitignore). 향후 저장/인제스트는 Claude API, 런타임은 Gemini 이원화 예정이나 현재 테스트 단계에서는 전부 Gemini Flash.

---

### Phase 1 완료: 모듈화 (2026-06-18)
- Core/UI/Models/Interop 분리 완료
- 미사용 파일 제거 (TaskPaneControl.cs)
- HeaderAsset WPF→System.Drawing 전환 완료
- 빌드 검증 완료 (에러/경고 0)

### Phase 2 완료: WPF 지연 초기화 검증 (2026-06-18)
- **VSTO 전환 불필요** — COM Add-in 그대로 유지
- WPF ElementHost를 `OnSizeChanged(Width > 0)`에서 생성하면 COM 충돌 없음 확인
- 핵심 원리: 생성자(COM 초기화 중)가 아닌 지연 시점에서 ElementHost 생성
- csproj에 WPF 참조 추가 완료 (PresentationCore, PresentationFramework, WindowsFormsIntegration 등)

### Phase 3 완료: WPF UI 전환 (2026-06-18)
- `UI/Wpf/AssetCard.cs`: WPF 카드 (호버 스케일 애니메이션, 클릭/드래그 이벤트)
- `UI/Wpf/AssetPanel.cs`: WPF 패널 (헤더 + ScrollViewer + 상태바)
- `TaskPaneHost.cs`: OnSizeChanged에서 ElementHost 생성, WPF 드래그는 Win32 Capture로 처리
- 기존 WinForms UI (CardControl, DragHandler) 폴백으로 보존
- GhostWindow은 Win32 기반으로 WPF/WinForms 공통 재사용
- **주의**: CoordinateConverter에 폴백 로직 추가 금지 (기존 non-fatal 패턴 유지)

### Phase 4 완료: 에셋 스키마 확장 + AI 애니메이션 (2026-06-18)
- Step 8: ThumbnailService.cs 분리 (LoadThumbnail, LoadImageNoLock → TaskPaneHost에서 추출)
- Step 8.5: HeaderAsset에 Tags/Colors/JsonExtensionData 추가, AssetColors 클래스 신규, assets.json에 tags+colors 필드 채움
- Step 9: AI탭 애니메이션 5종 — 빈 상태 아이콘 펄스, 로딩 점 바운스, 타이핑 효과, 카드 스태거 슬라이드인, 크로스페이드

### Phase A 완료: 에셋 데이터 스키마 v2 (2026-06-19)
- xUnit 테스트 프로젝트 구축 (MSBuild + vstest.console.exe 러너)
- AssetColor/AssetFont/AssetSlot 값 타입 + HeaderAsset v2 확장 (SchemaVersion/Kind/Scope/Provenance/Fonts/Slots, Colors→List)
- AssetSchemaMigrator: v1 객체형 colors→v2 역할 배열, scope/kind 기본값 자동 부여
- AssetLoader가 마이그레이터 경유 (JArray→Migrate→ToObject)
- DesignConcept + CatalogEntry + CatalogBuilder (컴팩트 런타임 투영)
- ConceptResolver: 역할 치환 순수 함수 (locked·missing 존중)
- assets.json 7개 항목 v2로 재작성 (roles/fonts/slots/scope)
- 기존 AssetColors 클래스 삭제 (소비처 없음 확인 완료)
- 12개 단위테스트 전부 PASS

### Phase E (다음): AI 서비스 연동 — Gemini Flash
- MockAiService → GeminiAiService 교체 (IAiService 인터페이스 유지)
- **Gemini Flash** API 연동 (테스트 단계, 전부 Gemini Flash로 진행)
- API 키: `src/TeampptAddin/Assets/api-keys.json` (gitignore 처리)
- 향후 계획: 저장/인제스트 시 Claude API, 사용자 런타임은 Gemini — 현재는 전부 Gemini Flash

### 잔여 TODO (우선순위 순)

#### Step 10 (Phase E): Gemini Flash API 연동
- `MockAiService` → `GeminiAiService` 교체 (IAiService 인터페이스 그대로 유지)
- API 키: `api-keys.json` 파일 (gitignore 처리)
- 사용자 입력 + CatalogEntry 컨텍스트 → Gemini Flash API 호출
- 응답을 `AiRecommendation` JSON으로 파싱 (에셋 파일명 3개 + 메시지)
- UI 변경 없음

#### Step 11: 드래그 UX 개선
- 드래그 중 커서가 PowerPoint 슬라이드 위에 있을 때 투명 오버레이 창 표시
  - Win32 `WindowFromPoint`로 슬라이드 캔버스 감지
  - `WS_EX_TRANSPARENT + WS_EX_LAYERED` 오버레이로 삽입 위치 미리보기
- 드롭 시 커서 위치 대신 템플릿 원래 포지션으로 삽입
  - `CoordinateConverter.PositionShapesAtCursor` 대신 원본 좌표 유지

#### Step 12: 색상 일괄 변경 완성
- 현재: 텍스트 색만 변경
- 추가: shape fill 색도 팔레트에 따라 치환
  - `asset.colors.main` → `palette.colors.main` 1:1 매핑
  - COM API로 슬라이드 내 해당 fill 색 일괄 교체

#### Step 13: 데이터 저장
- 마지막 선택 팔레트/폰트 저장 → 앱 재시작 후 복원
- 저장 위치: `%AppData%\TEAMPPT\settings.json`

#### Step 14: 에셋 자동 감지 + LLM 메타데이터 생성
- `FileSystemWatcher`로 Assets 폴더 감시
- 신규 .pptx 감지 → 썸네일 추출 → Claude API로 name/category/use_when/tags/colors 자동 생성
- `assets.json` 자동 업데이트 → 앱 핫 리로드

#### Step 15: 초안 기반 추천 (킬러 피처)
- AI탭에 "초안 입력" 모드 추가
- 사용자가 발표 개요/목차 입력
- Claude가 슬라이드 구성 제안 + 각 슬라이드에 맞는 에셋 매핑
- 슬라이드별 카드 그룹으로 표시

#### Step 16: 키워드 검색
- 에셋 탭 상단 검색 인풋
- name, tags, use_when 필드 대상 필터링

#### Step 17: 인스톨러
- InnoSetup으로 `TEAMPPT_Setup.exe` 생성
- DLL + Assets + RegAsm 등록 + 레지스트리 키 자동 설정

#### 기술 부채
- WinForms 폴백 코드 정리 (WPF 안정화 확인 후 CardControl/DragHandler 제거 검토)
- Strong Name 서명 (RegAsm 경고 해결)
- 다중 모니터 / DPI 스케일링 테스트

---

## 8. 알려진 이슈 / 주의사항

1. **PowerPoint가 DLL 잠금**: 빌드 전 반드시 PPT 종료. 안 그러면 MSB3027 에러
2. **RegAsm unsigned 경고**: 기능에는 영향 없음. Strong Name 서명으로 해결 가능
3. **Control 카테고리 누락**: RegAsm 재실행 시 `{40FC6ED4-...}` 카테고리가 사라질 수 있음. 항상 수동 추가 필요
4. **OnSizeChanged에서 WPF 초기화**: `Width > 0`인 첫 SizeChanged에서 InitWpfUI 실행. COM 호스팅 특성상 Handle 생성 시 Size가 0x0이므로 이 지연 패턴이 필수
5. **CoordinateConverter 폴백 금지**: PositionShapesAtCursor에 중앙 폴백을 추가하면 정상 동작하는 좌표 변환까지 덮어씌움. "Illegal value" 에러는 PPT 뷰 상태에 따른 간헐 이슈이므로 기존 non-fatal catch-and-log 유지
5. **썸네일 캐시 버전**: shape-only export로 변경 후 반드시 캐시 삭제 필요. 안 그러면 이전 slide-level 썸네일이 계속 사용됨 (수정일 비교로 방지하지만, 같은 pptx면 재생성 안 함)
6. **로그 인코딩**: `debug.log`가 UTF-8인데, 한글 에러 메시지가 깨질 수 있음 (EUC-KR COM 에러)

---

## 9. 참조 경로

| 항목 | 경로 |
|------|------|
| 소스 코드 | `C:\Projects\teamppt-addin\src\TeampptAddin\` |
| 빌드 출력 | `bin\Debug\TeampptAddin.dll` |
| Assets 원본 (pptx) | `assets\headers\header_N.pptx` |
| Assets 빌드 복사본 | `bin\Debug\Assets\header_N.pptx` |
| 썸네일 캐시 | `%LocalAppData%\TeampptAddin\thumbnails\` |
| 디버그 로그 | `%LocalAppData%\TeampptAddin\debug.log` |
| 인계문서 | `C:\Projects\teamppt-addin\HANDOFF.md` |
| Add-in 레지스트리 | `HKCU\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect` |
| COM CLSID (Host) | `HKCR\CLSID\{2D4E6F8A-1B3C-5D7E-9F0A-4C6E8D2B1A3F}` |
| COM CLSID (Connect) | `HKCR\CLSID\{7B3A4D1E-9F2C-4A85-B6D0-3E8F1C5A7B92}` |

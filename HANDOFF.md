# TEAMPPT Add-in 개발 인계문서

> 최종 업데이트: 2026-06-17  
> 프로젝트 경로: `C:\Projects\teamppt-addin\src\TeampptAddin`

---

## 1. 프로젝트 개요

PowerPoint COM Add-in (.NET Framework 4.8). 오른쪽 Task Pane에 헤더 에셋 카드 목록을 표시하고, 클릭 또는 드래그앤드롭으로 현재 슬라이드에 Shape을 삽입하는 도구.

### 핵심 기능
- **카드 리스트**: `Assets/header_1.pptx` ~ `header_7.pptx`를 읽어서 카드로 표시
- **썸네일**: Shape-only PNG export (투명 배경) → 캐시 → 카드/고스트에 사용
- **클릭 삽입**: 카드 클릭 → `ShapeInserter.InsertToActiveSlide` → 슬라이드 중앙에 삽입
- **드래그앤드롭**: 마우스 캡처 방식 (OLE DragDrop 아님) → 고스트 윈도우 → 드롭 위치에 삽입

---

## 2. 파일 구조 및 역할

> 2026-06-18 모듈화 완료. Core/UI/Models/Interop 분리.

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

### UI/ — WinForms 전용 (Phase 3에서 WPF XAML로 교체 예정)

| 파일 | 역할 |
|------|------|
| `TaskPaneHost.cs` | COM 호스팅 껍데기 (InitUI, LoadCards, 썸네일 캐시, IObjectSafety) |
| `CardControl.cs` | 카드 렌더링 (OnPaint, 호버) |
| `DragHandler.cs` | 드래그 상태 관리 (BeginDrag/EndDrag, GhostWindow, 클릭 삽입) |
| `GhostWindow.cs` | WS_EX_LAYERED 투명 윈도우 (UpdateLayeredWindow P/Invoke) |

### Models/

| 파일 | 역할 |
|------|------|
| `HeaderAsset.cs` | 에셋 데이터 모델 (System.Drawing.Image 기반, WPF 의존성 제거) |

### Interop/

| 파일 | 역할 |
|------|------|
| `IObjectSafety.cs` | COM IObjectSafety 인터페이스 정의 |

---

## 3. 아키텍처 핵심 결정사항

### 3.1 WPF → WinForms 전환
- **원인**: `TaskPaneHost` 생성자에서 WPF `ElementHost` 초기화 시 COM 컨텍스트에서 "지정한 ActiveX 컨트롤을 만들 수 없습니다" 에러 발생
- **해결**: 순수 WinForms UserControl + 커스텀 `OnPaint`로 전면 전환
- **결과**: COM 호스팅 안정적 동작 확인

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

### 4.1 빌드
```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "C:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.sln" /p:Configuration=Debug /v:minimal
```
> **주의**: PowerPoint가 DLL을 잠그므로 빌드 전 반드시 PPT 종료

### 4.2 COM 등록 (관리자 권한 필요)
```powershell
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" `
  "C:\Projects\teamppt-addin\src\TeampptAddin\bin\Debug\TeampptAddin.dll" /codebase /tlb

# Control 카테고리 수동 등록 (필수!)
New-Item -Path "Registry::HKEY_CLASSES_ROOT\CLSID\{2D4E6F8A-1B3C-5D7E-9F0A-4C6E8D2B1A3F}\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}" -Force
```

### 4.3 PowerPoint Add-in 레지스트리
```powershell
$regPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect"
New-Item -Path $regPath -Force
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "TEAMPPT"
Set-ItemProperty -Path $regPath -Name "Description" -Value "TEAMPPT Header Assets Add-in"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Type DWord -Value 3
```

### 4.4 썸네일 캐시 클리어 (shape-only 재생성 시)
```powershell
Get-ChildItem "$env:LOCALAPPDATA\TeampptAddin\thumbnails" -File |
  ForEach-Object { Remove-Item $_.FullName -Force }
```

### 4.5 전체 원커맨드 배포
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

## 6. 현재 동작 상태 (2026-06-18 모듈화 후)

| 기능 | 상태 | 비고 |
|------|------|------|
| Task Pane 표시 | **동작** | COM 호스팅 OK |
| 카드 리스트 렌더링 | **동작** | 커스텀 OnPaint, 다크 테마 |
| 썸네일 (shape-only PNG) | **동작** | Group→Export→Ungroup, 캐시 |
| 클릭 삽입 | **동작** | InsertToActiveSlide |
| 드래그앤드롭 | **동작** | 마우스 캡처 + 클립보드 Paste |
| 고스트 윈도우 | **동작** | UpdateLayeredWindow, 실제 사이즈 |
| 드롭 위치 삽입 | **동작** | PointsToScreenPixels 역변환, 슬라이드 좌표 클램핑 |
| 호버 애니메이션 | **부분** | 색상 변경만 (스케일 애니메이션 미구현) |

---

## 7. 다음 실행사항 (TODO)

### Phase 1 완료: 모듈화 (2026-06-18)
- Core/UI/Models/Interop 분리 완료
- 미사용 파일 제거 (TaskPaneControl.cs)
- WPF 참조 제거 (PresentationCore, PresentationFramework 등)
- HeaderAsset WPF→System.Drawing 전환 완료
- 빌드 검증 완료 (에러/경고 0)

### Phase 2: VS 2022 + VSTO 전환
1. VSTO 프로젝트 생성 (PowerPoint Add-in)
2. Core/ 파일 그대로 복사
3. COM 수동 등록 → VSTO 자동 배포로 전환

### Phase 3: WPF XAML UI 재작성
1. UI/ 폴더를 WPF XAML로 교체 (Core/는 그대로 유지)
2. VSTO의 WPF 공식 지원으로 ElementHost 문제 없음

### Phase 4: AI 채팅 기능 추가
1. Claude API 연동

### 기존 TODO (우선순위 순)
1. **UI 최종 확인**: 썸네일 캐시 클리어 후 shape-only PNG가 투명 배경으로 정확히 나오는지 확인
2. **드롭 위치 정확도 테스트**: 여러 줌 레벨/슬라이드 크기에서 좌표 변환 정확도 검증
3. **다중 모니터 / DPI 스케일링**: `PointsToScreenPixelsX/Y` 좌표 변환 테스트
4. **Assets 폴더 자동 포함**: csproj에 빌드 시 Assets 복사 설정 추가
5. **에셋 동적 로딩**: 하드코딩된 `header_1~7` 대신 Assets 폴더 자동 스캔
6. **Strong Name 서명**: RegAsm 경고 해결
7. **호버 스케일 애니메이션**: WinForms Timer로 카드 확대/축소
10. **인스톨러 제작**: MSI 또는 ClickOnce로 원클릭 설치

---

## 8. 알려진 이슈 / 주의사항

1. **PowerPoint가 DLL 잠금**: 빌드 전 반드시 PPT 종료. 안 그러면 MSB3027 에러
2. **RegAsm unsigned 경고**: 기능에는 영향 없음. Strong Name 서명으로 해결 가능
3. **Control 카테고리 누락**: RegAsm 재실행 시 `{40FC6ED4-...}` 카테고리가 사라질 수 있음. 항상 수동 추가 필요
4. **OnSizeChanged에서 카드 로드**: `Width > 0`인 첫 SizeChanged에서만 LoadCards 실행. COM 호스팅 특성상 Handle 생성 시 Size가 0x0이므로 이 패턴이 필수
5. **썸네일 캐시 버전**: shape-only export로 변경 후 반드시 캐시 삭제 필요. 안 그러면 이전 slide-level 썸네일이 계속 사용됨 (수정일 비교로 방지하지만, 같은 pptx면 재생성 안 함)
6. **로그 인코딩**: `debug.log`가 UTF-8인데, 한글 에러 메시지가 깨질 수 있음 (EUC-KR COM 에러)

---

## 9. 참조 경로

| 항목 | 경로 |
|------|------|
| 소스 코드 | `C:\Projects\teamppt-addin\src\TeampptAddin\` |
| 빌드 출력 | `bin\Debug\TeampptAddin.dll` |
| Assets (pptx) | `bin\Debug\Assets\header_N.pptx` |
| 썸네일 캐시 | `%LocalAppData%\TeampptAddin\thumbnails\` |
| 디버그 로그 | `%LocalAppData%\TeampptAddin\debug.log` |
| 인계문서 | `C:\Projects\teamppt-addin\HANDOFF.md` |
| Add-in 레지스트리 | `HKCU\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect` |
| COM CLSID (Host) | `HKCR\CLSID\{2D4E6F8A-1B3C-5D7E-9F0A-4C6E8D2B1A3F}` |
| COM CLSID (Connect) | `HKCR\CLSID\{7B3A4D1E-9F2C-4A85-B6D0-3E8F1C5A7B92}` |

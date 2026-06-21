# 패널 버튼화 + 창별 중복 본질 해결 — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 우측 TEAMPPT 패널을 시작 시 자동 표시에서 **리본 전용 탭의 토글 버튼**으로 전환하고, `Dictionary<HWND,_CustomTaskPane>`로 **창마다 1개**를 소유·추적·해제하여 누적 중복을 구조적으로 제거한다.

**Architecture:** `Connect`는 PowerPoint 콜백/이벤트를 받아 전부 `TaskPaneManager`(신규)에 위임만 한다. `TaskPaneManager`가 "창마다 1개"의 모든 진실(팩토리·리본·`Dictionary<int HWND,_CustomTaskPane>`)을 소유하고 4개 메서드(`Toggle`/`IsVisible`/`SweepClosedWindows`/`ReleaseAll`)만 노출한다. 회수 판단은 순수함수 `WindowSweep.ToReclaim`으로 분리해 TDD한다. `TaskPaneHost`는 무변경(지연 초기화 유지 = ActiveX 충돌 구조적 보장).

**Tech Stack:** .NET Framework 4.8 COM Add-in, Microsoft.Office.Core (`IRibbonExtensibility`/`ICustomTaskPaneConsumer`), Microsoft.Office.Interop.PowerPoint, xUnit (순수함수 테스트).

## Global Constraints

- **의존성 추가 금지** — Newtonsoft.Json + Office Interop 외 새 패키지 불가.
- **`TaskPaneHost` 무변경** — 지연 초기화(`OnSizeChanged(Width>0)`) 유지 = ActiveX 충돌 구조적 보장.
- **`LoadBehavior=3` 유지** — 리본 버튼 표시 위해 자동 로드는 유지, 자동 *생성*만 제거. 레지스트리 변경 없음.
- **딕셔너리 키 = `DocumentWindow.HWND`(int)** — 창마다 고유·안정, COM 객체 직접 키 금지.
- **해제 3종 세트** — CTP 회수 시 항상 `ctp.Delete()` + `Marshal.ReleaseComObject(ctp)` + `dict.Remove(hwnd)` 세 가지 모두.
- **COM 등록 빌드는 관리자 권한 필수** — `Start-Process -Verb RunAs` (`RegisterForComInterop=true`). 단위테스트 빌드는 관리자 불필요(아래 명령 참조).
- **⚠️ 빌드 메커니즘 실측(Task 0):** CLAUDE.md의 `Start-Process -Verb RunAs cmd /c "MSBuild ... > build.log"` elevated 래퍼는 **이 세션에서 실제로 안 돌았다**(build.log 갱신 안 됨 → 스테일 로그 오판). **검증된 작동 명령:** `MSBuild TeampptAddin.csproj /t:Rebuild /p:Configuration=Debug /p:Platform=AnyCPU /p:RegisterForComInterop=false /v:minimal` (비관리자, 이 셸 직접 실행 → DLL 갱신). 이번 작업은 **새 COM coclass/GUID가 없음**(Connect에 `IRibbonExtensibility` 구현 추가 = 런타임 QI로 발견, 재등록 불필요; `TaskPaneManager`는 COM-invisible). → **초기 1회 등록(이미 완료)** 후엔 직접 recompile만으로 새 코드가 로드됨. 빌드 후 DLL `LastWriteTime`이 갱신됐는지 반드시 확인.
- **검증 시 PowerPoint 완전히 닫고 시작** — 기존 로드 인스턴스에 붙으면 패널 중복 + COM 충돌.
- **신규 `.cs`는 `TeampptAddin.csproj`의 `<ItemGroup>`에 `<Compile Include>`로 등록**해야 빌드에 포함됨(SDK 스타일 아님 — 자동 포함 안 됨).
- 노출 API 4종(타입 고정): `void Toggle(int hwnd, bool pressed)`, `bool IsVisible(int hwnd)`, `void SweepClosedWindows()`, `void ReleaseAll()`. 추가로 `void SetFactory(ICTPFactory)`, `void SetRibbon(IRibbonUI)`.
- 리본 컨트롤 ID 고정: 탭 `teampptTab`, 토글 버튼 `teampptToggle`.

---

## File Structure

| 파일 | 책임 | 변경 |
|---|---|---|
| `src/TeampptAddin/Core/WindowSweep.cs` | 순수함수: (추적 HWND 집합, 살아있는 HWND 집합) → 회수할 HWND 집합 | 신규 |
| `src/TeampptAddin.Tests/WindowSweepTest.cs` | `WindowSweep.ToReclaim` 단위테스트 4케이스 | 신규 |
| `src/TeampptAddin/UI/TaskPaneManager.cs` | 패널 생명주기 전담. 팩토리·리본·`Dictionary<int,_CustomTaskPane>` 소유. 4+2 메서드. | 신규 |
| `src/TeampptAddin/Connect.cs` | COM 진입점. `IRibbonExtensibility` 추가, 리본 콜백·앱 이벤트 → Manager 위임. 자동 생성 제거. | 수정 |
| `src/TeampptAddin/TeampptAddin.csproj` | 신규 `.cs` 2개 `<Compile Include>` 등록 | 수정 |
| `src/TeampptAddin/UI/TaskPaneHost.cs` | (무변경 — 비목표) | — |

---

### Task 0: 진단 확정 — systematic-debugging (계측·실측, 구현 선행)

> **REQUIRED SUB-SKILL: superpowers:systematic-debugging.** 구현 코드를 손대기 전에 spec §2 진단(CTPFactoryAvailable이 창마다 호출 + 필드 1개 덮어쓰기 + 해제 없음 = 누적 중복)을 `debug.log` 증거로 확정한다. 동작 변경 없음 — **로깅(계측)만** 추가한다.

**Files:**
- Modify: `src/TeampptAddin/Connect.cs:62-80` (CTPFactoryAvailable에 로그만 추가), `src/TeampptAddin/Connect.cs:35-40` (OnConnection에 임시 WindowActivate 구독)

**Interfaces:**
- Consumes: 기존 `Globals.Application`, `Logger.Log`.
- Produces: 없음(계측 전용, 이후 Task에서 대체됨).

- [ ] **Step 1: CTPFactoryAvailable에 계측 로그 추가 (동작 변경 없음)**

`src/TeampptAddin/Connect.cs`의 `CTPFactoryAvailable` 본문 try 첫 줄에 추가:

```csharp
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
{
    try
    {
        int wins = 0;
        try { wins = Globals.Application?.Windows?.Count ?? -1; } catch { wins = -2; }
        Logger.Log($"DIAG CTPFactoryAvailable called. Windows.Count={wins}");

        _taskPane = CTPFactoryInst.CreateCTP(
            "TeampptAddin.TaskPaneHost",
            "TEAMPPT");
        _taskPane.Width = 660;
        _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        _taskPane.Visible = true;
    }
    catch (Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(
            $"TEAMPPT Task Pane 생성 실패:\n{ex.Message}",
            "TEAMPPT", System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}
```

- [ ] **Step 2: OnConnection에 임시 WindowActivate 계측 구독 (CTP↔활성창 바인딩 확인용)**

`src/TeampptAddin/Connect.cs`의 `OnConnection` 본문에 추가:

```csharp
public void OnConnection(object Application, ext_ConnectMode ConnectMode,
    object AddInInst, ref Array custom)
{
    _app = (PowerPoint.Application)Application;
    Globals.Application = _app;

    _app.WindowActivate += (pres, win) =>
    {
        int hwnd = 0;
        try { hwnd = win.HWND; } catch { }
        Logger.Log($"DIAG WindowActivate. HWND={hwnd}");
    };
}
```

- [ ] **Step 3: 관리자 권한으로 COM 등록 빌드**

Run (관리자 권한 — `Start-Process -Verb RunAs`):

```powershell
Start-Process -FilePath "cmd.exe" -ArgumentList '/c "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.sln" /t:Build /p:Configuration=Debug "/p:Platform=Any CPU" /verbosity:minimal > "c:\Projects\teamppt-addin\build.log" 2>&1' -Verb RunAs -Wait -WindowStyle Hidden
```

Expected: `build.log` 끝 5줄에 `0 Error(s)`.

- [ ] **Step 4: 사용자 수동 실측 — 시나리오 실행 후 debug.log 첨부 요청**

> 빌드 산출물이 자동 등록되지 않으면 HANDOFF.md §4.3 RegAsm 절차 수행. PowerPoint를 **완전히 닫은 상태**에서 시작.

사용자에게 다음 시나리오를 부탁하고 `%LOCALAPPDATA%\TeampptAddin\debug.log`를 붙여달라고 요청:
1. PowerPoint 새로 실행 → 빈 프레젠테이션 1개.
2. 새 프레젠테이션 추가(Ctrl+N)로 둘째 창 열기.
3. 두 창을 번갈아 클릭(활성 전환) 2~3회.

- [ ] **Step 5: 진단 확정 — 증거로 spec §2 검증**

`debug.log`에서 확인:
- `DIAG CTPFactoryAvailable called`가 **창 생성마다 반복**되는가? + 각 호출에서 `Constructor STA`가 함께 찍히는가? → §2-(1) "창마다 새 CTP 생성" 확정.
- `DIAG WindowActivate. HWND=...`의 HWND가 활성창마다 **다른 값**으로 바뀌는가? → CTP↔활성창은 HWND로 구분 가능함 확정(딕셔너리 키 결정 근거).

확정 결과를 `docs/superpowers/specs/2026-06-21-panel-button-per-window-design.md` §2 아래 또는 PROGRESS-BOARD에 1~2줄로 기록. **진단이 spec과 다르면 여기서 멈추고 사용자와 재논의**(systematic-debugging: 가정 깨지면 멈춤).

- [ ] **Step 6: 계측 코드 커밋 (이후 Task에서 대체됨)**

```bash
git add src/TeampptAddin/Connect.cs
git commit -m "chore(diag): CTPFactoryAvailable/WindowActivate 계측 로그로 누적중복 진단 실측"
```

---

### Task 1: WindowSweep 순수함수 + 단위테스트 (TDD)

**Files:**
- Create: `src/TeampptAddin/Core/WindowSweep.cs`
- Test: `src/TeampptAddin.Tests/WindowSweepTest.cs`
- Modify: `src/TeampptAddin/TeampptAddin.csproj` (Compile 등록)

**Interfaces:**
- Produces: `static IReadOnlyList<int> WindowSweep.ToReclaim(IEnumerable<int> tracked, IEnumerable<int> live)` — `tracked`(추적 중 HWND) 중 `live`(살아있는 창 HWND)에 없는 것들을 반환. 입력 순서 보존, 중복 제거.

- [ ] **Step 1: 실패하는 테스트 작성**

Create `src/TeampptAddin.Tests/WindowSweepTest.cs`:

```csharp
using System.Collections.Generic;
using Xunit;

namespace TeampptAddin.Tests
{
    public class WindowSweepTest
    {
        [Fact]
        public void AllAlive_ReclaimsNothing()
        {
            var result = WindowSweep.ToReclaim(new[] { 1, 2, 3 }, new[] { 1, 2, 3 });
            Assert.Empty(result);
        }

        [Fact]
        public void OneClosed_ReclaimsThatOne()
        {
            var result = WindowSweep.ToReclaim(new[] { 1, 2, 3 }, new[] { 1, 3 });
            Assert.Equal(new[] { 2 }, result);
        }

        [Fact]
        public void SeveralClosed_ReclaimsThem_InTrackedOrder()
        {
            var result = WindowSweep.ToReclaim(new[] { 1, 2, 3, 4 }, new[] { 3 });
            Assert.Equal(new[] { 1, 2, 4 }, result);
        }

        [Fact]
        public void EmptyTracked_ReclaimsNothing()
        {
            Assert.Empty(WindowSweep.ToReclaim(new int[0], new[] { 1, 2 }));
        }
    }
}
```

- [ ] **Step 2: 테스트 실패 확인**

Run (단위테스트 — 관리자 불필요):

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.csproj" /t:Build /p:Configuration=Debug /p:Platform=AnyCPU /p:RegisterForComInterop=false /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" /t:Build /p:Configuration=Debug /p:BuildProjectReferences=false /v:minimal
dotnet test "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" --no-build --no-restore --filter "FullyQualifiedName~WindowSweepTest"
```

Expected: 컴파일 실패(`WindowSweep` 형식이 없음, CS0103/CS0246).

- [ ] **Step 3: 최소 구현 작성**

Create `src/TeampptAddin/Core/WindowSweep.cs`:

```csharp
using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    /// <summary>회수 판단 순수함수. 추적 중 HWND 중 살아있지 않은 것을 골라낸다.</summary>
    public static class WindowSweep
    {
        public static IReadOnlyList<int> ToReclaim(IEnumerable<int> tracked, IEnumerable<int> live)
        {
            var liveSet = new HashSet<int>(live ?? Enumerable.Empty<int>());
            var seen = new HashSet<int>();
            var result = new List<int>();
            foreach (var hwnd in tracked ?? Enumerable.Empty<int>())
            {
                if (liveSet.Contains(hwnd)) continue;
                if (!seen.Add(hwnd)) continue;
                result.Add(hwnd);
            }
            return result;
        }
    }
}
```

- [ ] **Step 4: csproj에 Compile 등록**

`src/TeampptAddin/TeampptAddin.csproj`의 `<Compile Include="Core\CoordinateConverter.cs" />` 다음 줄에 추가:

```xml
    <Compile Include="Core\WindowSweep.cs" />
```

- [ ] **Step 5: 테스트 통과 확인**

Run (Step 2와 동일 3줄):

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.csproj" /t:Build /p:Configuration=Debug /p:Platform=AnyCPU /p:RegisterForComInterop=false /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" /t:Build /p:Configuration=Debug /p:BuildProjectReferences=false /v:minimal
dotnet test "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" --no-build --no-restore --filter "FullyQualifiedName~WindowSweepTest"
```

Expected: `Passed!  - Failed: 0, Passed: 4`.

- [ ] **Step 6: 커밋**

```bash
git add src/TeampptAddin/Core/WindowSweep.cs src/TeampptAddin.Tests/WindowSweepTest.cs src/TeampptAddin/TeampptAddin.csproj
git commit -m "feat(panel): 회수판단 순수함수 WindowSweep + 단위테스트 4 GREEN"
```

---

### Task 2: TaskPaneManager — 생명주기 전담 (창마다 1개 + 해제 3종 세트)

**Files:**
- Create: `src/TeampptAddin/UI/TaskPaneManager.cs`
- Modify: `src/TeampptAddin/TeampptAddin.csproj` (Compile 등록)

**Interfaces:**
- Consumes: `WindowSweep.ToReclaim` (Task 1), `Globals.Application`, `Logger.Log`, `Microsoft.Office.Core.ICTPFactory`/`IRibbonUI`/`_CustomTaskPane`/`MsoCTPDockPosition`.
- Produces (Connect가 호출):
  - `void SetFactory(ICTPFactory factory)` — 팩토리 보관(생성 안 함).
  - `void SetRibbon(IRibbonUI ribbon)` — 리본 참조 보관(버튼 갱신용).
  - `void Toggle(int hwnd, bool pressed)` — 없고 pressed면 생성·등록·표시, 있으면 `Visible=pressed`.
  - `bool IsVisible(int hwnd)` — dict에 있고 `Visible==true`.
  - `void SweepClosedWindows()` — 사라진 창 CTP를 해제 3종 세트로 회수 + 버튼 무효화.
  - `void ReleaseAll()` — 전량 해제 + clear.

- [ ] **Step 1: TaskPaneManager 작성**

Create `src/TeampptAddin/UI/TaskPaneManager.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace TeampptAddin
{
    /// <summary>
    /// 패널 생명주기 전담. "창마다 CTP 1개"의 모든 진실을 소유한다.
    /// 키 = DocumentWindow.HWND(int). 해제는 항상 3종 세트:
    /// Delete() + Marshal.ReleaseComObject + dict.Remove.
    /// </summary>
    public sealed class TaskPaneManager
    {
        private const string ToggleControlId = "teampptToggle";

        private readonly Dictionary<int, _CustomTaskPane> _panes
            = new Dictionary<int, _CustomTaskPane>();

        private ICTPFactory _factory;
        private IRibbonUI _ribbon;

        public void SetFactory(ICTPFactory factory)
        {
            _factory = factory;
            Logger.Log("Manager.SetFactory");
        }

        public void SetRibbon(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Logger.Log("Manager.SetRibbon");
        }

        public bool IsVisible(int hwnd)
        {
            if (hwnd != 0 && _panes.TryGetValue(hwnd, out var pane))
            {
                try { return pane.Visible; } catch { return false; }
            }
            return false;
        }

        public void Toggle(int hwnd, bool pressed)
        {
            if (hwnd == 0) { Logger.Log("Manager.Toggle ignored: hwnd=0"); return; }

            if (_panes.TryGetValue(hwnd, out var existing))
            {
                try { existing.Visible = pressed; } catch (Exception ex) { Logger.Log($"Toggle visible failed: {ex.Message}"); }
                Logger.Log($"Manager.Toggle existing hwnd={hwnd} pressed={pressed}");
                return;
            }

            if (!pressed) { Logger.Log($"Manager.Toggle no-op hwnd={hwnd} (none, pressed=false)"); return; }
            if (_factory == null) { Logger.Log("Manager.Toggle abort: factory null"); return; }

            try
            {
                var pane = _factory.CreateCTP("TeampptAddin.TaskPaneHost", "TEAMPPT");
                pane.Width = 660;
                pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _panes[hwnd] = pane;
                SubscribeVisibleStateChange(pane);
                pane.Visible = true;
                Logger.Log($"Manager.Toggle created hwnd={hwnd}. count={_panes.Count}");
            }
            catch (Exception ex)
            {
                Logger.Log($"Manager.Toggle create FAILED hwnd={hwnd}: {ex}");
            }
        }

        public void SweepClosedWindows()
        {
            var live = new List<int>();
            try
            {
                var windows = Globals.Application?.Windows;
                if (windows != null)
                {
                    foreach (PowerPoint.DocumentWindow w in windows)
                    {
                        try { live.Add(w.HWND); } catch { }
                    }
                }
            }
            catch (Exception ex) { Logger.Log($"Sweep enum failed: {ex.Message}"); }

            var reclaim = WindowSweep.ToReclaim(new List<int>(_panes.Keys), live);
            foreach (var hwnd in reclaim)
            {
                ReleaseOne(hwnd);
                Logger.Log($"Manager.Sweep reclaimed hwnd={hwnd}. count={_panes.Count}");
            }
            if (reclaim.Count > 0) InvalidateButton();
        }

        public void ReleaseAll()
        {
            foreach (var hwnd in new List<int>(_panes.Keys))
                ReleaseOne(hwnd);
            _panes.Clear();
            Logger.Log("Manager.ReleaseAll done");
        }

        // 해제 3종 세트: Delete() + ReleaseComObject + dict.Remove
        private void ReleaseOne(int hwnd)
        {
            if (!_panes.TryGetValue(hwnd, out var pane)) return;
            try { pane.Delete(); } catch (Exception ex) { Logger.Log($"ReleaseOne Delete failed hwnd={hwnd}: {ex.Message}"); }
            try { Marshal.ReleaseComObject(pane); } catch { }
            _panes.Remove(hwnd);
        }

        private void SubscribeVisibleStateChange(_CustomTaskPane pane)
        {
            try
            {
                ((_CustomTaskPaneEvents_Event)pane).VisibleStateChange +=
                    _ => InvalidateButton();
            }
            catch (Exception ex) { Logger.Log($"SubscribeVisibleStateChange failed: {ex.Message}"); }
        }

        private void InvalidateButton()
        {
            try { _ribbon?.InvalidateControl(ToggleControlId); }
            catch (Exception ex) { Logger.Log($"InvalidateButton failed: {ex.Message}"); }
        }
    }
}
```

> **참고(Task 0 확인 대상):** `_CustomTaskPaneEvents_Event` 캐스트로 `VisibleStateChange`를 구독한다(Office PIA가 동일 객체에 이 이벤트 소스 인터페이스를 노출). Task 0에서 CTP 이벤트 바인딩이 다르게 확인되면 이 구독부만 조정하고 나머지 구조는 유지.

- [ ] **Step 2: 파일 상단 PowerPoint alias 확인**

위 코드는 `PowerPoint.DocumentWindow`를 쓴다. 파일 최상단 using에 다음이 포함됐는지 확인(없으면 추가):

```csharp
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
```

- [ ] **Step 3: csproj에 Compile 등록**

`src/TeampptAddin/TeampptAddin.csproj`의 `<Compile Include="UI\TaskPaneHost.cs" />` 다음 줄에 추가:

```xml
    <Compile Include="UI\TaskPaneManager.cs" />
```

- [ ] **Step 4: 컴파일 검증 (관리자 불필요, 등록 없이 빌드만)**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.csproj" /t:Build /p:Configuration=Debug /p:Platform=AnyCPU /p:RegisterForComInterop=false /v:minimal
```

Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 5: 커밋**

```bash
git add src/TeampptAddin/UI/TaskPaneManager.cs src/TeampptAddin/TeampptAddin.csproj
git commit -m "feat(panel): TaskPaneManager 신규 — 창별 CTP 추적/해제 3종 세트"
```

---

### Task 3: Connect 리본 버튼화 + 위임 (자동 생성 제거)

> 이 Task의 deliverable: 시작 시 자동으로 안 뜨고, **리본 TEAMPPT 탭의 토글 버튼**으로 활성창에 패널 1개를 켜고 끈다. (수동검증 체크리스트 1~4)

**Files:**
- Modify: `src/TeampptAddin/Connect.cs` (전면 — `IRibbonExtensibility` 추가, 리본 콜백, Manager 위임, 자동 생성 제거)

**Interfaces:**
- Consumes: `TaskPaneManager`(Task 2)의 `SetFactory`/`SetRibbon`/`Toggle`/`IsVisible`/`ReleaseAll`.
- Produces (PowerPoint가 리본 XML 콜백으로 호출): `string GetCustomUI(string)`, `void OnRibbonLoad(IRibbonUI)`, `void OnToggleAction(IRibbonControl, bool)`, `bool GetTogglePressed(IRibbonControl)`.

- [ ] **Step 1: Connect.cs 전면 교체**

Replace 전체 `src/TeampptAddin/Connect.cs` (Task 0 계측 코드 대체):

```csharp
using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// COM Add-in 진입점. 직접 패널을 만들지 않고 전부 TaskPaneManager에 위임한다.
    /// - IDTExtensibility2: 생명주기. OnConnection에서 앱 이벤트 구독.
    /// - ICustomTaskPaneConsumer: CTPFactoryAvailable에서 팩토리를 Manager에 보관(생성 안 함).
    /// - IRibbonExtensibility: TEAMPPT 탭 토글 버튼. 버튼이 활성창 패널을 토글.
    /// LoadBehavior=3 유지(리본 표시). 자동 생성만 제거.
    /// </summary>
    [ComVisible(true)]
    [Guid("7B3A4D1E-9F2C-4A85-B6D0-3E8F1C5A7B92")]
    [ProgId("TeampptAddin.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2, ICustomTaskPaneConsumer, IRibbonExtensibility
    {
        private PowerPoint.Application _app;
        private readonly TaskPaneManager _manager = new TaskPaneManager();

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
            object AddInInst, ref Array custom)
        {
            _app = (PowerPoint.Application)Application;
            Globals.Application = _app;

            _app.WindowActivate += App_WindowActivate;
            _app.PresentationClose += App_PresentationClose;
            Logger.Log("Connect.OnConnection: events wired");
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (_app != null)
                {
                    _app.WindowActivate -= App_WindowActivate;
                    _app.PresentationClose -= App_PresentationClose;
                }
            }
            catch (Exception ex) { Logger.Log($"OnDisconnection unwire failed: {ex.Message}"); }

            _manager.ReleaseAll();
            _app = null;
            Globals.Application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom)
        {
            _manager.ReleaseAll();
        }

        #endregion

        #region App events

        private void App_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            _manager.SweepClosedWindows();
        }

        private void App_PresentationClose(PowerPoint.Presentation Pres)
        {
            _manager.SweepClosedWindows();
        }

        #endregion

        #region ICustomTaskPaneConsumer

        // 자동 생성 제거: 팩토리만 보관(idempotent). 패널은 버튼으로만 생성.
        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            _manager.SetFactory(CTPFactoryInst);
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            return RibbonXml;
        }

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _manager.SetRibbon(ribbon);
        }

        public void OnToggleAction(IRibbonControl control, bool pressed)
        {
            _manager.Toggle(ActiveHwnd(), pressed);
        }

        public bool GetTogglePressed(IRibbonControl control)
        {
            return _manager.IsVisible(ActiveHwnd());
        }

        private int ActiveHwnd()
        {
            try { return _app?.ActiveWindow?.HWND ?? 0; }
            catch { return 0; }
        }

        private const string RibbonXml =
@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='teampptTab' label='TEAMPPT'>
        <group id='teampptGroup' label='패널'>
          <toggleButton id='teampptToggle' label='TEAMPPT 패널'
                        size='large' imageMso='PaneInsert'
                        onAction='OnToggleAction' getPressed='GetTogglePressed'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        #endregion
    }
}
```

> **참고:** `RibbonXml`의 한글 라벨은 C# 소스(UTF-8) 문자열 리터럴이라 안전(ASCII 제약은 `.ps1` 한정). `imageMso='PaneInsert'`는 내장 아이콘.

- [ ] **Step 2: COM 등록 빌드 (관리자 권한 필수)**

Run:

```powershell
Start-Process -FilePath "cmd.exe" -ArgumentList '/c "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.sln" /t:Build /p:Configuration=Debug "/p:Platform=Any CPU" /verbosity:minimal > "c:\Projects\teamppt-addin\build.log" 2>&1' -Verb RunAs -Wait -WindowStyle Hidden
```

Expected: `build.log` 끝 5줄에 `0 Error(s)`.

- [ ] **Step 3: 수동검증 체크리스트 1~4 (사용자 실행 → debug.log/관찰 첨부)**

> PowerPoint 완전히 닫고 시작. 필요 시 HANDOFF.md §4.3 RegAsm 재등록.

1. PowerPoint 시작 → **패널 자동으로 안 뜸**, 리본에 **TEAMPPT 탭 + 토글 버튼** 보임.
2. 한 창에서 버튼 클릭 → 우측 패널 1개 생성. 또 클릭 → 숨김, 또 → 표시(버튼 눌림 상태 일치).
3. 둘째 창 열기(Ctrl+N) → 둘째 창엔 패널 **자동으로 안 뜸**(버튼 꺼짐).
4. 둘째 창 버튼 클릭 → 그 창에 독립 패널, 두 창 패널 각각 작동.

`debug.log`에 `Manager.Toggle created`가 창마다 다른 hwnd로 1번씩, 중복 생성 로그 없음을 확인.

- [ ] **Step 4: 커밋**

```bash
git add src/TeampptAddin/Connect.cs
git commit -m "feat(panel): 리본 토글 버튼화 + Manager 위임, 자동 생성 제거"
```

---

### Task 4: 동기화·회수 견고화 검증 (창 전환/닫기/종료)

> 이 Task의 deliverable: 창 전환 시 버튼이 활성창 상태를 따르고, 패널 X 닫기·창 닫기·PPT 종료에서 추적/해제가 정확. (수동검증 체크리스트 5~10, 본질 해결의 최종 증거 8·9·10)
>
> 코드는 Task 2·3에서 이미 배선됨(`SweepClosedWindows`는 `WindowActivate`/`PresentationClose`에, `InvalidateButton`은 `VisibleStateChange`/sweep에, `ReleaseAll`은 종료에). 이 Task는 **그 배선이 실제 PowerPoint에서 명세대로 동작하는지 검증**하고, 어긋나면 systematic-debugging으로 고친다.

**Files:**
- Modify (필요 시에만): `src/TeampptAddin/UI/TaskPaneManager.cs`, `src/TeampptAddin/Connect.cs` (검증 중 발견된 동기화 결함 수정)

**Interfaces:**
- Consumes: Task 2·3의 전체 표면. 새 공개 API 없음.

- [ ] **Step 1: 수동검증 체크리스트 5~10 (사용자 실행 → debug.log 첨부)**

5. 두 창 번갈아 활성화 → **버튼 눌림 상태가 활성창의 패널 표시 여부를 따라감**(getPressed 재평가).
6. 패널을 **X로 닫기** → 같은 창 버튼이 "꺼짐"으로 동기화(VisibleStateChange → InvalidateControl).
7. 패널 켠 창 하나를 닫기 → **그 창 패널만 회수**, 다른 창 패널 멀쩡(`Manager.Sweep reclaimed` 1건).
8. 창을 **여러 번 껐다 켰다** 반복 → 패널 **누적 0**(잔존 없음). ← 본질 해결 축①
9. 패널 열 때 **ActiveX 충돌/크래시 없음**(`Constructor STA` 정상, 예외 로그 없음). ← 본질 해결 축②
10. PowerPoint 종료 → `Manager.ReleaseAll done`, COM 누수/에러 없음.

- [ ] **Step 2: 증거 판정 — 어긋나면 systematic-debugging**

`debug.log`로 각 항목 증거 확인:
- 5/6: `InvalidateButton` 호출 + 버튼 시각 상태 일치.
- 7/8: `Manager.Sweep reclaimed hwnd=...`가 닫힌 창마다 정확히 1건, `count` 단조 감소.
- 9: `Constructor STA=STA` + create FAILED 로그 없음.
- 10: `Manager.ReleaseAll done`.

**불일치 시:** 멈추고 superpowers:systematic-debugging으로 근본 원인(예: VisibleStateChange 미발화, WindowActivate 미발생)을 `debug.log`로 좁혀 최소 수정. 수정 후 해당 항목 재검증.

- [ ] **Step 3: 회귀 단위테스트 재확인 (수정이 있었다면)**

Run:

```powershell
dotnet test "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" --no-build --no-restore --filter "FullyQualifiedName~WindowSweepTest"
```

Expected: `Passed: 4`.

- [ ] **Step 4: 커밋 (수정이 있었을 때만)**

```bash
git add -A
git commit -m "fix(panel): 창 전환/닫기 동기화·회수 견고화 (수동검증 5~10 PASS)"
```

---

### Task 5: 마무리 — 문서 갱신 + 브랜치 통합

**Files:**
- Modify: `PROGRESS-BOARD.md`, `HANDOFF.md`

- [ ] **Step 1: PROGRESS-BOARD 잎 교체**

`PROGRESS-BOARD.md`의 나무/잎을 갱신: 이 plan 완료 표시, 나무를 "B 에셋 인제스트 재개(LLM 이해 어댑터)"로 교체. (CLAUDE.md 규칙: 끝난 잎 지우고 교체, 숲 골격 유지.)

- [ ] **Step 2: HANDOFF 재개 지점 갱신**

`HANDOFF.md` §재개 지점을 "패널 버튼화 완료·main 머지"로 갱신하고 다음 할 일을 인제스트 재개로 설정.

- [ ] **Step 3: 커밋 + 통합 옵션 제시**

```bash
git add PROGRESS-BOARD.md HANDOFF.md
git commit -m "docs: 패널 버튼화 완료 반영 + 인계 갱신"
```

이후 superpowers:finishing-a-development-branch로 main 통합(merge/PR) 옵션을 사용자에게 제시.

---

## Self-Review

**Spec coverage:**
- §2 진단 → Task 0 (실측 확정). §3 버튼+딕셔너리 → Task 2·3. §4 컴포넌트표(Connect/Manager/Host/Ribbon) → Task 3(Connect+Ribbon), Task 2(Manager), Host 무변경. §4 노출 4메서드 → Manager API(+SetFactory/SetRibbon). §5 이벤트 8개 → 1:CTPFactoryAvailable(SetFactory, T3), 2:onLoad(OnRibbonLoad, T3), 3:onAction(OnToggleAction, T3), 4:getPressed(GetTogglePressed, T3), 5:WindowActivate(App_WindowActivate→Sweep+Invalidate, T3/T4), 6:VisibleStateChange(SubscribeVisibleStateChange, T2/T4), 7:PresentationClose(App_PresentationClose, T3/T4), 8:Shutdown/Disconnection(ReleaseAll, T3). §6 해제 3종 세트 → `ReleaseOne` (T2). §6 4메서드 동작 → T2. §7 ActiveX 타이밍(Host 무변경) → 비목표 준수. §8 단위테스트 순수함수 4케이스 → Task 1. §8 수동검증 1~10 → Task 3(1~4)+Task 4(5~10). §9 범위/비목표 → 준수(Host 내부·인제스트·아이콘 정교화 제외).
- 갭 없음.

**Placeholder scan:** 모든 코드 스텝에 실제 코드/명령/기대출력 포함. TBD·"적절히 처리" 없음.

**Type consistency:** `Toggle(int,bool)`/`IsVisible(int)`/`SweepClosedWindows()`/`ReleaseAll()`/`SetFactory(ICTPFactory)`/`SetRibbon(IRibbonUI)`가 Task 2 정의와 Task 3 호출에서 일치. `WindowSweep.ToReclaim(IEnumerable<int>,IEnumerable<int>)`가 Task 1 정의와 Task 2 사용에서 일치. 리본 ID `teampptToggle`이 Manager 상수·Ribbon XML에서 일치. 리본 콜백 이름(`OnRibbonLoad`/`OnToggleAction`/`GetTogglePressed`)이 XML 속성과 Connect 메서드에서 일치.

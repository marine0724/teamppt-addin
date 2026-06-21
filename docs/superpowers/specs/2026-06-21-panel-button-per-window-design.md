# 패널 버튼화 + 창별 중복 본질 해결 — 설계

> 작성: 2026-06-21 · 상태: **승인됨**(brainstorming 완료) → 다음: writing-plans
> 브랜치: **신규 별도 브랜치**(Connect.cs/TaskPaneHost 영역 → 인제스트와 분리)

## 1. 문제

우측 TEAMPPT 패널이 두 가지로 오동작:
- PowerPoint **시작 시 자동으로** 뜬다(사용자가 참조용으로 PPT를 여러 개 켜도 매번 관여).
- 창이 여럿이면 패널이 **누적 중복**되거나 엉뚱한(원래) 창에 붙는다.

## 2. 근본 원인 (코드 진단)

`Connect.cs`의 `CTPFactoryAvailable`이 트리거. PowerPoint는 이 콜백을 **프레젠테이션 창이 새로 열릴 때마다 다시 호출**한다(로그에 창마다 `Constructor STA` 반복 = 창마다 CTP 생성 증거). 현재 코드는:

1. **"이미 있나?" 체크 없이** 매번 새 CTP 생성.
2. `_CustomTaskPane _taskPane` **필드 1개에 덮어쓰기** → 이전 패널은 참조만 잃고 화면에 잔존.
3. 창 닫을 때 **해제 없음**(`OnBeginShutdown`은 `Visible=false`만, `OnDisconnection`은 패널 정리 안 함).

이 셋이 겹쳐 누적 중복. = **자동·무통제·무정리** 생성.

### 2.1 실측 확정 (Task 0, debug.log 증거 — 2026-06-21)

계측 빌드로 실측해 위 진단을 **확정**했다. 시나리오 = 에셋 .pptx 파일 4개를 연속으로 열기:

```
CTPFactoryAvailable Count=0 → Constructor(파일1) → WindowActivate HWND=1640220
CTPFactoryAvailable Count=1 → Constructor(파일2) → WindowActivate HWND=2359364
CTPFactoryAvailable Count=2 → Constructor(파일3) → WindowActivate HWND=395176
CTPFactoryAvailable Count=3 → Constructor(파일4) → WindowActivate HWND=460716
```

- **확정①:** 프레젠테이션 파일을 열 때마다 `CTPFactoryAvailable`이 다시 호출되고 매번 새 CTP(Constructor) 생성 → 필드 덮어쓰기로 이전 패널 잔존 = 누적 중복.
- **확정②:** 각 창 HWND가 고유·안정적 → 딕셔너리 키=HWND 타당. WindowActivate가 전환마다 정확히 발화 → 활성창 식별·빗질 신뢰 가능.
- **참고:** 단순 `Ctrl+N` 새 창 시나리오에선 1회만 호출됐다(트리거는 프레젠테이션 *파일 열기* 경로에서 확실히 재현). 해법(접근 A)은 호출 빈도와 무관하게 견고(생성을 버튼으로만 수행).

## 3. 채택 접근 (A): 리본 토글 버튼 + 창별 CTP 추적

버튼은 **무분별한 자동 생성을 막는 게이트**, 딕셔너리는 **창마다 1개 보장·해제**의 본질 해결. 둘을 함께 쓴다.

- 목표 동작: 시작 시 자동으로 안 뜸 → 사용자가 **각 창에서 버튼**을 눌러야 그 창에 패널 1개. **2개 이상 창에서 각각 독립 작동**.

## 4. 아키텍처 & 컴포넌트

| 구성요소 | 역할 | 상태 |
|---|---|---|
| **Connect** | COM 진입점. `IDTExtensibility2` + `ICustomTaskPaneConsumer` + **`IRibbonExtensibility`(신규)**. 직접 패널 생성 안 함 — **전부 Manager에 위임**. | 수정 |
| **TaskPaneManager** | 패널 생명주기 전담. `ICTPFactory`·`IRibbonUI` 보관 + **`Dictionary<int HWND, _CustomTaskPane>`**. | 신규 |
| **TaskPaneHost** | 패널 본체(WPF). CTP 1개당 인스턴스 1개 = 창마다 독립. **무변경**(지연 초기화 유지). | 무변경 |
| **Ribbon XML** | `customUI` → 전용 **TEAMPPT 탭** → 그룹 → **토글 버튼 1개**. | 신규 |

**모듈화 원칙:** `Connect`는 "PowerPoint 콜백을 받아 Manager에 넘김"만. `TaskPaneManager`는 "창마다 1개"의 모든 진실을 소유하고 바깥엔 4개 메서드만 노출: `Toggle(hwnd, pressed)`, `IsVisible(hwnd)`, `SweepClosedWindows()`, `ReleaseAll()`.

**딕셔너리 키 = `DocumentWindow.HWND`(int):** 창마다 고유·안정적, COM 객체 직접 키보다 비교 정확.

**레지스트리/로드 변화 없음:** `IRibbonExtensibility`는 Office가 자동 인식. `LoadBehavior=3`(자동 로드)은 **리본 버튼 표시 위해 유지** — 자동 *생성*만 제거.

## 5. 이벤트 흐름 (트리거 → 동작)

| # | 트리거 | 동작 |
|---|---|---|
| 1 | CTPFactoryAvailable (PPT, 로드/창 생성) | 팩토리 없으면 **보관만**. 생성 안 함(idempotent). |
| 2 | 리본 onLoad | `IRibbonUI` 참조 보관(버튼 갱신용). |
| 3 | 버튼 onAction(pressed) | 활성창 HWND → `Manager.Toggle`: 없으면 생성·등록·표시, 있으면 `Visible=pressed`. |
| 4 | 버튼 getPressed | `Manager.IsVisible(활성창 hwnd)` → 버튼이 현재 창 상태 반영. |
| 5 | App.WindowActivate (창 전환) | `SweepClosedWindows()` + 버튼 `InvalidateControl` → getPressed 재실행. |
| 6 | CTP VisibleStateChange (패널 X로 닫음) | 버튼 무효화 → "꺼짐" 동기화. |
| 7 | App.PresentationClose | `SweepClosedWindows()` → 사라진 창 CTP 해제. |
| 8 | OnBeginShutdown / OnDisconnection | `Manager.ReleaseAll()`. |

**핵심:** 버튼 상태는 **활성창 기준**(4·5). 패널 X로 닫아도 버튼 동기화(6).

## 6. 정리(해제) 로직 — 중복 본질 해결

- **`Toggle(hwnd, pressed)`**: dict에 없고 pressed면 CTP 생성→폭/도킹 설정→dict 등록→`VisibleStateChange` 구독→`Visible=true`. 있으면 `Visible=pressed`. "있나?" 체크가 한 창 두 개를 원천 차단.
- **`SweepClosedWindows()`**: `live = Application.Windows의 HWND`. dict의 hwnd 중 live에 없는 것 → `ctp.Delete()` + `Marshal.ReleaseComObject` + `dict.Remove`. (창 전환·프레젠테이션 닫힘마다 호출 → 잔존 즉시 회수.)
- **`IsVisible(hwnd)`**: dict에 있고 `Visible==true`.
- **`ReleaseAll()`**: 전량 `Delete()`+`ReleaseComObject`+clear. COM 누수 0.

**해제 3종 세트(빠져서 버그였음):** `Delete()`(화면·컬렉션에서 실제 제거) + `ReleaseComObject`(COM 참조) + `dict.Remove`(추적 제거).

## 7. ActiveX 충돌 타이밍 (보너스 효과)

기존 ActiveX 충돌은 `ElementHost`(WPF 인터롭)를 ActiveX 사이트 *생성 중*에 만들어 생긴 것. 해법은 [TaskPaneHost.cs](../../../src/TeampptAddin/UI/TaskPaneHost.cs)의 `OnSizeChanged(Width>0)` **지연 초기화** — 표준 패턴이라 **유지**한다(구조적 보장).

버튼 방식은 추가로, CTP 생성을 **시작 혼잡 구간 → 사용자 클릭의 평온한 시점**으로 옮겨 충돌 유발 타이밍을 제거. → 지연 초기화(보장) + 버튼(혼잡 제거) 두 겹으로 재현 조건이 사실상 안 생김. **`TaskPaneHost`는 손대지 않음.**

## 8. 테스트 전략

- **단위테스트(TDD, 순수):** 회수 판단 순수 함수 `(dict HWND 집합, live HWND 집합) → 회수할 HWND 집합`. 케이스: 전부 살아있음→0, 하나 닫힘→그 하나, 여럿 닫힘→그것들, 빈 dict→빈 결과.
- **수동검증(PowerPoint 실제, debug.log 증거):** 생성/토글/빗질/해제마다 `Logger.Log`. 체크리스트 1~10(아래). 8(창 여러 번 껐다켜기→누적 0)·9(패널 열 때 ActiveX 충돌 없음)가 본질 해결 두 축의 최종 증거.

수동검증 체크리스트:
1. 한 창 버튼 → 패널 1개
2. 또 클릭/또 → 숨김→표시(토글), 버튼표시 일치
3. 둘째 창 열기 → 자동으로 안 뜸
4. 둘째 창 버튼 → 독립 패널, 둘 다 작동
5. 창 전환 → 버튼표시가 활성창 따라감
6. 패널 X로 닫기 → 버튼 꺼짐 동기화
7. 한 창 닫기 → 그 창 패널만 회수, 다른 창 멀쩡
8. 창 여러 번 껐다켜기 → 누적 0
9. 패널 열 때 → ActiveX 충돌/크래시 없음
10. PPT 종료 → ReleaseAll, 누수/에러 없음

## 9. 범위 / 비목표 / 선행 확인

- **범위:** `Connect.cs`(리본+위임), `TaskPaneManager`(신규), Ribbon XML, 순수함수 단위테스트, 수동검증 스크립트/절차.
- **비목표:** `TaskPaneHost` 내부 변경, 인제스트/LLM/Supabase, 버튼 아이콘 정교화(텍스트/기본 imageMso로 충분).
- **선행 확인(writing-plans의 Task 0 = systematic-debugging):** 구현 전 `debug.log`로 **`CTPFactoryAvailable`이 실제로 언제 몇 번 불리는지** 실측해 진단(§2) 확정. PowerPoint의 CTP↔활성창 바인딩 동작도 확인.
- **별도 브랜치**에서 진행(인제스트와 분리).

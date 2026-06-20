# 로컬 인제스트 코어 구현 계획 (Phase B — 1번 plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 디자이너 묶음 pptx(PowerPoint 섹션 분류)를 받아 섹션을 읽고, 슬라이드 1장 = 에셋 1개 단위로 개별 pptx + 썸네일 PNG(긴 변 768px)로 쪼개는 로컬 인제스트 코어를 만든다. 외부 의존성(Supabase/LLM) 없음.

**Architecture:** 순수 로직(섹션→에셋 분할 계획, 에셋 ID 생성)은 xUnit으로 TDD. COM/Interop(섹션 읽기, 슬라이드 split, PNG 렌더)은 인터페이스 뒤의 얇은 어댑터로 두고 PowerPoint에서 수동 검증. 오케스트레이터가 이 둘을 연결. 기존 코드의 Core(COM)/Services·Models(로직) 분리를 그대로 따른다.

**Tech Stack:** .NET Framework 4.8, Microsoft.Office.Interop.PowerPoint, Newtonsoft.Json 13.0.3, xUnit 2.9.

## Global Constraints

- Core/Connect.cs/Globals.cs 직접 수정 금지 (신규 파일만 추가).
- 의존성 추가 금지 — Newtonsoft.Json + Office Interop만.
- 단위테스트는 관리자 불필요: 본프로젝트 `/p:RegisterForComInterop=false` 빌드 → 테스트프로젝트 `/p:BuildProjectReferences=false` 빌드 → `dotnet test --no-build --no-restore`.
- COM 객체는 반드시 `Marshal.ReleaseComObject` + try/finally로 해제 (기존 ThumbnailGenerator 패턴).
- PNG 렌더 해상도 상수 `LlmImageLongEdgePx = 768` (긴 변 기준). 모든 화면읽기 경로 공유.
- 에셋 단위 = 섹션 안 슬라이드 1장. 섹션명 = category.
- MSBuild = `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`.

---

## File Structure

| 파일 | 책임 | 테스트 |
|---|---|---|
| `Models/IngestModels.cs` (신규) | `SectionInfo`, `AssetSplitItem` 데이터 레코드 (순수) | 직렬화/동등성 불필요 |
| `Services/AssetIdGenerator.cs` (신규) | 섹션명+순번 → 파일安全 에셋 ID(`표지_01`) (순수) | `AssetIdGeneratorTest.cs` |
| `Services/IngestPlanner.cs` (신규) | `List<SectionInfo>` → `List<AssetSplitItem>` (순수) | `IngestPlannerTest.cs` |
| `Core/SectionReader.cs` (신규) | `Presentation.SectionProperties` → `List<SectionInfo>` (COM) | 수동 |
| `Core/SlideSplitter.cs` (신규) | 소스 pres + 슬라이드 인덱스 → 단일 슬라이드 pptx 저장 (COM) | 수동 |
| `Core/SlideImageRenderer.cs` (신규) | 슬라이드 → PNG, 긴 변 px 지정 (COM) | 수동 |
| `Services/IngestRunner.cs` (신규) | 리더→플래너→스플리터→렌더러 연결 (COM 오케스트레이션) | 수동 |

순수 3개(Models/AssetIdGenerator/IngestPlanner)가 이번 plan의 단위테스트 가능 산출물. COM 4개는 인터페이스로 분리하되 PowerPoint 수동 검증.

---

### Task 1: 인제스트 데이터 모델

**Files:**
- Create: `src/TeampptAddin/Models/IngestModels.cs`

**Interfaces:**
- Produces: `SectionInfo { string Name; int FirstSlideIndex; int SlideCount }`, `AssetSplitItem { int SourceSlideIndex; string Category; string AssetId; string PptxFileName; string ThumbFileName }` — 이후 모든 Task가 사용.

- [ ] **Step 1: 데이터 모델 작성**

```csharp
namespace TeampptAddin
{
    /// <summary>PowerPoint 섹션 1개 = 카테고리 1개. Interop SectionProperties에서 읽음.</summary>
    public class SectionInfo
    {
        public string Name { get; set; }           // 섹션명 = category (예: "표지")
        public int FirstSlideIndex { get; set; }    // 1-based, 이 섹션 첫 슬라이드
        public int SlideCount { get; set; }         // 이 섹션 슬라이드 수
    }

    /// <summary>슬라이드 1장 = 에셋 1개. 분할 계획의 한 항목.</summary>
    public class AssetSplitItem
    {
        public int SourceSlideIndex { get; set; }   // 1-based, 원본 묶음 pptx 안 위치
        public string Category { get; set; }        // 섹션명
        public string AssetId { get; set; }         // "표지_01"
        public string PptxFileName { get; set; }    // "표지_01.pptx"
        public string ThumbFileName { get; set; }   // "표지_01.png"
    }
}
```

- [ ] **Step 2: 빌드 확인**

Run (관리자 불필요):
```
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.csproj" /t:Build /p:Configuration=Debug "/p:Platform=AnyCPU" /p:RegisterForComInterop=false /verbosity:minimal
```
Expected: Build succeeded.

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Models/IngestModels.cs
git commit -m "feat(ingest): SectionInfo/AssetSplitItem 데이터 모델"
```

---

### Task 2: 에셋 ID 생성기

**Files:**
- Create: `src/TeampptAddin/Services/AssetIdGenerator.cs`
- Test: `src/TeampptAddin.Tests/AssetIdGeneratorTest.cs`

**Interfaces:**
- Consumes: 없음 (순수 문자열).
- Produces: `static string AssetIdGenerator.Make(string category, int sequence)` → `"표지_01"`. `sequence`는 1-based, 2자리 zero-pad. 파일명 불가 문자(`\ / : * ? " < > |` 및 공백)는 `_`로 치환.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetIdGeneratorTest
    {
        [Fact]
        public void Make_Pads_Sequence_To_Two_Digits()
        {
            Assert.Equal("표지_01", AssetIdGenerator.Make("표지", 1));
            Assert.Equal("표지_12", AssetIdGenerator.Make("표지", 12));
        }

        [Fact]
        public void Make_Replaces_Filename_Unsafe_Chars_In_Category()
        {
            Assert.Equal("레이아웃_표지__01", AssetIdGenerator.Make("레이아웃(표지)", 1).Replace("(", "_").Replace(")", "_"));
        }

        [Fact]
        public void Make_Replaces_Slash_And_Space()
        {
            Assert.Equal("3단_강점_01", AssetIdGenerator.Make("3단 강점", 1));
            Assert.Equal("a_b_03", AssetIdGenerator.Make("a/b", 3));
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

빌드 후 테스트 (아래 "테스트 실행 절차" 참조). Expected: FAIL — `AssetIdGenerator` 미정의로 컴파일 에러.

- [ ] **Step 3: 최소 구현**

```csharp
using System.Linq;

namespace TeampptAddin
{
    /// <summary>섹션명 + 순번 → 파일명 안전한 에셋 ID. 예: ("표지", 1) → "표지_01".</summary>
    public static class AssetIdGenerator
    {
        private static readonly char[] Unsafe = { '\\', '/', ':', '*', '?', '"', '<', '>', '|', ' ', '(', ')' };

        public static string Make(string category, int sequence)
        {
            var safe = new string((category ?? "asset")
                .Select(c => Unsafe.Contains(c) ? '_' : c)
                .ToArray());
            return $"{safe}_{sequence:D2}";
        }
    }
}
```

참고: Step 1의 두 번째 테스트는 괄호도 `Make`가 직접 치환하므로 `레이아웃_표지__01`이 됨 (`(`→`_`, `)`→`_`). 테스트의 추가 `.Replace`는 no-op로 통과.

- [ ] **Step 4: 통과 확인**

테스트 실행. Expected: 3 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/AssetIdGenerator.cs src/TeampptAddin.Tests/AssetIdGeneratorTest.cs
git commit -m "feat(ingest): AssetIdGenerator (섹션명+순번 → 파일安全 ID)"
```

---

### Task 3: 인제스트 플래너

**Files:**
- Create: `src/TeampptAddin/Services/IngestPlanner.cs`
- Test: `src/TeampptAddin.Tests/IngestPlannerTest.cs`

**Interfaces:**
- Consumes: `AssetIdGenerator.Make`, `SectionInfo`, `AssetSplitItem`.
- Produces: `static List<AssetSplitItem> IngestPlanner.Plan(IReadOnlyList<SectionInfo> sections)`. 각 섹션의 슬라이드를 1장씩 펼쳐 항목 생성. 순번은 **섹션(카테고리)별로** 1부터. `SourceSlideIndex = section.FirstSlideIndex + offset`. 파일명 = `{AssetId}.pptx` / `.png`.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using System.Collections.Generic;
using Xunit;

namespace TeampptAddin.Tests
{
    public class IngestPlannerTest
    {
        [Fact]
        public void Plan_Expands_Each_Slide_To_One_Item_With_PerCategory_Sequence()
        {
            var sections = new List<SectionInfo>
            {
                new SectionInfo { Name = "표지", FirstSlideIndex = 1, SlideCount = 2 },
                new SectionInfo { Name = "목차", FirstSlideIndex = 3, SlideCount = 1 },
            };

            var items = IngestPlanner.Plan(sections);

            Assert.Equal(3, items.Count);

            Assert.Equal(1, items[0].SourceSlideIndex);
            Assert.Equal("표지", items[0].Category);
            Assert.Equal("표지_01", items[0].AssetId);
            Assert.Equal("표지_01.pptx", items[0].PptxFileName);
            Assert.Equal("표지_01.png", items[0].ThumbFileName);

            Assert.Equal(2, items[1].SourceSlideIndex);
            Assert.Equal("표지_02", items[1].AssetId);

            Assert.Equal(3, items[2].SourceSlideIndex);
            Assert.Equal("목차", items[2].Category);
            Assert.Equal("목차_01", items[2].AssetId);
        }

        [Fact]
        public void Plan_Empty_Sections_Returns_Empty()
        {
            Assert.Empty(IngestPlanner.Plan(new List<SectionInfo>()));
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

테스트 실행. Expected: FAIL — `IngestPlanner` 미정의.

- [ ] **Step 3: 최소 구현**

```csharp
using System.Collections.Generic;

namespace TeampptAddin
{
    /// <summary>섹션 목록 → 슬라이드 1장=에셋 1개 분할 계획. 순번은 카테고리별 1부터.</summary>
    public static class IngestPlanner
    {
        public static List<AssetSplitItem> Plan(IReadOnlyList<SectionInfo> sections)
        {
            var result = new List<AssetSplitItem>();
            if (sections == null) return result;

            foreach (var section in sections)
            {
                for (int offset = 0; offset < section.SlideCount; offset++)
                {
                    int sequence = offset + 1;
                    var id = AssetIdGenerator.Make(section.Name, sequence);
                    result.Add(new AssetSplitItem
                    {
                        SourceSlideIndex = section.FirstSlideIndex + offset,
                        Category = section.Name,
                        AssetId = id,
                        PptxFileName = id + ".pptx",
                        ThumbFileName = id + ".png",
                    });
                }
            }
            return result;
        }
    }
}
```

- [ ] **Step 4: 통과 확인**

테스트 실행. Expected: 2 PASS (전체 기존 테스트 포함 여전히 GREEN).

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/IngestPlanner.cs src/TeampptAddin.Tests/IngestPlannerTest.cs
git commit -m "feat(ingest): IngestPlanner (섹션→슬라이드 단위 분할 계획)"
```

---

### Task 4: 섹션 리더 (COM, 수동 검증)

**Files:**
- Create: `src/TeampptAddin/Core/SectionReader.cs`

**Interfaces:**
- Consumes: `Microsoft.Office.Interop.PowerPoint.Presentation`, `SectionInfo`.
- Produces: `static List<SectionInfo> SectionReader.Read(PowerPoint.Presentation pres)`. `pres.SectionProperties`를 1-based 순회하여 각 섹션의 `Name(i)`, `FirstSlide(i)`, `SlidesCount(i)`로 `SectionInfo` 구성. 섹션이 0개면 빈 리스트.

- [ ] **Step 1: 구현 작성**

```csharp
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 열린 Presentation의 PowerPoint 섹션을 읽어 SectionInfo 목록으로 변환.
    /// SectionProperties는 PowerPoint 2013+ 제공. 1-based 인덱스.
    /// </summary>
    public static class SectionReader
    {
        public static List<SectionInfo> Read(PowerPoint.Presentation pres)
        {
            var result = new List<SectionInfo>();
            var props = pres.SectionProperties;
            int count = props.Count;
            for (int i = 1; i <= count; i++)
            {
                result.Add(new SectionInfo
                {
                    Name = props.Name(i),
                    FirstSlideIndex = props.FirstSlide(i),
                    SlideCount = props.SlidesCount(i),
                });
            }
            return result;
        }
    }
}
```

- [ ] **Step 2: 빌드 확인**

Run: Task 1 Step 2의 MSBuild 명령. Expected: Build succeeded. (`SectionProperties.Name/FirstSlide/SlidesCount` API가 Interop에 존재함을 컴파일로 확인.)

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Core/SectionReader.cs
git commit -m "feat(ingest): SectionReader (Interop SectionProperties 읽기)"
```

> 수동 검증은 Task 7(오케스트레이터)에서 일괄 수행.

---

### Task 5: 슬라이드 스플리터 (COM, 수동 검증)

**Files:**
- Create: `src/TeampptAddin/Core/SlideSplitter.cs`

**Interfaces:**
- Consumes: `PowerPoint.Application`(=`Globals.Application`), `PowerPoint.Presentation`(소스).
- Produces: `static void SlideSplitter.SplitSlide(PowerPoint.Application app, PowerPoint.Presentation source, int slideIndex, string outputPptxPath)`. 빈 Presentation을 새로 만들고, 소스 슬라이드를 복사(`Slide.Copy()` → `Slides.Paste()`)하여 단일 슬라이드 pptx로 저장 후 닫기. 새 Presentation은 `Marshal.ReleaseComObject`.

- [ ] **Step 1: 구현 작성**

```csharp
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>소스 묶음 pptx에서 슬라이드 1장을 복사해 단일 슬라이드 pptx로 저장.</summary>
    public static class SlideSplitter
    {
        public static void SplitSlide(
            PowerPoint.Application app,
            PowerPoint.Presentation source,
            int slideIndex,
            string outputPptxPath)
        {
            var dir = Path.GetDirectoryName(outputPptxPath);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            PowerPoint.Presentation dest = null;
            try
            {
                // 화면에 띄우지 않고 새 프레젠테이션 생성
                dest = app.Presentations.Add(MsoTriState.msoFalse);

                source.Slides[slideIndex].Copy();
                dest.Slides.Paste();  // 인덱스 1에 붙음

                // 빈 기본 슬라이드가 있으면 제거 (Add 직후 기본 슬라이드 없음 — Paste만 존재)
                dest.SaveAs(
                    outputPptxPath,
                    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    MsoTriState.msoFalse);
                dest.Close();
            }
            finally
            {
                if (dest != null) Marshal.ReleaseComObject(dest);
            }
        }
    }
}
```

- [ ] **Step 2: 빌드 확인**

Run: Task 1 Step 2의 MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Core/SlideSplitter.cs
git commit -m "feat(ingest): SlideSplitter (슬라이드 복사→단일 pptx 저장)"
```

---

### Task 6: 슬라이드 PNG 렌더러 (COM, 수동 검증)

**Files:**
- Create: `src/TeampptAddin/Core/SlideImageRenderer.cs`

**Interfaces:**
- Consumes: `PowerPoint.Presentation`(소스), 슬라이드 인덱스, 긴 변 px.
- Produces: `static class SlideImageRenderer { const int LlmImageLongEdgePx = 768; static void Render(PowerPoint.Presentation source, int slideIndex, string outputPngPath, int longEdgePx = LlmImageLongEdgePx); }`. 슬라이드 비율을 `pres.PageSetup.SlideWidth/SlideHeight`로 구해 긴 변을 `longEdgePx`에 맞춘 정수 W/H를 계산하고 `slide.Export(path, "PNG", W, H)` 호출.

- [ ] **Step 1: 구현 작성**

```csharp
using System;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 슬라이드를 PNG로 렌더. 긴 변을 LlmImageLongEdgePx에 맞춰 다운스케일.
    /// LLM 화면읽기 전역 기본 = 768px (Gemini 1타일/258토큰, Claude ~442토큰).
    /// 텍스트는 pptx XML에서 읽으므로 OCR용 아님 → 768로 충분.
    /// </summary>
    public static class SlideImageRenderer
    {
        public const int LlmImageLongEdgePx = 768;

        public static void Render(
            PowerPoint.Presentation source,
            int slideIndex,
            string outputPngPath,
            int longEdgePx = LlmImageLongEdgePx)
        {
            var dir = Path.GetDirectoryName(outputPngPath);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            float slideW = source.PageSetup.SlideWidth;   // points
            float slideH = source.PageSetup.SlideHeight;

            int w, h;
            if (slideW >= slideH)
            {
                w = longEdgePx;
                h = (int)Math.Round(longEdgePx * (slideH / slideW));
            }
            else
            {
                h = longEdgePx;
                w = (int)Math.Round(longEdgePx * (slideW / slideH));
            }

            source.Slides[slideIndex].Export(outputPngPath, "PNG", w, h);
        }
    }
}
```

- [ ] **Step 2: 빌드 확인**

Run: Task 1 Step 2의 MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Core/SlideImageRenderer.cs
git commit -m "feat(ingest): SlideImageRenderer (긴변 768px PNG 렌더)"
```

---

### Task 7: 인제스트 오케스트레이터 + PowerPoint 수동 검증

**Files:**
- Create: `src/TeampptAddin/Services/IngestRunner.cs`

**Interfaces:**
- Consumes: `Globals.Application`, `SectionReader.Read`, `IngestPlanner.Plan`, `SlideSplitter.SplitSlide`, `SlideImageRenderer.Render`.
- Produces: `static List<AssetSplitItem> IngestRunner.Run(string bundlePptxPath, string outputDir)`. 묶음 pptx를 열고 → 섹션 읽기 → 계획 → 각 항목마다 split + render → 소스 닫기 → 계획 리스트 반환. pptx/png는 `outputDir` 아래에 저장.

- [ ] **Step 1: 구현 작성**

```csharp
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 로컬 인제스트 코어: 묶음 pptx → 섹션 읽기 → 슬라이드 1장=에셋 1개 split + 768px 썸네일.
    /// LLM/Supabase 없음. 분할된 pptx/png는 outputDir에 저장.
    /// </summary>
    public static class IngestRunner
    {
        public static List<AssetSplitItem> Run(string bundlePptxPath, string outputDir)
        {
            var app = Globals.Application;
            if (app == null)
            {
                Logger.Log("IngestRunner: app is null");
                return new List<AssetSplitItem>();
            }
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            PowerPoint.Presentation source = null;
            try
            {
                source = app.Presentations.Open(
                    bundlePptxPath,
                    MsoTriState.msoTrue,    // ReadOnly
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);  // WithWindow=False

                var sections = SectionReader.Read(source);
                Logger.Log($"Ingest: {sections.Count} sections in {Path.GetFileName(bundlePptxPath)}");

                var plan = IngestPlanner.Plan(sections);

                foreach (var item in plan)
                {
                    var pptxPath = Path.Combine(outputDir, item.PptxFileName);
                    var pngPath = Path.Combine(outputDir, item.ThumbFileName);
                    SlideSplitter.SplitSlide(app, source, item.SourceSlideIndex, pptxPath);
                    SlideImageRenderer.Render(source, item.SourceSlideIndex, pngPath);
                    Logger.Log($"Ingest: {item.AssetId} (slide {item.SourceSlideIndex})");
                }

                source.Close();
                return plan;
            }
            finally
            {
                if (source != null) Marshal.ReleaseComObject(source);
            }
        }
    }
}
```

- [ ] **Step 2: 본프로젝트 빌드 확인**

Run: Task 1 Step 2의 MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 전체 단위테스트 GREEN 확인**

"테스트 실행 절차" 수행. Expected: 기존 19개 + 신규(AssetIdGenerator 3, IngestPlanner 2) 모두 PASS.

- [ ] **Step 4: PowerPoint 수동 검증 (COM 경로)**

샘플 묶음 pptx(섹션 2개 이상, 섹션마다 슬라이드 2장 이상)를 준비하고, 임시 실행 지점(예: 디버그용 임시 메뉴 버튼 또는 즉시 창에서 `IngestRunner.Run(@"...bundle.pptx", @"%LOCALAPPDATA%\TeampptAddin\ingest-test")` 호출)으로 실행한다. 확인:
- `outputDir`에 `{섹션명}_NN.pptx`가 슬라이드 수만큼 생성됨.
- 각 pptx는 슬라이드 1장만 포함.
- `{섹션명}_NN.png`가 함께 생성되고 긴 변이 768px.
- `debug.log`에 섹션 수와 각 에셋 ID가 기록됨.

(임시 실행 지점은 검증 후 제거하거나 Phase B 후속 plan의 관리자 모드 버튼으로 대체. 이 plan에는 영구 UI 미포함.)

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/IngestRunner.cs
git commit -m "feat(ingest): IngestRunner 오케스트레이터 + 수동 검증 완료"
```

---

## 테스트 실행 절차 (모든 TDD Task 공통, 관리자 불필요)

```
# 1) 본프로젝트 (COM 등록 끔)
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.csproj" /t:Build /p:Configuration=Debug "/p:Platform=AnyCPU" /p:RegisterForComInterop=false /verbosity:minimal
# 2) 테스트프로젝트 (프로젝트참조 재빌드 안 함)
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" /t:Build /p:Configuration=Debug /p:BuildProjectReferences=false /verbosity:minimal
# 3) 테스트 실행
dotnet test "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" --no-build --no-restore
```

---

## 완료 정의

- 순수 로직 3개(IngestModels/AssetIdGenerator/IngestPlanner) 단위테스트 GREEN.
- COM 어댑터 3개(SectionReader/SlideSplitter/SlideImageRenderer) + IngestRunner 빌드 성공.
- PowerPoint 수동 검증: 섹션 분류된 묶음 pptx → 섹션명_NN.pptx(1슬라이드씩) + 768px PNG 생성 확인.
- 외부 의존성(Supabase/LLM/Gemini) 없음 — Supabase 준비와 병렬 진행 가능.

## 다음 plan (이 plan 밖)

- LLM 이해(Gemini/Claude 어댑터, responseSchema 확장으로 name/use_when/tags/slots/colors/fonts 생성).
- 임베딩 + Supabase 업로드 + 관리자 게이트(admin.json).
- 추천·삽입 읽기 경로.

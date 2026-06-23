# 화면 공유 진단 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** AI 탭에 「화면 공유 진단」 버튼을 추가해, 누르면 현재 슬라이드를 768px PNG로 캡처→Gemini 멀티모달 1회 호출로 개선점 진단 + 추천질문 3개를 채팅에 띄운다 (버튼식 1회성).

**Architecture:** 순수 로직(스키마/파서)은 HTTP에서 분리해 xUnit으로 테스트하고(기존 `UnderstandingSchema`/`UnderstandingParser` 패턴), COM 캡처와 WPF UI는 수동 검증한다. `IAiService`에 `DiagnoseSlideAsync`를 추가하고 3개 구현체(`GeminiAiService` 실구현, `VectorRecommendService` 위임, `MockAiService` 더미)에 채운다.

**Tech Stack:** C# .NET Framework 4.8, WPF(코드비하인드), Microsoft.Office.Interop.PowerPoint(COM), Newtonsoft.Json, xUnit 2.9.

## Global Constraints

- 빌드는 MSBuild 관리자 권한(`Start-Process -Verb RunAs`), 빌드 후 DLL 타임스탬프(1분 이내) + `build.log` 오류 0건 검증 (CLAUDE.md). MSBuild 경로: `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`.
- API 키를 문서·커밋에 평문으로 절대 넣지 않는다.
- 슬라이드 읽기는 이미지(PNG) 단독. 텍스트/도형 추출 없음.
- 호출은 버튼 1회성. 슬라이드 변경 자동 감지·폴링 없음.
- 진단은 개선점 중심 + 추천질문 정확히 3개.
- Gemini 모델: `gemini-2.5-flash`, 멀티모달 `inline_data`(`mime_type=image/png`, base64).
- 기존 UI 패턴 재사용: 타이핑 버블(`AddAiBubble`), 칩(`BuildChip`), 펄스/스캔 애니메이션.
- 스펙: `docs/superpowers/specs/2026-06-23-screen-share-diagnosis-design.md`.

---

### Task 1: 진단 모델 + 스키마 + 파서 (순수 로직, 테스트)

**Files:**
- Create: `src/TeampptAddin/Models/SlideDiagnosis.cs`
- Create: `src/TeampptAddin/Services/SlideDiagnosisSchema.cs`
- Create: `src/TeampptAddin/Services/SlideDiagnosisParser.cs`
- Test: `src/TeampptAddin.Tests/SlideDiagnosisParserTest.cs`
- Test: `src/TeampptAddin.Tests/SlideDiagnosisSchemaTest.cs`

**Interfaces:**
- Produces:
  - `class SlideDiagnosis { string Message; List<string> SuggestedQuestions; }`
  - `static SlideDiagnosisSchema.BuildResponseSchema() -> JObject`
  - `static SlideDiagnosisSchema.BuildSystemPrompt() -> string`
  - `static SlideDiagnosisParser.Parse(string llmJson) -> SlideDiagnosis`

- [ ] **Step 1: Write the failing tests**

`src/TeampptAddin.Tests/SlideDiagnosisParserTest.cs`:
```csharp
using System.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class SlideDiagnosisParserTest
    {
        [Fact]
        public void Parse_Reads_Message_And_Questions()
        {
            var json = @"{""message"":""대비가 약합니다."",""questions"":[""제목을 키울까?"",""색을 바꿀까?"",""여백을 줄일까?""]}";
            var d = SlideDiagnosisParser.Parse(json);
            Assert.Equal("대비가 약합니다.", d.Message);
            Assert.Equal(3, d.SuggestedQuestions.Count);
            Assert.Equal("제목을 키울까?", d.SuggestedQuestions[0]);
        }

        [Fact]
        public void Parse_Missing_Fields_Yields_Empty_Defaults()
        {
            var d = SlideDiagnosisParser.Parse("{}");
            Assert.Equal("", d.Message);
            Assert.Empty(d.SuggestedQuestions);
        }

        [Fact]
        public void Parse_Caps_Questions_At_Three()
        {
            var json = @"{""message"":""x"",""questions"":[""1"",""2"",""3"",""4"",""5""]}";
            var d = SlideDiagnosisParser.Parse(json);
            Assert.Equal(3, d.SuggestedQuestions.Count);
        }
    }
}
```

`src/TeampptAddin.Tests/SlideDiagnosisSchemaTest.cs`:
```csharp
using Xunit;

namespace TeampptAddin.Tests
{
    public class SlideDiagnosisSchemaTest
    {
        [Fact]
        public void ResponseSchema_Has_Message_And_Questions()
        {
            var schema = SlideDiagnosisSchema.BuildResponseSchema();
            var props = schema["properties"];
            Assert.NotNull(props["message"]);
            Assert.NotNull(props["questions"]);
            Assert.Equal("array", props["questions"]["type"].ToString());
        }

        [Fact]
        public void SystemPrompt_Mentions_Diagnosis_And_Three_Questions()
        {
            var p = SlideDiagnosisSchema.BuildSystemPrompt();
            Assert.Contains("개선", p);
            Assert.Contains("3", p);
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj --filter "FullyQualifiedName~SlideDiagnosis"`
Expected: FAIL — `SlideDiagnosis`/`SlideDiagnosisSchema`/`SlideDiagnosisParser` 미정의 컴파일 에러.

- [ ] **Step 3: Write the model**

`src/TeampptAddin/Models/SlideDiagnosis.cs`:
```csharp
using System.Collections.Generic;

namespace TeampptAddin
{
    public class SlideDiagnosis
    {
        public string Message { get; set; } = "";
        public List<string> SuggestedQuestions { get; set; } = new List<string>();
    }
}
```

- [ ] **Step 4: Write the schema**

`src/TeampptAddin/Services/SlideDiagnosisSchema.cs`:
```csharp
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public static class SlideDiagnosisSchema
    {
        public static JObject BuildResponseSchema()
        {
            return new JObject
            {
                ["type"] = "object",
                ["properties"] = new JObject
                {
                    ["message"] = new JObject { ["type"] = "string" },
                    ["questions"] = new JObject
                    {
                        ["type"] = "array",
                        ["items"] = new JObject { ["type"] = "string" }
                    }
                },
                ["required"] = new JArray { "message", "questions" }
            };
        }

        public static string BuildSystemPrompt()
        {
            return @"너는 발표 슬라이드를 진단하는 디자인 코치야. 슬라이드 이미지 1장을 보고 개선점 중심으로 진단해.

## message
- 강점이 있으면 한 문장으로 짧게 인정하고, 곧바로 구체적인 개선점 2~3가지를 제시해.
- 추상적 칭찬·일반론 금지. ""제목 대비가 약해 잘 안 읽힌다"", ""여백이 좌우로 치우쳤다"" 처럼 이 슬라이드에서 보이는 것만.
- 한국어, 친근하지만 군더더기 없이. 4~6문장 이내.

## questions
- 사용자가 이 진단을 보고 이어서 물어볼 만한 자연어 질문을 정확히 3개.
- 이 슬라이드 맥락에 붙는 실질 질문으로 (예: ""제목을 어떻게 키우면 좋을까?"", ""이 색 조합이 적절해?"").
- 각 질문은 한 문장.

모르면 지어내지 말고 보수적으로.";
        }
    }
}
```

- [ ] **Step 5: Write the parser**

`src/TeampptAddin/Services/SlideDiagnosisParser.cs`:
```csharp
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public static class SlideDiagnosisParser
    {
        public static SlideDiagnosis Parse(string llmJson)
        {
            var o = JObject.Parse(llmJson);
            var questions = (o["questions"] as JArray)?
                .Select(t => t.ToString())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Take(3)
                .ToList() ?? new List<string>();

            return new SlideDiagnosis
            {
                Message = o["message"]?.ToString() ?? "",
                SuggestedQuestions = questions
            };
        }
    }
}
```

- [ ] **Step 6: Add the new files to the main csproj if not auto-included**

Check `src/TeampptAddin/TeampptAddin.csproj` — if it uses explicit `<Compile Include=...>` items (old-style), add lines for the 3 new files under the appropriate `<ItemGroup>`. If it is SDK-style (globbing), skip.

Run: `grep -c "<Compile Include" src/TeampptAddin/TeampptAddin.csproj`
- If output > 0 → old-style; add:
```xml
    <Compile Include="Models\SlideDiagnosis.cs" />
    <Compile Include="Services\SlideDiagnosisSchema.cs" />
    <Compile Include="Services\SlideDiagnosisParser.cs" />
```
- If output is 0 → SDK-style; no edit needed.

- [ ] **Step 7: Run tests to verify they pass**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj --filter "FullyQualifiedName~SlideDiagnosis"`
Expected: PASS (5 tests).

- [ ] **Step 8: Commit**

```bash
git add src/TeampptAddin/Models/SlideDiagnosis.cs src/TeampptAddin/Services/SlideDiagnosisSchema.cs src/TeampptAddin/Services/SlideDiagnosisParser.cs src/TeampptAddin.Tests/SlideDiagnosisParserTest.cs src/TeampptAddin.Tests/SlideDiagnosisSchemaTest.cs src/TeampptAddin/TeampptAddin.csproj
git commit -m "feat(diagnosis): SlideDiagnosis 모델+스키마+파서 (테스트)"
```

---

### Task 2: IAiService.DiagnoseSlideAsync + 3개 구현체

**Files:**
- Modify: `src/TeampptAddin/Services/IAiService.cs` (인터페이스에 메서드 추가 + `MockAiService` 구현)
- Modify: `src/TeampptAddin/Services/GeminiAiService.cs` (실구현)
- Modify: `src/TeampptAddin/Services/VectorRecommendService.cs` (`_selector`로 위임)
- Test: `src/TeampptAddin.Tests/MockDiagnoseTest.cs`

**Interfaces:**
- Consumes (Task 1): `SlideDiagnosis`, `SlideDiagnosisSchema.BuildResponseSchema()`, `SlideDiagnosisSchema.BuildSystemPrompt()`, `SlideDiagnosisParser.Parse(string)`.
- Produces: `Task<SlideDiagnosis> IAiService.DiagnoseSlideAsync(string pngPath)`.

- [ ] **Step 1: Write the failing test**

`src/TeampptAddin.Tests/MockDiagnoseTest.cs`:
```csharp
using System.Threading.Tasks;
using Xunit;

namespace TeampptAddin.Tests
{
    public class MockDiagnoseTest
    {
        [Fact]
        public async Task Mock_Returns_Message_And_Three_Questions()
        {
            IAiService svc = new MockAiService();
            var d = await svc.DiagnoseSlideAsync("dummy.png");
            Assert.False(string.IsNullOrEmpty(d.Message));
            Assert.Equal(3, d.SuggestedQuestions.Count);
        }
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj --filter "FullyQualifiedName~MockDiagnose"`
Expected: FAIL — `IAiService`에 `DiagnoseSlideAsync` 없음 컴파일 에러.

- [ ] **Step 3: Add method to interface + MockAiService**

`src/TeampptAddin/Services/IAiService.cs` — 인터페이스에 메서드 추가:
```csharp
    public interface IAiService
    {
        Task<AiRecommendation> RecommendAsync(
            string userIntent,
            IEnumerable<HeaderAsset> assets,
            IEnumerable<StylePalette> palettes,
            IEnumerable<StyleFont> fonts);

        Task<SlideDiagnosis> DiagnoseSlideAsync(string pngPath);
    }
```
같은 파일 `MockAiService` 클래스 안에 구현 추가 (기존 `RecommendAsync` 아래):
```csharp
        public Task<SlideDiagnosis> DiagnoseSlideAsync(string pngPath)
        {
            return Task.FromResult(new SlideDiagnosis
            {
                Message = "제목 대비가 약해 한눈에 안 들어와요. 본문 여백이 좌우로 치우쳤고, 색이 3개를 넘어 산만합니다.",
                SuggestedQuestions = new System.Collections.Generic.List<string>
                {
                    "제목을 어떻게 키우면 좋을까?",
                    "색을 몇 개로 줄이면 좋을까?",
                    "여백을 어떻게 맞추지?"
                }
            });
        }
```

- [ ] **Step 4: Implement in GeminiAiService**

`src/TeampptAddin/Services/GeminiAiService.cs` — `using System.IO;`가 이미 있음(확인). `RecommendAsync` 아래에 추가:
```csharp
        public async Task<SlideDiagnosis> DiagnoseSlideAsync(string pngPath)
        {
            var base64 = Convert.ToBase64String(File.ReadAllBytes(pngPath));

            // 이미지+진단요청을 대화 히스토리에 적재 → 칩 후속 질문이 같은 슬라이드 맥락 유지
            _history.Add(new JObject
            {
                ["role"] = "user",
                ["parts"] = new JArray
                {
                    new JObject
                    {
                        ["inline_data"] = new JObject
                        {
                            ["mime_type"] = "image/png",
                            ["data"] = base64
                        }
                    },
                    new JObject { ["text"] = "이 슬라이드를 개선점 중심으로 진단해줘." }
                }
            });

            while (_history.Count > MaxHistoryTurns * 2)
                _history.RemoveAt(0);

            var requestBody = new JObject
            {
                ["contents"] = new JArray(_history.ToArray()),
                ["systemInstruction"] = new JObject
                {
                    ["parts"] = new JArray { new JObject { ["text"] = SlideDiagnosisSchema.BuildSystemPrompt() } }
                },
                ["generationConfig"] = new JObject
                {
                    ["temperature"] = 0.6,
                    ["responseMimeType"] = "application/json",
                    ["responseSchema"] = SlideDiagnosisSchema.BuildResponseSchema(),
                    ["thinkingConfig"] = new JObject { ["thinkingBudget"] = 0 }
                }
            };

            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={_apiKey}";
            var bodyString = requestBody.ToString(Formatting.None);

            const int maxAttempts = 3;
            HttpResponseMessage response = null;
            string body = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var content = new StringContent(bodyString, Encoding.UTF8, "application/json");
                Http.DefaultRequestHeaders.Authorization = null;
                response = await Http.PostAsync(url, content).ConfigureAwait(false);
                body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                Logger.Log($"[Diagnose] attempt {attempt}: HTTP {(int)response.StatusCode}");
                if (response.IsSuccessStatusCode) break;

                var status = (int)response.StatusCode;
                bool transient = status == 503 || status == 429 || status == 500;
                if (transient && attempt < maxAttempts)
                {
                    await Task.Delay(500 * (1 << (attempt - 1))).ConfigureAwait(false);
                    continue;
                }
                throw new HttpRequestException($"Gemini 진단 API 오류 ({status}): {body}");
            }

            var root = JObject.Parse(body);
            LogTokenUsage(root);
            var text = root["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();
            if (string.IsNullOrEmpty(text))
                throw new InvalidOperationException("Gemini 진단 응답에 텍스트가 없습니다.");

            _history.Add(new JObject
            {
                ["role"] = "model",
                ["parts"] = new JArray { new JObject { ["text"] = text } }
            });

            return SlideDiagnosisParser.Parse(text);
        }
```

- [ ] **Step 5: Implement in VectorRecommendService (위임)**

`src/TeampptAddin/Services/VectorRecommendService.cs` — `RecommendAsync` 아래에 추가. `_selector`는 진단·추천이 같은 `GeminiAiService` 인스턴스라 히스토리(이미지 맥락)가 공유된다:
```csharp
        public Task<SlideDiagnosis> DiagnoseSlideAsync(string pngPath)
            => _selector.DiagnoseSlideAsync(pngPath);
```

- [ ] **Step 6: Run test to verify it passes**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj --filter "FullyQualifiedName~MockDiagnose"`
Expected: PASS.

- [ ] **Step 7: Run full test suite (no regressions)**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj`
Expected: PASS (all existing + new).

- [ ] **Step 8: Commit**

```bash
git add src/TeampptAddin/Services/IAiService.cs src/TeampptAddin/Services/GeminiAiService.cs src/TeampptAddin/Services/VectorRecommendService.cs src/TeampptAddin.Tests/MockDiagnoseTest.cs
git commit -m "feat(diagnosis): IAiService.DiagnoseSlideAsync + Gemini 멀티모달 구현"
```

---

### Task 3: SlideCaptureService (현재 슬라이드 PNG 캡처, COM)

**Files:**
- Create: `src/TeampptAddin/Services/SlideCaptureService.cs`

**Interfaces:**
- Consumes: 기존 `SlideImageRenderer.Render(PowerPoint.Presentation, int, string)`, `Globals.Application`.
- Produces:
  - `class SlideCapture { string PngPath; int SlideNumber; }`
  - `static SlideCapture SlideCaptureService.CaptureCurrentSlide()` — 활성 슬라이드 없으면 `null`.

> COM 의존이라 자동 단위 테스트 불가 → 수동 검증(Task 5). 코드 자체는 완결.

- [ ] **Step 1: Write the service**

`src/TeampptAddin/Services/SlideCaptureService.cs`:
```csharp
using System;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public class SlideCapture
    {
        public string PngPath { get; set; }
        public int SlideNumber { get; set; }
    }

    /// <summary>
    /// 현재 편집 중인 슬라이드를 768px PNG로 캡처. 활성 슬라이드가 없으면 null.
    /// 버튼식 1회성 — 폴링/이벤트 구독 없음.
    /// </summary>
    public static class SlideCaptureService
    {
        public static SlideCapture CaptureCurrentSlide()
        {
            var app = Globals.Application;
            var win = app?.ActiveWindow;
            if (win == null) return null;

            PowerPoint.Slide slide;
            try { slide = win.View.Slide; }
            catch { return null; }   // 슬라이드쇼 등 Slide 접근 불가 뷰
            if (slide == null) return null;

            var pres = win.Presentation;
            int index = slide.SlideIndex;

            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TeampptAddin", "cache", "screen-share");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            var png = Path.Combine(dir, $"slide-{index}.png");
            SlideImageRenderer.Render(pres, index, png);

            return new SlideCapture { PngPath = png, SlideNumber = index };
        }
    }
}
```

- [ ] **Step 2: Add to csproj if old-style (same check as Task 1 Step 6)**

If `<Compile Include` count > 0, add:
```xml
    <Compile Include="Services\SlideCaptureService.cs" />
```

- [ ] **Step 3: Commit**

```bash
git add src/TeampptAddin/Services/SlideCaptureService.cs src/TeampptAddin/TeampptAddin.csproj
git commit -m "feat(diagnosis): SlideCaptureService 현재 슬라이드 PNG 캡처"
```

---

### Task 4: AssetPanel UI — 버튼 + 공유 인디케이터 + 진단 흐름

**Files:**
- Modify: `src/TeampptAddin/UI/Wpf/AssetPanel.cs`

**Interfaces:**
- Consumes: `SlideCaptureService.CaptureCurrentSlide()`, `IAiService.DiagnoseSlideAsync(string)`, `SlideDiagnosis`, 기존 `BuildInputBar()`, `AddAiBubble(string)`, `BuildChip(string)`, `RemoveLoadingBubble(FrameworkElement)`, `_chatStack`, `_chatScroll`, `_emptyState`, `_aiService`.

> WPF/COM UI → 수동 검증(Task 5). 아래 코드는 완결된 추가/수정.

- [ ] **Step 1: 필드 추가**

`AssetPanel` 클래스 필드 영역(예: `private Border _sendBtn;` 근처)에 추가:
```csharp
        private Border _shareBar;
```

- [ ] **Step 2: BuildAiTab의 row 1을 공유바+입력바 묶음으로 교체**

`BuildAiTab()`에서 현재:
```csharp
            var inputBar = BuildInputBar();
            Grid.SetRow(inputBar, 1);
            grid.Children.Add(inputBar);
```
를 아래로 교체:
```csharp
            var bottom = new StackPanel();
            bottom.Children.Add(BuildShareBar());
            bottom.Children.Add(BuildInputBar());
            Grid.SetRow(bottom, 1);
            grid.Children.Add(bottom);
```

- [ ] **Step 3: 공유바 빌드/상태 전환 메서드 추가**

`BuildInputBar()` 메서드 바로 아래에 추가:
```csharp
        private Border BuildShareBar()
        {
            _shareBar = new Border
            {
                Background = ThemeResources.BgChip,
                CornerRadius = new CornerRadius(10),
                Margin = new Thickness(10, 8, 10, 0),
                Padding = new Thickness(12, 8, 12, 8),
                Cursor = Cursors.Hand
            };
            SetShareBarIdle();
            return _shareBar;
        }

        private void SetShareBarIdle()
        {
            _shareBar.RenderTransform = null;
            _shareBar.Opacity = 1;
            _shareBar.Background = ThemeResources.BgChip;
            _shareBar.Cursor = Cursors.Hand;

            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            row.Children.Add(new TextBlock
            {
                Text = "🖥", FontSize = 13,
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = VerticalAlignment.Center
            });
            row.Children.Add(new TextBlock
            {
                Text = "화면 공유 진단", FontSize = 12, FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextAccent, FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center
            });
            _shareBar.Child = row;

            _shareBar.MouseLeftButtonUp -= ShareBarClick;
            _shareBar.MouseLeftButtonUp += ShareBarClick;
        }

        private async void ShareBarClick(object sender, MouseButtonEventArgs e)
        {
            await RunDiagnosisAsync();
        }

        private void EnterSharingState(int slideNumber)
        {
            _shareBar.MouseLeftButtonUp -= ShareBarClick;
            _shareBar.Cursor = Cursors.Arrow;
            _shareBar.Background = ThemeResources.BgCategoryActive;

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var dotScale = new ScaleTransform(1, 1);
            var dot = new Border
            {
                Width = 8, Height = 8, CornerRadius = new CornerRadius(4),
                Background = new SolidColorBrush(Color.FromRgb(79, 92, 245)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0),
                RenderTransform = dotScale,
                RenderTransformOrigin = new Point(0.5, 0.5)
            };
            try
            {
                var pulse = new DoubleAnimation
                {
                    From = 1.0, To = 1.5,
                    Duration = TimeSpan.FromSeconds(0.8),
                    AutoReverse = true,
                    RepeatBehavior = RepeatBehavior.Forever,
                    EasingFunction = new SineEase()
                };
                dotScale.BeginAnimation(ScaleTransform.ScaleXProperty, pulse);
                dotScale.BeginAnimation(ScaleTransform.ScaleYProperty, pulse);
            }
            catch { }
            Grid.SetColumn(dot, 0);
            grid.Children.Add(dot);

            var lbl = new TextBlock
            {
                Text = $"슬라이드 {slideNumber} 공유 중",
                FontSize = 12, FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextAccent, FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(lbl, 1);
            grid.Children.Add(lbl);

            var close = new TextBlock
            {
                Text = "✕", FontSize = 12,
                Foreground = ThemeResources.TextSub, FontFamily = ThemeResources.FontBase,
                Cursor = Cursors.Hand, VerticalAlignment = VerticalAlignment.Center
            };
            close.MouseLeftButtonUp += (s, e) => { e.Handled = true; ExitSharingState(); };
            Grid.SetColumn(close, 2);
            grid.Children.Add(close);

            _shareBar.Child = grid;

            // slide-in + fade ("팍" 등장)
            try
            {
                _shareBar.Opacity = 0;
                var tt = new TranslateTransform(0, 8);
                _shareBar.RenderTransform = tt;
                _shareBar.BeginAnimation(OpacityProperty,
                    new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(150)));
                tt.BeginAnimation(TranslateTransform.YProperty,
                    new DoubleAnimation(8, 0, TimeSpan.FromMilliseconds(150))
                    { EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut } });
            }
            catch { _shareBar.Opacity = 1; }
        }

        private void ExitSharingState()
        {
            SetShareBarIdle();
        }
```

- [ ] **Step 4: 진단 로딩 버블(스캔라인) 추가**

`AddAiLoadingBubble()` 메서드 아래에 추가 (에셋 슬롯머신과 별개로, "화면 읽는 중" 전용):
```csharp
        private FrameworkElement AddDiagnosisLoadingBubble()
        {
            var wrapper = new StackPanel { Margin = new Thickness(12, 4, 40, 4) };
            wrapper.Children.Add(new TextBlock
            {
                Text = "TEAMPPT AI", FontSize = 10, FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.Accent, Margin = new Thickness(4, 0, 0, 3)
            });

            var content = new StackPanel();
            content.Children.Add(new TextBlock
            {
                Text = "화면을 읽는 중···", FontSize = 11,
                Foreground = ThemeResources.TextDim, FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(0, 0, 0, 6)
            });

            var track = new Border
            {
                Height = 6, Width = 180, CornerRadius = new CornerRadius(3),
                Background = ThemeResources.BgSurface, ClipToBounds = true,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            var scan = new Border
            {
                Width = 50, CornerRadius = new CornerRadius(3),
                Background = ThemeResources.Accent,
                HorizontalAlignment = HorizontalAlignment.Left,
                RenderTransform = new TranslateTransform(-50, 0)
            };
            track.Child = scan;
            content.Children.Add(track);

            var border = new Border
            {
                Background = ThemeResources.BgAiResponse,
                CornerRadius = new CornerRadius(4, 13, 13, 13),
                Padding = new Thickness(14, 10, 14, 10),
                HorizontalAlignment = HorizontalAlignment.Left,
                MinWidth = 180,
                Child = content
            };
            wrapper.Children.Add(border);
            _chatStack.Children.Add(wrapper);

            // 스캔라인 왕복 (RemoveLoadingBubble이 border.Tag DispatcherTimer를 멈춤)
            var scanTt = (TranslateTransform)scan.RenderTransform;
            double x = -50; bool fwd = true;
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
            timer.Tick += (s, e) =>
            {
                x += fwd ? 8 : -8;
                if (x >= 180) fwd = false;
                if (x <= -50) fwd = true;
                scanTt.X = x;
            };
            timer.Start();
            border.Tag = timer;

            return wrapper;
        }
```

- [ ] **Step 5: 진단 실행 + 결과 렌더 메서드 추가**

위 메서드들 아래에 추가:
```csharp
        private async Task RunDiagnosisAsync()
        {
            if (_aiService == null)
            {
                AddAiBubble("AI 서비스를 초기화 중입니다.");
                return;
            }

            SlideCapture capture;
            try { capture = SlideCaptureService.CaptureCurrentSlide(); }
            catch (Exception ex)
            {
                AddAiBubble($"슬라이드를 읽지 못했어요: {ex.Message}");
                return;
            }
            if (capture == null)
            {
                AddAiBubble("공유할 슬라이드가 없어요. PowerPoint에서 슬라이드를 여세요.");
                return;
            }

            if (_emptyState != null && _emptyState.Visibility == Visibility.Visible)
                _emptyState.Visibility = Visibility.Collapsed;

            EnterSharingState(capture.SlideNumber);
            var loading = AddDiagnosisLoadingBubble();
            _chatScroll.ScrollToBottom();

            try
            {
                var diag = await _aiService.DiagnoseSlideAsync(capture.PngPath);
                RemoveLoadingBubble(loading);
                AddDiagnosisResult(diag);
            }
            catch (Exception ex)
            {
                RemoveLoadingBubble(loading);
                AddAiBubble($"진단 중 오류: {ex.Message}");
            }

            _chatScroll.ScrollToBottom();
        }

        private void AddDiagnosisResult(SlideDiagnosis diag)
        {
            AddAiBubble(diag?.Message ?? "진단 결과를 확인해보세요.");

            var qs = diag?.SuggestedQuestions ?? new List<string>();
            if (qs.Count == 0) return;

            _chatStack.Children.Add(new TextBlock
            {
                Text = "이어서 물어보기", FontSize = 11, FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextSub, Margin = new Thickness(12, 8, 0, 2)
            });

            var wrap = new WrapPanel { Margin = new Thickness(8, 0, 12, 4) };
            foreach (var q in qs.Take(3))
                wrap.Children.Add(BuildChip(q));
            _chatStack.Children.Add(wrap);
        }
```

- [ ] **Step 6: 빌드 (관리자 권한)**

Run:
```powershell
Start-Process -FilePath "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" -ArgumentList '"c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.sln" /t:Build /p:Configuration=Debug "/p:Platform=Any CPU" /verbosity:minimal /fileLogger /fileLoggerParameters:logfile=c:\Projects\teamppt-addin\build.log;verbosity=minimal' -Verb RunAs -Wait -WindowStyle Hidden
```
검증 (둘 다 통과해야 함):
```bash
stat -c '%y' c:/Projects/teamppt-addin/src/TeampptAddin/bin/Debug/TeampptAddin.dll   # 1분 이내
tail -5 c:/Projects/teamppt-addin/build.log                                          # 오류 0건
```

- [ ] **Step 7: Commit**

```bash
git add src/TeampptAddin/UI/Wpf/AssetPanel.cs
git commit -m "feat(diagnosis): AI탭 화면 공유 진단 버튼+인디케이터+진단 흐름"
```

---

### Task 5: 수동 검증 (PowerPoint 실제 동작)

**Files:** 없음 (검증 전용). 실패 시 해당 Task로 돌아가 수정.

- [ ] **Step 1: PowerPoint에서 애드인 로드, 슬라이드 1장 이상 연다**

- [ ] **Step 2: AI 탭 → 입력창 위 「🖥 화면 공유 진단」 버튼 확인 후 클릭**
  - 기대: 공유바가 "🔵 슬라이드 N 공유 중 ✕"로 **slide-in 애니메이션**과 함께 전환, 닷이 펄스.
  - 기대: 채팅에 "화면을 읽는 중···" 스캔라인 로딩 → 곧 진단 버블(타이핑) + "이어서 물어보기" 칩 3개.

- [ ] **Step 3: 토큰 로그로 이미지 1회 전송 확인**

Run: `tail -20 "$LOCALAPPDATA/TeampptAddin/logs"/*.log` (또는 Logger 출력 경로)
기대: `[Diagnose] attempt 1: HTTP 200`, `[Gemini] 토큰 사용:` input이 이미지 포함 규모(수백~)로 1회만.

- [ ] **Step 4: 칩 클릭 → 같은 슬라이드 맥락으로 답변되는지**
  - 기대: 칩 텍스트가 사용자 버블로 전송되고, AI가 그 슬라이드에 대한 답을 이어서 함(이미지가 히스토리에 있어 맥락 유지).

- [ ] **Step 5: ✕ 클릭 → 공유바가 「🖥 화면 공유 진단」 idle 버튼으로 복귀**

- [ ] **Step 6: 슬라이드 없는 상태(빈 프레젠테이션/슬라이드쇼) → 안내 메시지**
  - 기대: "공유할 슬라이드가 없어요…" 버블, 인디케이터 미전환.

- [ ] **Step 7: PROGRESS-BOARD.md + PITCH.md 갱신**
  - PROGRESS-BOARD: 잎 #3을 ✅로, 다음 잎으로 교체. A-2 첫 조각 완료 반영.
  - PITCH.md: §4의 "화면 공유 진단(설계 완료·구현 예정)"을 **§2 완성된 기능**으로 승격(비전문가 언어). (memory `feedback-pitch-and-board`)
  - Commit: `docs: 화면 공유 진단 완료 — BOARD/PITCH 갱신`

---

## Self-Review

**Spec coverage:**
- §4.1 SlideCaptureService → Task 3 ✅
- §4.2 SlideDiagnosis 모델 → Task 1 ✅
- §4.3 DiagnoseSlideAsync (3 구현체, 히스토리 적재, Mock) → Task 2 ✅
- §4.4 버튼+인디케이터(애니메이션)+스캔로딩+진단버블+칩 → Task 4 ✅
- §5 데이터 흐름 → Task 4 RunDiagnosisAsync ✅
- §6 에러 처리(슬라이드 없음/렌더 실패/Gemini 오류, 정리) → Task 3 null 반환 + Task 4 try/catch + RemoveLoadingBubble ✅
- §7 테스트(파서/스키마/Mock 단위 + 캡처/UI 수동) → Task 1·2 자동, Task 5 수동 ✅
- §8 빌드/검증 → Task 4 Step 6, Task 5 ✅

**Placeholder scan:** 모든 step에 실제 코드/명령 포함. "적절히 처리" 류 없음. ✅

**Type consistency:** `SlideDiagnosis{Message, SuggestedQuestions}`, `SlideCapture{PngPath, SlideNumber}`, `DiagnoseSlideAsync(string)→Task<SlideDiagnosis>`, `CaptureCurrentSlide()→SlideCapture?` — Task 1/2/3에서 정의, Task 4에서 동일 시그니처로 소비. ✅

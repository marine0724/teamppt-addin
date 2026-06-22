# A-1a · LLM 이해 어댑터 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 인제스트로 쪼갠 슬라이드(PNG + 섹션명)를 Gemini 멀티모달이 보고 이해하여, 구조화 에셋 레코드(`kind` 분류 · slots · 색역할 · 폰트 · use_when · content_fit · tags · 예시의도문장)와 임베딩용 `embed_text`를 생성한다. Supabase 없음(다음 plan).

**Architecture:** 순수 로직(응답 파서, embed_text 조립, responseSchema 빌더)은 xUnit으로 TDD. 멀티모달 HTTP 호출은 기존 [GeminiAiService](../../../src/TeampptAddin/Services/GeminiAiService.cs)의 HttpClient+재시도+responseSchema 패턴을 그대로 따르는 얇은 어댑터로 두고, API 키+실제 PNG가 필요하므로 PowerPoint/통합 수동 검증. 기존 Core(COM)/Services·Models(로직) 분리를 따른다.

**Tech Stack:** .NET Framework 4.8, Newtonsoft.Json 13.0.3, System.Net.Http.HttpClient, Gemini `gemini-2.5-flash`(멀티모달), xUnit 2.9.

## Global Constraints

- Core/Connect.cs/Globals.cs 직접 수정 금지 (신규 파일만 추가).
- 의존성 추가 금지 — Newtonsoft.Json + Office Interop만. HTTP는 `System.Net.Http.HttpClient` 직접.
- API 키는 평문 커밋 금지. 키는 `Assets/api-keys.json`의 `gemini` 필드(기존), gitignore 유지.
- PNG 렌더 해상도 = `SlideImageRenderer.LlmImageLongEdgePx`(768) 산출물 재사용. 추가 렌더 안 함.
- `kind` 고정 enum = **`layout` | `component`** (그래프·표·다이어그램은 component). LLM은 이 둘 중 하나만 배정.
- `category`는 LLM이 만들지 않음 — 섹션명(`AssetSplitItem.Category`)을 그대로 주입.
- 단위테스트: 본프로젝트 `/p:RegisterForComInterop=false` 빌드 → 테스트프로젝트 `/p:BuildProjectReferences=false` 빌드 → `dotnet test --no-build --no-restore` (절차는 맨 아래).
- MSBuild = `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`.

---

## File Structure

| 파일 | 책임 | 테스트 |
|---|---|---|
| `Models/AssetUnderstanding.cs` (신규) | LLM 이해 결과 = `HeaderAsset Asset` + `List<string> ExampleIntents`(임베딩용). 순수 데이터 | 불필요 |
| `Services/UnderstandingSchema.cs` (신규) | Gemini `responseSchema`(JObject) + system 프롬프트. `kind` enum 강제 | `UnderstandingSchemaTest.cs` |
| `Services/UnderstandingParser.cs` (신규) | LLM JSON 텍스트 + 주입 category/file → `AssetUnderstanding` (순수) | `UnderstandingParserTest.cs` |
| `Services/EmbedTextBuilder.cs` (신규) | `AssetUnderstanding` → 임베딩 원문 문자열 (순수) | `EmbedTextBuilderTest.cs` |
| `Services/AssetUnderstandingService.cs` (신규) | PNG 경로 + category → Gemini 멀티모달 호출 → `AssetUnderstanding` (HTTP 어댑터) | 수동/통합 |

순수 4개(Models/Schema/Parser/EmbedText)가 이번 plan의 단위테스트 가능 산출물. HTTP 1개(Service)는 빌드 + 수동 검증.

---

### Task 1: 이해 결과 모델

**Files:**
- Create: `src/TeampptAddin/Models/AssetUnderstanding.cs`

**Interfaces:**
- Produces: `AssetUnderstanding { HeaderAsset Asset; List<string> ExampleIntents }` — 이후 모든 Task가 사용. `Asset`은 기존 [HeaderAsset](../../../src/TeampptAddin/Models/HeaderAsset.cs)(File/Name/Kind/Category/Scope/ContentFit/UseWhen/Tags/Colors/Fonts/Slots) 재사용. `ExampleIntents`는 임베딩 전용(예: "투자 유치 IR 표지")이라 HeaderAsset에 안 넣고 분리.

- [ ] **Step 1: 모델 작성**

```csharp
using System.Collections.Generic;

namespace TeampptAddin
{
    /// <summary>
    /// LLM이 슬라이드 1장(PNG+섹션명)을 이해한 결과.
    /// Asset = Supabase metadata/검색 골격으로 저장될 구조화 레코드.
    /// ExampleIntents = 임베딩 원문(embed_text)에만 쓰는 예시 의도 문장들(저장 레코드엔 미포함).
    /// </summary>
    public class AssetUnderstanding
    {
        public HeaderAsset Asset { get; set; }
        public List<string> ExampleIntents { get; set; } = new List<string>();
    }
}
```

- [ ] **Step 2: 빌드 확인**

Run:
```
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.csproj" /t:Build /p:Configuration=Debug "/p:Platform=AnyCPU" /p:RegisterForComInterop=false /verbosity:minimal
```
Expected: Build succeeded.

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Models/AssetUnderstanding.cs
git commit -m "feat(ingest): AssetUnderstanding 이해 결과 모델"
```

---

### Task 2: 이해 응답 스키마 + 시스템 프롬프트

**Files:**
- Create: `src/TeampptAddin/Services/UnderstandingSchema.cs`
- Test: `src/TeampptAddin.Tests/UnderstandingSchemaTest.cs`

**Interfaces:**
- Consumes: 없음.
- Produces:
  - `static JObject UnderstandingSchema.BuildResponseSchema()` — Gemini `generationConfig.responseSchema`. 최상위 `required` = `["name","kind","use_when","content_fit","tags","example_intents","slots","colors","fonts"]`. `kind`는 `enum: ["layout","component"]`.
  - `static string UnderstandingSchema.BuildSystemPrompt(string category)` — 섹션명을 힌트로 주는 시스템 프롬프트.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using Newtonsoft.Json.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class UnderstandingSchemaTest
    {
        [Fact]
        public void Schema_Requires_Core_Fields()
        {
            var schema = UnderstandingSchema.BuildResponseSchema();
            var required = (JArray)schema["required"];
            Assert.Contains("name", required.ToObject<string[]>());
            Assert.Contains("kind", required.ToObject<string[]>());
            Assert.Contains("slots", required.ToObject<string[]>());
            Assert.Contains("colors", required.ToObject<string[]>());
            Assert.Contains("example_intents", required.ToObject<string[]>());
        }

        [Fact]
        public void Schema_Kind_Is_Constrained_Enum()
        {
            var schema = UnderstandingSchema.BuildResponseSchema();
            var kindEnum = (JArray)schema["properties"]["kind"]["enum"];
            Assert.Equal(2, kindEnum.Count);
            Assert.Contains("layout", kindEnum.ToObject<string[]>());
            Assert.Contains("component", kindEnum.ToObject<string[]>());
        }

        [Fact]
        public void SystemPrompt_Includes_Category_Hint()
        {
            var prompt = UnderstandingSchema.BuildSystemPrompt("표지");
            Assert.Contains("표지", prompt);
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

테스트 실행(맨 아래 절차). Expected: FAIL — `UnderstandingSchema` 미정의 컴파일 에러.

- [ ] **Step 3: 최소 구현**

```csharp
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>
    /// 인제스트 이해(understanding) 단계용 Gemini responseSchema + 시스템 프롬프트.
    /// 모델이 구조를 벗어날 수 없게 강제. kind는 layout/component 둘 중 하나로 제약.
    /// </summary>
    public static class UnderstandingSchema
    {
        public static JObject BuildResponseSchema()
        {
            JObject StrArray() => new JObject
            {
                ["type"] = "array",
                ["items"] = new JObject { ["type"] = "string" }
            };

            return new JObject
            {
                ["type"] = "object",
                ["properties"] = new JObject
                {
                    ["name"] = new JObject { ["type"] = "string" },
                    ["kind"] = new JObject
                    {
                        ["type"] = "string",
                        ["enum"] = new JArray { "layout", "component" }
                    },
                    ["use_when"] = new JObject { ["type"] = "string" },
                    ["content_fit"] = StrArray(),
                    ["tags"] = StrArray(),
                    ["example_intents"] = StrArray(),
                    ["slots"] = new JObject
                    {
                        ["type"] = "array",
                        ["items"] = new JObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JObject
                            {
                                ["name"] = new JObject { ["type"] = "string" },
                                ["type"] = new JObject
                                {
                                    ["type"] = "string",
                                    ["enum"] = new JArray { "text", "image", "chart", "table" }
                                },
                                ["perSlide"] = new JObject { ["type"] = "boolean" }
                            },
                            ["required"] = new JArray { "name", "type", "perSlide" }
                        }
                    },
                    ["colors"] = new JObject
                    {
                        ["type"] = "array",
                        ["items"] = new JObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JObject
                            {
                                ["role"] = new JObject { ["type"] = "string" },
                                ["value"] = new JObject { ["type"] = "string" },
                                ["locked"] = new JObject { ["type"] = "boolean" }
                            },
                            ["required"] = new JArray { "role", "value", "locked" }
                        }
                    },
                    ["fonts"] = new JObject
                    {
                        ["type"] = "array",
                        ["items"] = new JObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JObject
                            {
                                ["role"] = new JObject { ["type"] = "string" },
                                ["family"] = new JObject { ["type"] = "string" },
                                ["weight"] = new JObject { ["type"] = "string" }
                            },
                            ["required"] = new JArray { "role", "family" }
                        }
                    }
                },
                ["required"] = new JArray
                {
                    "name", "kind", "use_when", "content_fit",
                    "tags", "example_intents", "slots", "colors", "fonts"
                }
            };
        }

        public static string BuildSystemPrompt(string category)
        {
            return $@"너는 PPT 에셋을 분석하는 디자인 인제스트 엔진이야. 슬라이드 이미지 1장과 섹션명 힌트를 보고, 이 에셋을 미래에 검색·재사용하기 위한 구조화 메타데이터를 생성해.

## 섹션명(카테고리 힌트)
""{category}""

## 판단 규칙
- kind: 슬라이드 페이지 전체 틀이면 ""layout"", 틀 위에 얹는 부품(그래프/표/다이어그램 등)이면 ""component"".
- name: 이 에셋을 한눈에 구분할 짧은 한국어 이름 (예: ""우측정렬 연도강조 표지"").
- use_when: 어떤 상황에서 쓰면 좋은지 한 문장.
- content_fit: 들어가기 좋은 콘텐츠 종류 2~4개.
- tags: 검색용 키워드 3~8개.
- example_intents: 사용자가 이 에셋을 찾을 때 칠 법한 자연어 의도 문장 3~5개 (예: ""투자 유치 IR 표지"", ""회사 소개 첫 장""). 검색 임베딩 품질의 핵심이니 다양하게.
- slots: 사용자 글/이미지가 들어갈 이름 붙은 빈자리. 위치·크기·폰트 단서로 title/subtitle/body/image1 등 추론. type은 text|image|chart|table. 슬라이드마다 내용이 바뀌면 perSlide=true.
- colors: 핵심 색을 역할(main/sub1/sub2/text/accent 등)+hex로. 로고/브랜드 고정색은 locked=true.
- fonts: 보이는 폰트를 역할(heading/body 등)+family로.
- 모르면 지어내지 말고 빈 배열/보수적 값으로.";
        }
    }
}
```

- [ ] **Step 4: 통과 확인**

테스트 실행. Expected: 3 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/UnderstandingSchema.cs src/TeampptAddin.Tests/UnderstandingSchemaTest.cs
git commit -m "feat(ingest): UnderstandingSchema (responseSchema + kind enum 강제 + 프롬프트)"
```

---

### Task 3: 이해 응답 파서

**Files:**
- Create: `src/TeampptAddin/Services/UnderstandingParser.cs`
- Test: `src/TeampptAddin.Tests/UnderstandingParserTest.cs`

**Interfaces:**
- Consumes: `AssetUnderstanding`, `HeaderAsset`, `AssetColor`, `AssetFont`, `AssetSlot`.
- Produces: `static AssetUnderstanding UnderstandingParser.Parse(string llmJson, string category, string file)`. LLM JSON을 `HeaderAsset`로 매핑하되 **category/file은 인자로 주입**(LLM이 만들지 않음), `SchemaVersion=2`, `Scope="slide"`. `example_intents`는 `AssetUnderstanding.ExampleIntents`로. 누락 필드는 빈 리스트/빈 문자열로 안전 처리.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using Xunit;

namespace TeampptAddin.Tests
{
    public class UnderstandingParserTest
    {
        private const string Sample = @"{
          ""name"": ""우측정렬 연도강조 표지"",
          ""kind"": ""layout"",
          ""use_when"": ""연도를 강조하는 표지가 필요할 때"",
          ""content_fit"": [""표지"", ""연도 강조""],
          ""tags"": [""표지"", ""연도"", ""미니멀""],
          ""example_intents"": [""투자 유치 IR 표지"", ""회사 소개 첫 장""],
          ""slots"": [{""name"":""title"",""type"":""text"",""perSlide"":true}],
          ""colors"": [{""role"":""main"",""value"":""#1A2B4C"",""locked"":false}],
          ""fonts"": [{""role"":""heading"",""family"":""Pretendard""}]
        }";

        [Fact]
        public void Parse_Injects_Category_And_File_Not_From_Llm()
        {
            var u = UnderstandingParser.Parse(Sample, "표지", "표지_01.pptx");
            Assert.Equal("표지", u.Asset.Category);
            Assert.Equal("표지_01.pptx", u.Asset.File);
            Assert.Equal(2, u.Asset.SchemaVersion);
            Assert.Equal("slide", u.Asset.Scope);
        }

        [Fact]
        public void Parse_Maps_Core_Llm_Fields()
        {
            var u = UnderstandingParser.Parse(Sample, "표지", "표지_01.pptx");
            Assert.Equal("우측정렬 연도강조 표지", u.Asset.Name);
            Assert.Equal("layout", u.Asset.Kind);
            Assert.Equal("연도를 강조하는 표지가 필요할 때", u.Asset.UseWhen);
            Assert.Single(u.Asset.Slots);
            Assert.Equal("title", u.Asset.Slots[0].Name);
            Assert.Equal("#1A2B4C", u.Asset.Colors[0].Value);
            Assert.Equal("Pretendard", u.Asset.Fonts[0].Family);
        }

        [Fact]
        public void Parse_Extracts_Example_Intents_Separately()
        {
            var u = UnderstandingParser.Parse(Sample, "표지", "표지_01.pptx");
            Assert.Equal(2, u.ExampleIntents.Count);
            Assert.Contains("투자 유치 IR 표지", u.ExampleIntents);
        }

        [Fact]
        public void Parse_Missing_Arrays_Default_To_Empty()
        {
            var u = UnderstandingParser.Parse(@"{""name"":""x"",""kind"":""component""}", "표", "x.pptx");
            Assert.Empty(u.Asset.Tags);
            Assert.Empty(u.Asset.Slots);
            Assert.Empty(u.ExampleIntents);
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

테스트 실행. Expected: FAIL — `UnderstandingParser` 미정의.

- [ ] **Step 3: 최소 구현**

```csharp
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>LLM 이해 JSON → AssetUnderstanding. category/file은 코드가 주입(LLM 미생성).</summary>
    public static class UnderstandingParser
    {
        public static AssetUnderstanding Parse(string llmJson, string category, string file)
        {
            var o = JObject.Parse(llmJson);

            var asset = new HeaderAsset
            {
                SchemaVersion = 2,
                File = file,
                Category = category,
                Scope = "slide",
                Name = o["name"]?.ToString() ?? "",
                Kind = o["kind"]?.ToString() ?? "component",
                UseWhen = o["use_when"]?.ToString() ?? "",
                ContentFit = StrList(o["content_fit"]),
                Tags = StrList(o["tags"]),
                Colors = (o["colors"] as JArray)?.Select(c => new AssetColor
                {
                    Role = c["role"]?.ToString(),
                    Value = c["value"]?.ToString(),
                    Locked = c["locked"]?.Value<bool>() ?? false
                }).ToList() ?? new List<AssetColor>(),
                Fonts = (o["fonts"] as JArray)?.Select(f => new AssetFont
                {
                    Role = f["role"]?.ToString(),
                    Family = f["family"]?.ToString(),
                    Weight = f["weight"]?.ToString()
                }).ToList() ?? new List<AssetFont>(),
                Slots = (o["slots"] as JArray)?.Select(s => new AssetSlot
                {
                    Name = s["name"]?.ToString(),
                    Type = s["type"]?.ToString(),
                    PerSlide = s["perSlide"]?.Value<bool>() ?? false
                }).ToList() ?? new List<AssetSlot>()
            };

            return new AssetUnderstanding
            {
                Asset = asset,
                ExampleIntents = StrList(o["example_intents"])
            };
        }

        private static List<string> StrList(JToken token)
        {
            return (token as JArray)?.Select(t => t.ToString()).ToList() ?? new List<string>();
        }
    }
}
```

- [ ] **Step 4: 통과 확인**

테스트 실행. Expected: 4 PASS (기존 테스트 포함 전부 GREEN).

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/UnderstandingParser.cs src/TeampptAddin.Tests/UnderstandingParserTest.cs
git commit -m "feat(ingest): UnderstandingParser (LLM JSON→HeaderAsset, category/file 주입)"
```

---

### Task 4: 임베딩 원문 조립기

**Files:**
- Create: `src/TeampptAddin/Services/EmbedTextBuilder.cs`
- Test: `src/TeampptAddin.Tests/EmbedTextBuilderTest.cs`

**Interfaces:**
- Consumes: `AssetUnderstanding`.
- Produces: `static string EmbedTextBuilder.Build(AssetUnderstanding u)`. 검색 임베딩용 의미 문서 = `name` + `category` + `use_when` + `content_fit` + `tags` + `example_intents`를 줄바꿈으로 연결. 색 hex·폰트 family 등 "삽입용 구조 데이터"는 **제외**(검색과 무관, 토큰 절감 — embed_text↔metadata 분리 원칙).

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using System.Collections.Generic;
using Xunit;

namespace TeampptAddin.Tests
{
    public class EmbedTextBuilderTest
    {
        private static AssetUnderstanding Sample()
        {
            return new AssetUnderstanding
            {
                Asset = new HeaderAsset
                {
                    Name = "연도강조 표지",
                    Category = "표지",
                    UseWhen = "연도를 강조할 때",
                    ContentFit = new List<string> { "표지", "연도 강조" },
                    Tags = new List<string> { "표지", "연도" },
                    Colors = new List<AssetColor> { new AssetColor { Role = "main", Value = "#1A2B4C" } }
                },
                ExampleIntents = new List<string> { "투자 유치 IR 표지" }
            };
        }

        [Fact]
        public void Build_Includes_Search_Relevant_Fields()
        {
            var text = EmbedTextBuilder.Build(Sample());
            Assert.Contains("연도강조 표지", text);
            Assert.Contains("표지", text);
            Assert.Contains("연도를 강조할 때", text);
            Assert.Contains("투자 유치 IR 표지", text);
        }

        [Fact]
        public void Build_Excludes_Insertion_Only_Data()
        {
            var text = EmbedTextBuilder.Build(Sample());
            Assert.DoesNotContain("#1A2B4C", text);  // 색 hex는 삽입용, 검색 임베딩 제외
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

테스트 실행. Expected: FAIL — `EmbedTextBuilder` 미정의.

- [ ] **Step 3: 최소 구현**

```csharp
using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    /// <summary>
    /// 검색 임베딩용 "의미 문서" 조립. name+category+use_when+content_fit+tags+example_intents만.
    /// 색 hex·폰트 family 같은 삽입용 구조 데이터는 제외(embed_text↔metadata 분리).
    /// </summary>
    public static class EmbedTextBuilder
    {
        public static string Build(AssetUnderstanding u)
        {
            var a = u.Asset;
            var lines = new List<string> { a.Name, a.Category, a.UseWhen };
            lines.AddRange(a.ContentFit ?? new List<string>());
            lines.AddRange(a.Tags ?? new List<string>());
            lines.AddRange(u.ExampleIntents ?? new List<string>());
            return string.Join("\n", lines.Where(s => !string.IsNullOrWhiteSpace(s)));
        }
    }
}
```

- [ ] **Step 4: 통과 확인**

테스트 실행. Expected: 2 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/EmbedTextBuilder.cs src/TeampptAddin.Tests/EmbedTextBuilderTest.cs
git commit -m "feat(ingest): EmbedTextBuilder (검색 의미문서 조립, 삽입데이터 제외)"
```

---

### Task 5: 멀티모달 이해 서비스 (HTTP 어댑터, 수동 검증)

**Files:**
- Create: `src/TeampptAddin/Services/AssetUnderstandingService.cs`

**Interfaces:**
- Consumes: `UnderstandingSchema`, `UnderstandingParser`, `AssetUnderstanding`. 키 로딩은 `Globals.AssetsDir`의 `api-keys.json` `gemini` 필드(기존 [GeminiAiService.FromAssetsDir](../../../src/TeampptAddin/Services/GeminiAiService.cs#L25) 패턴 재사용).
- Produces: `class AssetUnderstandingService { AssetUnderstandingService(string apiKey); static AssetUnderstandingService FromAssetsDir(string assetsDir); Task<AssetUnderstanding> UnderstandAsync(string pngPath, string category, string file); }`. PNG를 base64 inline_data로 + 섹션명 텍스트로 멀티모달 호출 → `UnderstandingParser.Parse`. 503/429/500 재시도(기존 백오프 패턴).

- [ ] **Step 1: 구현 작성**

```csharp
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>
    /// 인제스트 이해 단계: 슬라이드 PNG(긴변 768px) + 섹션명 → Gemini 멀티모달 → 구조화 AssetUnderstanding.
    /// 기존 GeminiAiService의 HttpClient+재시도+responseSchema 패턴을 따른다.
    /// </summary>
    public class AssetUnderstandingService
    {
        private static readonly HttpClient Http = new HttpClient();
        private readonly string _apiKey;

        public AssetUnderstandingService(string apiKey) { _apiKey = apiKey; }

        public static AssetUnderstandingService FromAssetsDir(string assetsDir)
        {
            var path = Path.Combine(assetsDir, "api-keys.json");
            var obj = JObject.Parse(File.ReadAllText(path, Encoding.UTF8));
            var key = obj["gemini"]?.ToString()
                ?? throw new InvalidOperationException("api-keys.json에 'gemini' 키가 없습니다.");
            return new AssetUnderstandingService(key);
        }

        public async Task<AssetUnderstanding> UnderstandAsync(string pngPath, string category, string file)
        {
            var base64 = Convert.ToBase64String(File.ReadAllBytes(pngPath));

            var requestBody = new JObject
            {
                ["contents"] = new JArray
                {
                    new JObject
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
                            new JObject { ["text"] = $"섹션명: {category}. 이 슬라이드를 분석해 구조화 메타데이터를 생성해." }
                        }
                    }
                },
                ["systemInstruction"] = new JObject
                {
                    ["parts"] = new JArray { new JObject { ["text"] = UnderstandingSchema.BuildSystemPrompt(category) } }
                },
                ["generationConfig"] = new JObject
                {
                    ["temperature"] = 0.4,
                    ["responseMimeType"] = "application/json",
                    ["responseSchema"] = UnderstandingSchema.BuildResponseSchema()
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
                Logger.Log($"[Understand] {Path.GetFileName(pngPath)} attempt {attempt}: HTTP {(int)response.StatusCode}");

                if (response.IsSuccessStatusCode) break;

                var status = (int)response.StatusCode;
                bool transient = status == 503 || status == 429 || status == 500;
                if (transient && attempt < maxAttempts)
                {
                    await Task.Delay(500 * (1 << (attempt - 1))).ConfigureAwait(false);
                    continue;
                }
                throw new HttpRequestException($"Gemini 이해 API 오류 ({status}): {body}");
            }

            var root = JObject.Parse(body);
            var text = root["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();
            if (string.IsNullOrEmpty(text))
                throw new InvalidOperationException("Gemini 이해 응답에 텍스트가 없습니다.");

            return UnderstandingParser.Parse(text, category, file);
        }
    }
}
```

- [ ] **Step 2: 본프로젝트 빌드 확인**

Run: Task 1 Step 2의 MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 전체 단위테스트 GREEN 확인**

"테스트 실행 절차" 수행. Expected: 기존 + 신규(UnderstandingSchema 3, UnderstandingParser 4, EmbedTextBuilder 2) 모두 PASS.

- [ ] **Step 4: 수동/통합 검증 (실제 PNG + 키 필요)**

전제: 유효한 Gemini 키가 `Assets/api-keys.json`의 `gemini`에 있음(현재 키 작동 확인됨 — **키 형식이 `AIza`가 아니어도 손대지 말 것**. 신형식 키가 PPT 내 대화에서 정상 작동 검증됨). 로컬 인제스트 코어로 생성된 PNG 하나(예: `%LOCALAPPDATA%\TeampptAddin\ingest-test\표지_01.png`)를 대상으로, 임시 실행 지점(즉시 창/디버그 버튼)에서:
```csharp
var svc = AssetUnderstandingService.FromAssetsDir(Globals.AssetsDir);
var u = svc.UnderstandAsync(@"%LOCALAPPDATA%\TeampptAddin\ingest-test\표지_01.png", "표지", "표지_01.pptx").GetAwaiter().GetResult();
Logger.Log($"[Understand] name={u.Asset.Name}, kind={u.Asset.Kind}, slots={u.Asset.Slots.Count}, intents={u.ExampleIntents.Count}");
Logger.Log($"[Understand] embed_text=\n{EmbedTextBuilder.Build(u)}");
```
확인(`debug.log`):
- `kind`가 `layout` 또는 `component` 중 하나.
- `name`/`use_when`이 슬라이드 내용과 그럴듯하게 맞음.
- `slots`에 title 등 자리 추론됨.
- `example_intents`가 3~5개.
- `embed_text`에 색 hex가 없고 의미 문장만.

(임시 실행 지점은 검증 후 제거 — 영구 인제스트 버튼은 A-1b/c plan의 관리자 모드에서.)

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/AssetUnderstandingService.cs
git commit -m "feat(ingest): AssetUnderstandingService 멀티모달 이해 + 수동 검증"
```

---

## 테스트 실행 절차 (모든 TDD Task 공통, 관리자 불필요)

```
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.csproj" /t:Build /p:Configuration=Debug "/p:Platform=AnyCPU" /p:RegisterForComInterop=false /verbosity:minimal
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" /t:Build /p:Configuration=Debug /p:BuildProjectReferences=false /verbosity:minimal
dotnet test "c:\Projects\teamppt-addin\src\TeampptAddin.Tests\TeampptAddin.Tests.csproj" --no-build --no-restore
```

---

## 완료 정의

- 순수 4개(AssetUnderstanding/UnderstandingSchema/UnderstandingParser/EmbedTextBuilder) 단위테스트 GREEN.
- `AssetUnderstandingService` 빌드 성공 + 실제 PNG 1장 수동 검증(kind 분류·슬롯·예시의도·embed_text 확인).
- Supabase/임베딩 호출 없음 — 다음 plan(A-1b/c)에서 이 `AssetUnderstanding` + `EmbedTextBuilder` 출력을 임베딩→업로드로 연결.

## 다음 plan (이 plan 밖)

- **A-1b/c:** Supabase 인프라(테이블/RPC/Storage/RLS) + `text-embedding-004` 임베딩 호출 + 업로드 + `admin.json` 게이트(`IAccessPolicy` seam).
- **A-1d:** anon 벡터검색 읽기경로 + AI탭 텍스트 질의 추천 연결.

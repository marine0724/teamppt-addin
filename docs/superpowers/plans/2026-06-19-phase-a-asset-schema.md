# Phase A — 에셋 데이터 스키마 v2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** PowerPoint/COM/LLM 없이 순수 C#/JSON 로직만으로, 역할 기반 색·폰트 + 슬롯 + scope + kind(layout/component)를 담는 에셋 스키마 v2와 그 위의 마이그레이터·카탈로그 빌더·컨셉 리졸버를 단위테스트와 함께 구축한다.

**Architecture:** 기존 `Models/`·`Services/`를 확장한다. 저장 모델(`HeaderAsset`, 풍부·버전 있음)과 런타임 카탈로그(`CatalogEntry`, 컴팩트)를 분리하고, 둘 사이를 `CatalogBuilder`가 잇는다. 구버전 JSON은 `AssetSchemaMigrator`가 v1→v2로 흡수한다. 컨셉 적용(역할 치환)은 순수 함수 `ConceptResolver`로 분리해 단위테스트한다.

**Tech Stack:** .NET Framework 4.8, C# (LangVersion latest), Newtonsoft.Json 13.0.3, xUnit (신규 테스트 프로젝트).

## Global Constraints

- 대상 프레임워크: `net48` (메인 `src/TeampptAddin/TeampptAddin.csproj` 와 동일).
- JSON 직렬화는 **Newtonsoft.Json 13.0.3** 만 사용. `[JsonProperty]` 어트리뷰트로 snake_case 매핑(기존 패턴 그대로).
- 네임스페이스는 모두 `TeampptAddin` (기존 모델·서비스 동일).
- 메인 csproj는 old-style(non-SDK) + 명시적 `<Compile Include>`. **새 .cs 파일은 반드시 `<Compile Include>` 항목을 추가**해야 빌드에 포함된다.
- 색/폰트는 절대값이 아니라 **역할(role)** 단위로 저장한다: 색 `{role, value, locked}`, 폰트 `{role, family, fallback, weight, source}`.
- 에셋은 **`kind: layout | component`** 를 갖는다. `layout`=슬라이드 전체 틀(표지/목차/간지 등), `component`=레이아웃 위에 붙이는 부품(그래프/다이어그램/표). 기존 header_N은 `component` 기본값. v1 마이그레이션 시 kind 누락 → `"component"`.
- 슬롯 식별은 **텍스트 박스 + shape 이름 규약** (`slot.title`, `slot.image1` 등). Placeholder가 아님. 슬롯 type: `text | image | chart | table`.
- `HeaderAsset.Colors`(기존 `AssetColors` 객체형)는 **소비처가 전혀 없음**(확인 완료: UI의 `.Colors`는 전부 `StylePalette.PaletteColors`). 따라서 타입을 `List<AssetColor>`로 바꿔도 다른 코드가 깨지지 않는다. `AssetColors` 클래스는 삭제한다.
- `Core/` 폴더, `Connect.cs`, `Globals.cs` 는 **수정 금지**.
- COM 영역(`Slides`, `Application`, Vision, LLM)은 Phase A 범위 **밖**. 이 Phase의 모든 코드는 PowerPoint 없이 동작·테스트되어야 한다.

### Test Runner Notes (Task 1에서 확정)

- 1순위: `dotnet test`. 메인 프로젝트가 COM 등록(`RegisterForComInterop=true`)을 트리거하므로 **항상 `-p:RegisterForComInterop=false`** 를 붙여 관리자 권한 없이 빌드한다. PowerPoint가 열려 있으면 DLL 잠금으로 빌드 실패 → **빌드/테스트 전 PowerPoint 종료**.
- `dotnet` SDK가 legacy COM 메인 프로젝트 빌드에서 실패하면 폴백: VS의 `MSBuild.exe`로 솔루션 빌드(`/p:RegisterForComInterop=false`) 후 `vstest.console.exe`로 테스트 DLL 실행. Task 1에서 어느 경로가 동작하는지 확정하고, 이후 태스크는 그 경로의 명령을 사용한다.

---

## File Structure

| 파일 | 책임 | 변경 |
|------|------|------|
| `src/TeampptAddin/Models/AssetSchema.cs` | 신규 값 타입: `AssetColor`, `AssetFont`, `AssetSlot` | Create |
| `src/TeampptAddin/Models/HeaderAsset.cs` | 저장 모델 확장(SchemaVersion/Scope/Provenance/Fonts/Slots, Colors→List) | Modify |
| `src/TeampptAddin/Models/DesignConcept.cs` | 컨셉 모델(팔레트 역할→hex, 폰트 역할→family, styleTags) | Create |
| `src/TeampptAddin/Models/CatalogEntry.cs` | 런타임 컴팩트 카탈로그 항목 | Create |
| `src/TeampptAddin/Services/AssetSchemaMigrator.cs` | v1→v2 정규화(JObject→JObject) | Create |
| `src/TeampptAddin/Services/CatalogBuilder.cs` | 저장 레코드 → 컴팩트 카탈로그 | Create |
| `src/TeampptAddin/Services/ConceptResolver.cs` | 역할 치환 순수 함수 | Create |
| `src/TeampptAddin/Services/AssetLoader.cs` | 로드 시 마이그레이터 경유 | Modify |
| `src/TeampptAddin/Assets/assets.json` | 7개 항목 v2로 재작성 | Modify |
| `src/TeampptAddin/TeampptAddin.csproj` | 신규 .cs `<Compile Include>` 추가 | Modify |
| `src/TeampptAddin.Tests/TeampptAddin.Tests.csproj` | 신규 xUnit 테스트 프로젝트(net48) | Create |
| `src/TeampptAddin.Tests/*.cs` | 단위 테스트 | Create |

---

## Task 1: 테스트 하니스 구축 + 러너 확정

**Files:**
- Create: `src/TeampptAddin.Tests/TeampptAddin.Tests.csproj`
- Create: `src/TeampptAddin.Tests/SmokeTest.cs`
- Modify: `src/TeampptAddin/TeampptAddin.sln` (솔루션에 테스트 프로젝트 추가; `dotnet test`로 csproj 직접 지정 시 생략 가능)

**Interfaces:**
- Produces: 동작하는 테스트 러너 명령 1개(이후 모든 태스크가 재사용).

- [ ] **Step 1: SDK-style 테스트 프로젝트 생성**

`src/TeampptAddin.Tests/TeampptAddin.Tests.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <IsPackable>false</IsPackable>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.10.0" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\TeampptAddin\TeampptAddin.csproj" />
  </ItemGroup>
</Project>
```

- [ ] **Step 2: 트리비얼 테스트 작성 (실패 유도 불필요, 하니스 검증용)**

`src/TeampptAddin.Tests/SmokeTest.cs`:
```csharp
using Xunit;

namespace TeampptAddin.Tests
{
    public class SmokeTest
    {
        [Fact]
        public void Harness_Works()
        {
            Assert.Equal(4, 2 + 2);
        }
    }
}
```

- [ ] **Step 3: 러너 확정 — PowerPoint 종료 후 실행**

Run (1순위):
```
dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false
```
Expected: `Passed! - Failed: 0, Passed: 1`.
실패 시(legacy 프로젝트 빌드 오류 등) → Test Runner Notes의 MSBuild+vstest 폴백으로 전환하고, 동작한 명령을 이 파일 상단에 메모. **이후 모든 "Run:"은 확정된 러너를 사용**(아래 태스크들은 `dotnet test` 표기).

- [ ] **Step 4: 커밋**
```
git add src/TeampptAddin.Tests/
git commit -m "test: add xUnit harness for asset schema (Phase A)"
```

---

## Task 2: 스키마 값 타입 + HeaderAsset 확장

**Files:**
- Create: `src/TeampptAddin/Models/AssetSchema.cs`
- Modify: `src/TeampptAddin/Models/HeaderAsset.cs`
- Modify: `src/TeampptAddin/TeampptAddin.csproj` (`<Compile Include="Models\AssetSchema.cs" />`)
- Test: `src/TeampptAddin.Tests/HeaderAssetDeserializeTest.cs`

**Interfaces:**
- Produces:
  - `class AssetColor { string Role; string Value; bool Locked; }`
  - `class AssetFont { string Role; string Family; string Fallback; string Weight; string Source; }`
  - `class AssetSlot { string Name; string Type; bool PerSlide; }`
  - `HeaderAsset` 추가 필드: `int SchemaVersion`, `string Kind` (기본 `"component"`), `string Scope`, `string Provenance`, `List<AssetColor> Colors`, `List<AssetFont> Fonts`, `List<AssetSlot> Slots`.

- [ ] **Step 1: 실패 테스트 작성**

`src/TeampptAddin.Tests/HeaderAssetDeserializeTest.cs`:
```csharp
using Newtonsoft.Json;
using Xunit;

namespace TeampptAddin.Tests
{
    public class HeaderAssetDeserializeTest
    {
        [Fact]
        public void Deserializes_V2_Asset_With_Roles_Fonts_Slots()
        {
            const string json = @"{
              ""schemaVersion"": 2,
              ""file"": ""header_3.pptx"",
              ""name"": ""장점 나열"",
              ""kind"": ""layout"",
              ""category"": ""헤더"",
              ""scope"": ""deck"",
              ""provenance"": ""IR 덱"",
              ""colors"": [ { ""role"": ""main"", ""value"": ""#2563EB"", ""locked"": false } ],
              ""fonts"": [ { ""role"": ""heading"", ""family"": ""Pretendard"", ""fallback"": ""맑은 고딕"", ""weight"": ""Bold"", ""source"": ""bundled"" } ],
              ""slots"": [ { ""name"": ""title"", ""type"": ""text"", ""perSlide"": true } ]
            }";

            var a = JsonConvert.DeserializeObject<HeaderAsset>(json);

            Assert.Equal(2, a.SchemaVersion);
            Assert.Equal("layout", a.Kind);
            Assert.Equal("deck", a.Scope);
            Assert.Equal("IR 덱", a.Provenance);
            Assert.Single(a.Colors);
            Assert.Equal("main", a.Colors[0].Role);
            Assert.False(a.Colors[0].Locked);
            Assert.Equal("Pretendard", a.Fonts[0].Family);
            Assert.True(a.Slots[0].PerSlide);
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter HeaderAssetDeserializeTest`
Expected: 컴파일 실패(`AssetColor` 등 미정의 / `Colors` 타입 불일치).

- [ ] **Step 3: 값 타입 생성**

`src/TeampptAddin/Models/AssetSchema.cs`:
```csharp
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public class AssetColor
    {
        [JsonProperty("role")] public string Role { get; set; }
        [JsonProperty("value")] public string Value { get; set; }
        [JsonProperty("locked")] public bool Locked { get; set; }
    }

    public class AssetFont
    {
        [JsonProperty("role")] public string Role { get; set; }
        [JsonProperty("family")] public string Family { get; set; }
        [JsonProperty("fallback")] public string Fallback { get; set; }
        [JsonProperty("weight")] public string Weight { get; set; }
        [JsonProperty("source")] public string Source { get; set; }
    }

    public class AssetSlot
    {
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("type")] public string Type { get; set; }
        [JsonProperty("perSlide")] public bool PerSlide { get; set; }
        [JsonExtensionData] public Dictionary<string, JToken> Extra { get; set; }
    }
}
```

- [ ] **Step 4: HeaderAsset 확장 (Colors 타입 교체 + 신규 필드, `AssetColors` 삭제)**

`src/TeampptAddin/Models/HeaderAsset.cs` 를 다음으로 교체:
```csharp
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public class HeaderAsset
    {
        [JsonProperty("schemaVersion")] public int SchemaVersion { get; set; } = 1;
        [JsonProperty("file")] public string File { get; set; }
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("kind")] public string Kind { get; set; } = "component";
        [JsonProperty("category")] public string Category { get; set; }
        [JsonProperty("scope")] public string Scope { get; set; } = "slide";
        [JsonProperty("content_fit")] public List<string> ContentFit { get; set; }
        [JsonProperty("use_when")] public string UseWhen { get; set; }
        [JsonProperty("provenance")] public string Provenance { get; set; }
        [JsonProperty("grid_columns")] public int GridColumns { get; set; } = 1;
        [JsonProperty("tags")] public List<string> Tags { get; set; }
        [JsonProperty("colors")] public List<AssetColor> Colors { get; set; }
        [JsonProperty("fonts")] public List<AssetFont> Fonts { get; set; }
        [JsonProperty("slots")] public List<AssetSlot> Slots { get; set; }
        [JsonExtensionData] public Dictionary<string, JToken> Extra { get; set; }
    }
}
```

- [ ] **Step 5: csproj에 Compile 추가**

`src/TeampptAddin/TeampptAddin.csproj` 의 `<Compile Include="Models\HeaderAsset.cs" />` 바로 다음 줄에 추가:
```xml
    <Compile Include="Models\AssetSchema.cs" />
```

- [ ] **Step 6: 통과 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter HeaderAssetDeserializeTest`
Expected: PASS.

- [ ] **Step 7: 커밋**
```
git add src/TeampptAddin/Models/ src/TeampptAddin/TeampptAddin.csproj src/TeampptAddin.Tests/
git commit -m "feat: role-based color/font + slot/scope schema on HeaderAsset (Phase A)"
```

---

## Task 3: AssetSchemaMigrator (v1 → v2 정규화)

**Files:**
- Create: `src/TeampptAddin/Services/AssetSchemaMigrator.cs`
- Modify: `src/TeampptAddin/TeampptAddin.csproj`
- Test: `src/TeampptAddin.Tests/AssetSchemaMigratorTest.cs`

**Interfaces:**
- Produces: `static class AssetSchemaMigrator { static JObject Migrate(JObject raw); }`
  - v1(색이 객체 `{main,sub1,sub2,text}`)을 v2(색이 `[{role,value,locked}]` 배열)로 변환.
  - `schemaVersion` 누락 시 1로 간주 후 2로 승격. `scope` 누락 시 `"slide"`.
  - 이미 v2(색이 배열)면 그대로 통과.

- [ ] **Step 1: 실패 테스트 작성**

`src/TeampptAddin.Tests/AssetSchemaMigratorTest.cs`:
```csharp
using Newtonsoft.Json.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetSchemaMigratorTest
    {
        [Fact]
        public void Migrates_V1_Object_Colors_To_Role_Array()
        {
            var v1 = JObject.Parse(@"{
              ""file"": ""header_1.pptx"", ""name"": ""x"", ""category"": ""헤더"",
              ""colors"": { ""main"": ""#2563EB"", ""sub1"": ""#3B82F6"", ""sub2"": ""#93C5FD"", ""text"": ""#1E293B"" }
            }");

            var v2 = AssetSchemaMigrator.Migrate(v1);

            Assert.Equal(2, (int)v2["schemaVersion"]);
            Assert.Equal("slide", (string)v2["scope"]);
            var colors = (JArray)v2["colors"];
            Assert.Equal(4, colors.Count);
            Assert.Equal("main", (string)colors[0]["role"]);
            Assert.Equal("#2563EB", (string)colors[0]["value"]);
            Assert.False((bool)colors[0]["locked"]);
        }

        [Fact]
        public void Passes_Through_V2_Array_Colors()
        {
            var v2in = JObject.Parse(@"{
              ""schemaVersion"": 2, ""scope"": ""deck"",
              ""colors"": [ { ""role"": ""main"", ""value"": ""#000000"", ""locked"": true } ]
            }");

            var v2 = AssetSchemaMigrator.Migrate(v2in);

            Assert.Equal(2, (int)v2["schemaVersion"]);
            Assert.Equal("deck", (string)v2["scope"]);
            var colors = (JArray)v2["colors"];
            Assert.Single(colors);
            Assert.True((bool)colors[0]["locked"]);
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

Run: `dotnet test ... --filter AssetSchemaMigratorTest`
Expected: 컴파일 실패(`AssetSchemaMigrator` 미정의).

- [ ] **Step 3: 구현**

`src/TeampptAddin/Services/AssetSchemaMigrator.cs`:
```csharp
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public static class AssetSchemaMigrator
    {
        // v1의 객체형 colors를 역할 배열로 변환할 때 사용하는 순서/역할 키.
        private static readonly string[] V1ColorRoles = { "main", "sub1", "sub2", "text" };

        public static JObject Migrate(JObject raw)
        {
            var obj = (JObject)raw.DeepClone();

            if (obj["kind"] == null)
                obj["kind"] = "component";

            if (obj["scope"] == null)
                obj["scope"] = "slide";

            var colors = obj["colors"];
            if (colors is JObject colorObj)
            {
                var arr = new JArray();
                foreach (var role in V1ColorRoles)
                {
                    var val = colorObj[role];
                    if (val == null) continue;
                    arr.Add(new JObject
                    {
                        ["role"] = role,
                        ["value"] = val,
                        ["locked"] = false
                    });
                }
                obj["colors"] = arr;
            }

            obj["schemaVersion"] = 2;
            return obj;
        }
    }
}
```

- [ ] **Step 4: csproj에 Compile 추가**

`<Compile Include="Services\AssetLoader.cs" />` 위/주변에 추가:
```xml
    <Compile Include="Services\AssetSchemaMigrator.cs" />
```

- [ ] **Step 5: 통과 확인**

Run: `dotnet test ... --filter AssetSchemaMigratorTest`
Expected: PASS (2 tests).

- [ ] **Step 6: 커밋**
```
git add src/TeampptAddin/Services/AssetSchemaMigrator.cs src/TeampptAddin/TeampptAddin.csproj src/TeampptAddin.Tests/
git commit -m "feat: AssetSchemaMigrator v1->v2 (object colors -> role array)"
```

---

## Task 4: AssetLoader가 마이그레이터 경유

**Files:**
- Modify: `src/TeampptAddin/Services/AssetLoader.cs`
- Test: `src/TeampptAddin.Tests/AssetLoaderTest.cs`

**Interfaces:**
- Consumes: `AssetSchemaMigrator.Migrate(JObject)`, `HeaderAsset`.
- Produces: `AssetLoader.Load(string assetsDir)` 가 v1/v2 혼재 JSON을 모두 v2 `HeaderAsset`로 로드. 폴더 스캔 폴백·파일 존재 검증은 유지.

- [ ] **Step 1: 실패 테스트 작성**

`src/TeampptAddin.Tests/AssetLoaderTest.cs`:
```csharp
using System.IO;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetLoaderTest
    {
        [Fact]
        public void Load_Migrates_V1_File_To_Role_Colors()
        {
            var dir = Path.Combine(Path.GetTempPath(), "teamppt_test_" + System.Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                // 파일 존재 검증을 통과시키기 위해 더미 pptx 생성
                File.WriteAllText(Path.Combine(dir, "header_1.pptx"), "dummy");
                File.WriteAllText(Path.Combine(dir, "assets.json"), @"[
                  { ""file"": ""header_1.pptx"", ""name"": ""타이틀"", ""category"": ""헤더"",
                    ""colors"": { ""main"": ""#2563EB"", ""text"": ""#FFFFFF"" } }
                ]");

                var assets = AssetLoader.Load(dir);

                Assert.Single(assets);
                Assert.Equal(2, assets[0].SchemaVersion);
                Assert.NotNull(assets[0].Colors);
                Assert.Contains(assets[0].Colors, c => c.Role == "main" && c.Value == "#2563EB");
            }
            finally { Directory.Delete(dir, true); }
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

Run: `dotnet test ... --filter AssetLoaderTest`
Expected: FAIL (`SchemaVersion`이 0 또는 colors 바인딩 실패 — 기존 로더는 객체형 colors를 `List<AssetColor>`로 못 읽음).

- [ ] **Step 3: 구현 — LoadFromJson을 JArray→마이그레이터→바인딩으로 교체**

`src/TeampptAddin/Services/AssetLoader.cs` 의 `LoadFromJson` 를 교체하고 `using Newtonsoft.Json.Linq;` 추가:
```csharp
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public static class AssetLoader
    {
        public static List<HeaderAsset> Load(string assetsDir)
        {
            var jsonPath = Path.Combine(assetsDir, "assets.json");

            var assets = File.Exists(jsonPath)
                ? LoadFromJson(jsonPath)
                : ScanFolder(assetsDir);

            return assets
                .Where(a => File.Exists(Path.Combine(assetsDir, a.File)))
                .ToList();
        }

        private static List<HeaderAsset> LoadFromJson(string jsonPath)
        {
            var json = File.ReadAllText(jsonPath);
            var arr = JArray.Parse(json);
            var result = new List<HeaderAsset>();
            foreach (var token in arr)
            {
                if (!(token is JObject obj)) continue;
                var migrated = AssetSchemaMigrator.Migrate(obj);
                result.Add(migrated.ToObject<HeaderAsset>());
            }
            return result;
        }

        private static List<HeaderAsset> ScanFolder(string assetsDir)
        {
            if (!Directory.Exists(assetsDir))
                return new List<HeaderAsset>();

            return Directory.GetFiles(assetsDir, "header_*.pptx")
                .OrderBy(f => f)
                .Select(f => new HeaderAsset
                {
                    SchemaVersion = 2,
                    File = Path.GetFileName(f),
                    Name = Path.GetFileNameWithoutExtension(f),
                    Kind = "component",
                    Category = "헤더",
                    Scope = "slide",
                    ContentFit = new List<string>(),
                    UseWhen = "",
                    GridColumns = 1
                })
                .ToList();
        }
    }
}
```

- [ ] **Step 4: 통과 확인**

Run: `dotnet test ... --filter AssetLoaderTest`
Expected: PASS.

- [ ] **Step 5: 커밋**
```
git add src/TeampptAddin/Services/AssetLoader.cs src/TeampptAddin.Tests/
git commit -m "feat: AssetLoader migrates assets through AssetSchemaMigrator"
```

---

## Task 5: DesignConcept + CatalogEntry + CatalogBuilder

**Files:**
- Create: `src/TeampptAddin/Models/DesignConcept.cs`
- Create: `src/TeampptAddin/Models/CatalogEntry.cs`
- Create: `src/TeampptAddin/Services/CatalogBuilder.cs`
- Modify: `src/TeampptAddin/TeampptAddin.csproj`
- Test: `src/TeampptAddin.Tests/CatalogBuilderTest.cs`

**Interfaces:**
- Produces:
  - `class DesignConcept { string Id; string Name; List<string> StyleTags; Dictionary<string,string> Colors; Dictionary<string,string> Fonts; }` (Colors: 역할→hex, Fonts: 역할→family)
  - `class CatalogEntry { string File; string Name; string Kind; string Category; string Scope; List<string> Tags; string UseWhen; List<string> SlotNames; List<string> ColorRoles; List<string> FontRoles; }`
  - `static class CatalogBuilder { static List<CatalogEntry> Build(IEnumerable<HeaderAsset> assets); }`
  - 빌더는 **hex 값·폰트 family 등 무거운 값을 제외**하고 역할 이름·슬롯 이름만 투영(런타임 토큰 절감).

- [ ] **Step 1: 실패 테스트 작성**

`src/TeampptAddin.Tests/CatalogBuilderTest.cs`:
```csharp
using System.Collections.Generic;
using Xunit;

namespace TeampptAddin.Tests
{
    public class CatalogBuilderTest
    {
        [Fact]
        public void Build_Projects_Compact_Entry_Without_Hex_Values()
        {
            var asset = new HeaderAsset
            {
                File = "header_3.pptx", Name = "장점 나열", Category = "헤더", Scope = "deck",
                Tags = new List<string> { "장점", "나열" },
                UseWhen = "장점 3개",
                Colors = new List<AssetColor> { new AssetColor { Role = "main", Value = "#2563EB" } },
                Fonts = new List<AssetFont> { new AssetFont { Role = "heading", Family = "Pretendard" } },
                Slots = new List<AssetSlot> { new AssetSlot { Name = "title" }, new AssetSlot { Name = "body" } }
            };

            var entries = CatalogBuilder.Build(new[] { asset });

            Assert.Single(entries);
            var e = entries[0];
            Assert.Equal("header_3.pptx", e.File);
            Assert.Equal("deck", e.Scope);
            Assert.Equal(new[] { "title", "body" }, e.SlotNames);
            Assert.Equal(new[] { "main" }, e.ColorRoles);
            Assert.Equal(new[] { "heading" }, e.FontRoles);
            // 무거운 값은 투영되지 않음 — CatalogEntry에 hex/family 필드 자체가 없어야 한다.
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

Run: `dotnet test ... --filter CatalogBuilderTest`
Expected: 컴파일 실패(`DesignConcept`/`CatalogEntry`/`CatalogBuilder` 미정의).

- [ ] **Step 3: 모델 생성**

`src/TeampptAddin/Models/DesignConcept.cs`:
```csharp
using System.Collections.Generic;
using Newtonsoft.Json;

namespace TeampptAddin
{
    public class DesignConcept
    {
        [JsonProperty("id")] public string Id { get; set; }
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("styleTags")] public List<string> StyleTags { get; set; }
        // 역할 → hex
        [JsonProperty("colors")] public Dictionary<string, string> Colors { get; set; }
        // 역할 → family
        [JsonProperty("fonts")] public Dictionary<string, string> Fonts { get; set; }
    }
}
```

`src/TeampptAddin/Models/CatalogEntry.cs`:
```csharp
using System.Collections.Generic;

namespace TeampptAddin
{
    public class CatalogEntry
    {
        public string File { get; set; }
        public string Name { get; set; }
        public string Kind { get; set; }
        public string Category { get; set; }
        public string Scope { get; set; }
        public List<string> Tags { get; set; }
        public string UseWhen { get; set; }
        public List<string> SlotNames { get; set; }
        public List<string> ColorRoles { get; set; }
        public List<string> FontRoles { get; set; }
    }
}
```

- [ ] **Step 4: 빌더 구현**

`src/TeampptAddin/Services/CatalogBuilder.cs`:
```csharp
using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    public static class CatalogBuilder
    {
        public static List<CatalogEntry> Build(IEnumerable<HeaderAsset> assets)
        {
            return assets.Select(a => new CatalogEntry
            {
                File = a.File,
                Name = a.Name,
                Kind = a.Kind,
                Category = a.Category,
                Scope = a.Scope,
                Tags = a.Tags ?? new List<string>(),
                UseWhen = a.UseWhen,
                SlotNames = (a.Slots ?? new List<AssetSlot>()).Select(s => s.Name).ToList(),
                ColorRoles = (a.Colors ?? new List<AssetColor>()).Select(c => c.Role).ToList(),
                FontRoles = (a.Fonts ?? new List<AssetFont>()).Select(f => f.Role).ToList()
            }).ToList();
        }
    }
}
```

- [ ] **Step 5: csproj에 Compile 3건 추가**
```xml
    <Compile Include="Models\DesignConcept.cs" />
    <Compile Include="Models\CatalogEntry.cs" />
    <Compile Include="Services\CatalogBuilder.cs" />
```

- [ ] **Step 6: 통과 확인**

Run: `dotnet test ... --filter CatalogBuilderTest`
Expected: PASS.

- [ ] **Step 7: 커밋**
```
git add src/TeampptAddin/Models/ src/TeampptAddin/Services/CatalogBuilder.cs src/TeampptAddin/TeampptAddin.csproj src/TeampptAddin.Tests/
git commit -m "feat: DesignConcept + compact CatalogEntry + CatalogBuilder"
```

---

## Task 6: ConceptResolver (역할 치환 순수 함수)

**Files:**
- Create: `src/TeampptAddin/Services/ConceptResolver.cs`
- Modify: `src/TeampptAddin/TeampptAddin.csproj`
- Test: `src/TeampptAddin.Tests/ConceptResolverTest.cs`

**Interfaces:**
- Consumes: `HeaderAsset`, `AssetColor`, `AssetFont`, `DesignConcept`.
- Produces:
  - `class ResolvedColor { string Role; string Value; }`
  - `class ResolvedFont { string Role; string Family; }`
  - `static class ConceptResolver`:
    - `static List<ResolvedColor> ResolveColors(HeaderAsset asset, DesignConcept concept)`
    - `static List<ResolvedFont> ResolveFonts(HeaderAsset asset, DesignConcept concept)`
  - 규칙: 색이 `Locked==true`이거나 컨셉에 해당 역할이 없으면 **원본 value 유지**, 아니면 컨셉 값으로 치환. `concept==null`이면 전부 원본.

- [ ] **Step 1: 실패 테스트 작성**

`src/TeampptAddin.Tests/ConceptResolverTest.cs`:
```csharp
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class ConceptResolverTest
    {
        private static HeaderAsset Asset() => new HeaderAsset
        {
            Colors = new List<AssetColor>
            {
                new AssetColor { Role = "main",  Value = "#2563EB", Locked = false },
                new AssetColor { Role = "logo",  Value = "#FF0000", Locked = true },
                new AssetColor { Role = "text",  Value = "#1E293B", Locked = false }
            },
            Fonts = new List<AssetFont>
            {
                new AssetFont { Role = "heading", Family = "Pretendard" }
            }
        };

        private static DesignConcept Concept() => new DesignConcept
        {
            Colors = new Dictionary<string, string> { ["main"] = "#111827", ["text"] = "#374151" },
            Fonts = new Dictionary<string, string> { ["heading"] = "Noto Sans KR" }
        };

        [Fact]
        public void Unlocked_Role_With_Concept_Value_Is_Replaced()
        {
            var r = ConceptResolver.ResolveColors(Asset(), Concept());
            Assert.Equal("#111827", r.First(c => c.Role == "main").Value);
        }

        [Fact]
        public void Locked_Role_Keeps_Original()
        {
            var r = ConceptResolver.ResolveColors(Asset(), Concept());
            Assert.Equal("#FF0000", r.First(c => c.Role == "logo").Value);
        }

        [Fact]
        public void Role_Missing_From_Concept_Keeps_Original()
        {
            var concept = new DesignConcept { Colors = new Dictionary<string, string> { ["main"] = "#111827" } };
            var r = ConceptResolver.ResolveColors(Asset(), concept);
            Assert.Equal("#1E293B", r.First(c => c.Role == "text").Value);
        }

        [Fact]
        public void Null_Concept_Keeps_All_Original()
        {
            var r = ConceptResolver.ResolveColors(Asset(), null);
            Assert.Equal("#2563EB", r.First(c => c.Role == "main").Value);
        }

        [Fact]
        public void Fonts_Replaced_By_Concept_Family()
        {
            var r = ConceptResolver.ResolveFonts(Asset(), Concept());
            Assert.Equal("Noto Sans KR", r.First(f => f.Role == "heading").Family);
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

Run: `dotnet test ... --filter ConceptResolverTest`
Expected: 컴파일 실패(`ConceptResolver` 미정의).

- [ ] **Step 3: 구현**

`src/TeampptAddin/Services/ConceptResolver.cs`:
```csharp
using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    public class ResolvedColor
    {
        public string Role { get; set; }
        public string Value { get; set; }
    }

    public class ResolvedFont
    {
        public string Role { get; set; }
        public string Family { get; set; }
    }

    public static class ConceptResolver
    {
        public static List<ResolvedColor> ResolveColors(HeaderAsset asset, DesignConcept concept)
        {
            var colors = asset?.Colors ?? new List<AssetColor>();
            return colors.Select(c =>
            {
                var value = c.Value;
                if (!c.Locked && concept?.Colors != null &&
                    concept.Colors.TryGetValue(c.Role, out var conceptValue))
                {
                    value = conceptValue;
                }
                return new ResolvedColor { Role = c.Role, Value = value };
            }).ToList();
        }

        public static List<ResolvedFont> ResolveFonts(HeaderAsset asset, DesignConcept concept)
        {
            var fonts = asset?.Fonts ?? new List<AssetFont>();
            return fonts.Select(f =>
            {
                var family = f.Family;
                if (concept?.Fonts != null &&
                    concept.Fonts.TryGetValue(f.Role, out var conceptFamily))
                {
                    family = conceptFamily;
                }
                return new ResolvedFont { Role = f.Role, Family = family };
            }).ToList();
        }
    }
}
```

- [ ] **Step 4: csproj에 Compile 추가**
```xml
    <Compile Include="Services\ConceptResolver.cs" />
```

- [ ] **Step 5: 통과 확인**

Run: `dotnet test ... --filter ConceptResolverTest`
Expected: PASS (5 tests).

- [ ] **Step 6: 커밋**
```
git add src/TeampptAddin/Services/ConceptResolver.cs src/TeampptAddin/TeampptAddin.csproj src/TeampptAddin.Tests/
git commit -m "feat: ConceptResolver role substitution (respects locked + missing roles)"
```

---

## Task 7: assets.json 을 v2로 재작성 + 통합 스모크

**Files:**
- Modify: `src/TeampptAddin/Assets/assets.json`
- Test: `src/TeampptAddin.Tests/AssetsJsonIntegrationTest.cs`

**Interfaces:**
- Consumes: `AssetLoader.Load`, `CatalogBuilder.Build`.

- [ ] **Step 1: 실패 테스트 작성 (현재 v1 파일 → 7개 + scope 기대)**

`src/TeampptAddin.Tests/AssetsJsonIntegrationTest.cs`:
```csharp
using System.IO;
using System.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetsJsonIntegrationTest
    {
        private static string AssetsDir()
        {
            // 테스트 실행 디렉터리에서 리포 내 Assets 폴더를 찾는다.
            var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (dir != null && !Directory.Exists(Path.Combine(dir.FullName, "src", "TeampptAddin", "Assets")))
                dir = dir.Parent;
            Assert.NotNull(dir);
            return Path.Combine(dir.FullName, "src", "TeampptAddin", "Assets");
        }

        [Fact]
        public void Real_AssetsJson_Loads_As_V2_With_Header_Scope()
        {
            var assets = AssetLoader.Load(AssetsDir());

            Assert.Equal(7, assets.Count);
            Assert.All(assets, a => Assert.Equal(2, a.SchemaVersion));
            // 헤더 1/7 (표지/엔드 성격)은 deck scope 로 표기되어야 한다.
            var hero = assets.First(a => a.File == "header_1.pptx");
            Assert.Equal("deck", hero.Scope);
            Assert.NotEmpty(hero.Slots);

            var catalog = CatalogBuilder.Build(assets);
            Assert.Equal(7, catalog.Count);
        }
    }
}
```

- [ ] **Step 2: 실패 확인**

Run: `dotnet test ... --filter AssetsJsonIntegrationTest`
Expected: FAIL (현재 파일엔 `scope`/`slots` 없음 → `Scope`는 마이그레이터 기본 `"slide"`, `Slots`는 null).

- [ ] **Step 3: assets.json 재작성**

`src/TeampptAddin/Assets/assets.json` 의 7개 항목을 v2로 재작성한다. 각 항목에 `schemaVersion: 2`, `scope`(header_1=타이틀/header_7=마무리는 `"deck"`, 나머지는 본문 성격이면 `"slide"` — 단 헤더 일관성 정책상 반복 헤더는 `"deck"`; 우선 header_1·header_7만 `deck`, 2~6은 `slide`로 시작), `colors`를 역할 배열로, `fonts`(heading/body, Pretendard + fallback "맑은 고딕"), `slots`(최소 `title`/`subtitle`/`body` 중 해당 항목)을 추가한다. 기존 `colors` 객체 값은 그대로 역할 배열로 옮긴다. 예 (header_1):
```jsonc
{
  "schemaVersion": 2,
  "file": "header_1.pptx",
  "name": "타이틀 히어로",
  "category": "헤더",
  "scope": "deck",
  "content_fit": ["제목 강조", "프로젝트 소개", "브랜드 첫인상"],
  "use_when": "프레젠테이션 첫 슬라이드에 임팩트 있는 타이틀이 필요할 때",
  "grid_columns": 1,
  "tags": ["타이틀", "임팩트", "풀블리드", "첫슬라이드"],
  "colors": [
    { "role": "main", "value": "#2563EB", "locked": false },
    { "role": "sub1", "value": "#3B82F6", "locked": false },
    { "role": "sub2", "value": "#93C5FD", "locked": false },
    { "role": "text", "value": "#FFFFFF", "locked": false }
  ],
  "fonts": [
    { "role": "heading", "family": "Pretendard", "fallback": "맑은 고딕", "weight": "Bold", "source": "bundled" },
    { "role": "body",    "family": "Pretendard", "fallback": "맑은 고딕", "weight": "Regular", "source": "bundled" }
  ],
  "slots": [
    { "name": "title",    "type": "text", "perSlide": true },
    { "name": "subtitle", "type": "text", "perSlide": true }
  ]
}
```
나머지 header_2~7도 동일 구조로, 기존 name/category/content_fit/use_when/tags/색값을 보존하며 변환한다.

- [ ] **Step 4: 통과 확인**

Run: `dotnet test ... --filter AssetsJsonIntegrationTest`
Expected: PASS.

- [ ] **Step 5: 전체 테스트 통과 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false`
Expected: 모든 테스트 PASS.

- [ ] **Step 6: 커밋**
```
git add src/TeampptAddin/Assets/assets.json src/TeampptAddin.Tests/
git commit -m "feat: migrate assets.json to schema v2 (roles/fonts/slots/scope)"
```

---

## Self-Review (작성자 체크 완료)

- **Spec 커버리지:** 설계 §5 스키마 v2(역할 색/폰트·슬롯·scope·provenance·schemaVersion) → Task 2. 마이그레이터(저장↔구버전 흡수) → Task 3·4. 컨셉/카탈로그(런타임 컴팩트) → Task 5. 역할 치환 키스톤 → Task 6. 실제 데이터 v2 전환 → Task 7. (B/C/D/E/F는 별도 Phase, 의도적으로 범위 밖.)
- **Placeholder:** 없음 — 모든 코드 스텝에 완전한 코드 포함.
- **타입 일관성:** `AssetColor{Role,Value,Locked}`·`AssetFont{Role,Family,Fallback,Weight,Source}`·`AssetSlot{Name,Type,PerSlide}`·`HeaderAsset.Colors:List<AssetColor>`·`DesignConcept.Colors/Fonts:Dictionary<string,string>`·`CatalogEntry`·`ConceptResolver.ResolveColors/ResolveFonts`가 태스크 간 동일 시그니처로 사용됨.
- **비파괴 확인:** `HeaderAsset.Colors` 소비처 없음(grep 확인) → 타입 교체 안전. `AssetColors` 삭제는 미사용 클래스 제거.

## Execution Handoff

**Plan complete and saved to `docs/superpowers/plans/2026-06-19-phase-a-asset-schema.md`. Two execution options:**

**1. Subagent-Driven (recommended)** — task별 fresh 서브에이전트 디스패치 + 태스크 사이 리뷰, 빠른 반복.

**2. Inline Execution** — 현재 세션에서 executing-plans로 체크포인트 배치 실행.

**Which approach?**

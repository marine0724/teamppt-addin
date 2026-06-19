# 구조화 출력 (Gemini responseSchema) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 응답 JSON 형식을 프롬프트 텍스트로 설명하던 방식을 Gemini `responseSchema`(controlled generation)로 교체하여, 프롬프트 토큰을 줄이고 파싱 안정성을 보장한다.

**Architecture:** `GeminiPromptBuilder`에 응답 스키마를 JObject로 반환하는 `BuildResponseSchema()`를 추가한다. `BuildSystemPrompt`에서는 JSON 형식 설명 블록을 제거하고 의미적 지시만 남긴다. `GeminiAiService`는 요청 본문의 `generationConfig.responseSchema`에 그 스키마를 넣어, 모델이 스키마를 벗어날 수 없게 강제한다. `ParseResponse`는 변경 없음(스키마가 형식을 보장하므로 더 안정적).

**Tech Stack:** .NET Framework 4.8, System.Net.Http.HttpClient, Newtonsoft.Json 13.0.3, xUnit 2.9.2

## Global Constraints

- Target: .NET Framework 4.8 (SDK-style csproj 아님, 수동 Compile Include — 이 계획은 신규 파일이 없으므로 csproj 변경 없음)
- DI 프레임워크 없음, 직접 인스턴스화
- Newtonsoft.Json 13.0.3 사용 (System.Text.Json 아님)
- 테스트: xUnit 2.9.2, net48
- 한국어 UI/프롬프트
- COM 애드인(PowerPoint 인프로세스) — 의존성 추가 금지, HttpClient 직접 사용 유지
- Gemini `responseSchema`는 OpenAPI 3.0 서브셋. nullable 필드는 `"nullable": true` 사용

---

### Task 1: GeminiPromptBuilder에 BuildResponseSchema() 추가 (additive)

이 태스크는 순수 추가다. 기존 `BuildSystemPrompt`는 건드리지 않으므로 커밋 후에도 런타임 동작이 변하지 않는다.

**Files:**
- Modify: `src/TeampptAddin/Services/GeminiPromptBuilder.cs`
- Modify: `src/TeampptAddin.Tests/GeminiPromptBuilderTest.cs`

**Interfaces:**
- Consumes: 없음 (정적 메서드, 인자 없음)
- Produces: `GeminiPromptBuilder.BuildResponseSchema()` → `Newtonsoft.Json.Linq.JObject` — Gemini `generationConfig.responseSchema`에 넣을 스키마 객체

- [ ] **Step 1: 실패하는 테스트 작성**

`src/TeampptAddin.Tests/GeminiPromptBuilderTest.cs`의 클래스 안(기존 `UserPrompt_Contains_Intent` 테스트 뒤)에 추가:

```csharp
        [Fact]
        public void ResponseSchema_Has_Expected_Top_Level_Properties()
        {
            var schema = GeminiPromptBuilder.BuildResponseSchema();

            Assert.Equal("object", schema["type"]?.ToString());
            var props = schema["properties"];
            Assert.NotNull(props["message"]);
            Assert.NotNull(props["assets"]);
            Assert.NotNull(props["palette"]);
            Assert.NotNull(props["font"]);
        }

        [Fact]
        public void ResponseSchema_Assets_Is_Array_Of_File_Reason()
        {
            var schema = GeminiPromptBuilder.BuildResponseSchema();
            var assets = schema["properties"]["assets"];

            Assert.Equal("array", assets["type"]?.ToString());
            var itemProps = assets["items"]["properties"];
            Assert.NotNull(itemProps["file"]);
            Assert.NotNull(itemProps["reason"]);
        }

        [Fact]
        public void ResponseSchema_Palette_And_Font_Are_Nullable()
        {
            var schema = GeminiPromptBuilder.BuildResponseSchema();

            Assert.True(schema["properties"]["palette"]["nullable"].Value<bool>());
            Assert.True(schema["properties"]["font"]["nullable"].Value<bool>());
        }
```

또한 파일 상단 using에 `Newtonsoft.Json.Linq`가 없으면 추가:

```csharp
using Newtonsoft.Json.Linq;
```

- [ ] **Step 2: 테스트 실패 확인**

```powershell
dotnet test src\TeampptAddin.Tests --filter "FullyQualifiedName~GeminiPromptBuilderTest" --no-restore -v minimal
```

Expected: 컴파일 에러 — `BuildResponseSchema` 정의 없음

- [ ] **Step 3: BuildResponseSchema 구현**

`src/TeampptAddin/Services/GeminiPromptBuilder.cs` 상단 using에 추가:

```csharp
using Newtonsoft.Json.Linq;
```

`BuildUserPrompt` 메서드 뒤에 추가:

```csharp
        /// <summary>
        /// Gemini generationConfig.responseSchema에 넣을 응답 스키마.
        /// 모델이 이 구조를 벗어날 수 없게 강제하므로, 프롬프트에 형식을 설명할 필요가 없다.
        /// palette/font는 질문/부적합 케이스에서 null이 될 수 있어 nullable.
        /// </summary>
        public static JObject BuildResponseSchema()
        {
            return new JObject
            {
                ["type"] = "object",
                ["properties"] = new JObject
                {
                    ["message"] = new JObject { ["type"] = "string" },
                    ["assets"] = new JObject
                    {
                        ["type"] = "array",
                        ["items"] = new JObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JObject
                            {
                                ["file"] = new JObject { ["type"] = "string" },
                                ["reason"] = new JObject { ["type"] = "string" }
                            },
                            ["required"] = new JArray { "file", "reason" }
                        }
                    },
                    ["palette"] = new JObject
                    {
                        ["type"] = "object",
                        ["nullable"] = true,
                        ["properties"] = new JObject
                        {
                            ["id"] = new JObject { ["type"] = "string" },
                            ["reason"] = new JObject { ["type"] = "string" }
                        }
                    },
                    ["font"] = new JObject
                    {
                        ["type"] = "object",
                        ["nullable"] = true,
                        ["properties"] = new JObject
                        {
                            ["name"] = new JObject { ["type"] = "string" },
                            ["reason"] = new JObject { ["type"] = "string" }
                        }
                    }
                },
                ["required"] = new JArray { "message", "assets" }
            };
        }
```

- [ ] **Step 4: 테스트 통과 확인**

```powershell
dotnet test src\TeampptAddin.Tests --filter "FullyQualifiedName~GeminiPromptBuilderTest" --no-restore -v minimal
```

Expected: 기존 + 신규 3개 모두 PASS

- [ ] **Step 5: Commit**

```bash
git add src/TeampptAddin/Services/GeminiPromptBuilder.cs src/TeampptAddin.Tests/GeminiPromptBuilderTest.cs
git commit -m "feat: add BuildResponseSchema for Gemini structured output"
```

---

### Task 2: 프롬프트에서 형식 블록 제거 + 요청에 responseSchema 연결

이 태스크에서 두 변경이 함께 일어나야 일관성이 유지된다(프롬프트의 형식 설명 제거 ↔ 스키마 강제). 그래서 한 태스크로 묶는다.

**Files:**
- Modify: `src/TeampptAddin/Services/GeminiPromptBuilder.cs` (BuildSystemPrompt에서 응답 형식 블록 제거)
- Modify: `src/TeampptAddin/Services/GeminiAiService.cs` (generationConfig에 responseSchema 추가)
- Modify: `src/TeampptAddin.Tests/GeminiPromptBuilderTest.cs` (형식 블록 부재 검증으로 교체)

**Interfaces:**
- Consumes: `GeminiPromptBuilder.BuildResponseSchema()` → `JObject` (Task 1)
- Produces: 변경된 `BuildSystemPrompt`(형식 블록 없음), responseSchema가 포함된 Gemini 요청 본문

- [ ] **Step 1: 기존 형식 검증 테스트를 "형식 블록 부재" 검증으로 교체**

`src/TeampptAddin.Tests/GeminiPromptBuilderTest.cs`에서 기존 `SystemPrompt_Contains_Json_Response_Schema` 테스트를 아래로 **교체**:

```csharp
        [Fact]
        public void SystemPrompt_Does_Not_Embed_Json_Skeleton()
        {
            var prompt = GeminiPromptBuilder.BuildSystemPrompt(
                MakeCatalog(), MakePalettes(), MakeFonts());

            // 형식은 responseSchema가 강제하므로 프롬프트에 JSON 골격을 넣지 않는다
            Assert.DoesNotContain("```json", prompt);
            Assert.DoesNotContain("\"reason\":", prompt);
        }
```

- [ ] **Step 2: 테스트 실패 확인**

```powershell
dotnet test src\TeampptAddin.Tests --filter "FullyQualifiedName~SystemPrompt_Does_Not_Embed_Json_Skeleton" --no-restore -v minimal
```

Expected: FAIL — 현재 프롬프트에는 아직 ```json 블록이 있음

- [ ] **Step 3: BuildSystemPrompt에서 응답 형식 블록 제거**

`src/TeampptAddin/Services/GeminiPromptBuilder.cs`의 `BuildSystemPrompt` return 문에서, `## 핵심 원칙` 섹션까지만 남기고 `## 응답 형식` 이하 전체를 제거. 최종 return 문은 다음과 같이:

```csharp
            return $@"너는 PPT 디자인 어시스턴트야. 사용자와 대화하며 적합한 에셋과 스타일을 추천해.

## 에셋 카탈로그
{catalogJson}

## 팔레트 목록
{palettesJson}

## 폰트 목록
{fontsJson}

## 핵심 원칙
- 카탈로그에 있는 에셋만 추천할 수 있다. 없는 에셋을 지어내지 마.
- 사용자의 의도와 에셋의 use_when/content_fit/tags를 비교해서, 실제로 적합할 때만 추천해.
- 적합한 에셋이 없으면 솔직하게 ""현재 보유한 에셋 중에는 딱 맞는 것이 없다""고 말해. 이때 assets는 빈 배열, palette/font는 null로 둬.
- 사용자의 요청이 모호하면 바로 추천하지 말고, 먼저 질문해서 의도를 파악해.
- message는 한국어 1~2문장. 각 에셋 추천에는 구체적 reason을 달아.";
```

> 형식(JSON 키/구조)은 이제 `responseSchema`가 강제하므로 프롬프트에서 설명하지 않는다. 의미적 지시(언제 빈 배열/null인지, message 길이)만 남긴다.

- [ ] **Step 4: GeminiAiService 요청에 responseSchema 추가**

`src/TeampptAddin/Services/GeminiAiService.cs`의 `generationConfig` JObject(현재 70행 부근)를 다음으로 변경:

변경 전:
```csharp
                ["generationConfig"] = new JObject
                {
                    ["temperature"] = 0.7,
                    ["responseMimeType"] = "application/json",
                    ["thinkingConfig"] = new JObject { ["thinkingBudget"] = 1024 }
                }
```

변경 후:
```csharp
                ["generationConfig"] = new JObject
                {
                    ["temperature"] = 0.7,
                    ["responseMimeType"] = "application/json",
                    ["responseSchema"] = GeminiPromptBuilder.BuildResponseSchema(),
                    ["thinkingConfig"] = new JObject { ["thinkingBudget"] = 1024 }
                }
```

- [ ] **Step 5: 전체 테스트 통과 확인**

```powershell
dotnet test src\TeampptAddin.Tests --no-restore -v minimal
```

Expected: 모든 테스트 PASS (교체한 테스트 포함)

- [ ] **Step 6: 빌드 확인**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" src\TeampptAddin\TeampptAddin.csproj /t:Build /p:Configuration=Debug /v:minimal
```

Expected: 빌드 성공

- [ ] **Step 7: Commit**

```bash
git add src/TeampptAddin/Services/GeminiPromptBuilder.cs src/TeampptAddin/Services/GeminiAiService.cs src/TeampptAddin.Tests/GeminiPromptBuilderTest.cs
git commit -m "feat: enforce response shape via responseSchema, drop prompt format block"
```

---

### Task 3: PowerPoint 수동 검증 + 토큰/지연 측정

코드가 아니라 실제 동작·효과를 확인하는 태스크. 토큰 절감과 추천 속도 개선을 눈으로 검증한다.

**Files:** 없음 (측정만)

- [ ] **Step 1: 관리자 권한 빌드 + 등록**

```powershell
Start-Process -Verb RunAs -FilePath "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" -ArgumentList "src\TeampptAddin\TeampptAddin.csproj /t:Build /p:Configuration=Debug /v:minimal"
```

- [ ] **Step 2: PowerPoint에서 추천 요청 → 응답 형식 검증**

1. PowerPoint 열기 → Task Pane → AI 채팅에 "투자 유치용 깔끔한 표지" 입력
2. 에셋 카드 + 스타일 추천이 정상 표시되는지 확인
3. 이어서 "음... 잘 모르겠어" 같은 모호한 입력 → assets 빈 배열/질문 응답이 정상인지 확인 (nullable 동작)

- [ ] **Step 3: 토큰 사용량 before/after 비교**

```powershell
Get-Content "$env:LOCALAPPDATA\TeampptAddin\debug.log" -Tail 10
```

Expected: `[Gemini] 토큰 사용: input=...` 의 input 토큰이 형식 블록 제거 전보다 줄었는지 확인 (형식 설명 ~25줄만큼 감소 기대)

- [ ] **Step 4: (선택, 튜닝) thinkingBudget 조정으로 추천 속도 개선 시도**

추천 응답이 여전히 느리면 `GeminiAiService.cs`의 `["thinkingBudget"] = 1024`를 `512`로 낮춰 빌드 후 다시 측정. 응답 품질이 유지되는 최저값을 찾는다. 품질 저하가 느껴지면 원복.

> 주의: 이건 측정 기반 튜닝이다. 임의로 0으로 끄지 말고, Step 2의 추천 품질을 보면서 단계적으로 낮춘다.

- [ ] **Step 5: 변경이 있으면 Commit**

```bash
git add src/TeampptAddin/Services/GeminiAiService.cs
git commit -m "perf: tune thinkingBudget for faster recommendations"
```

---

## Self-Review

**Spec coverage:** 이 계획은 대화에서 확정한 5개 결정 중 "구조화 출력(responseSchema)" 하나만 다룬다. 나머지(Supabase 저장, 벡터 검색, 대화 기억)는 의도적으로 제외 — 별도 설계 스펙 + 코드 계획으로 분리(scope-check). 구조화 출력 범위 안에서: 스키마 정의(Task 1), 프롬프트 형식 제거 + 요청 연결(Task 2), 효과 검증(Task 3) 모두 커버됨.

**Placeholder scan:** 모든 코드 스텝에 실제 코드 포함. TODO/TBD 없음.

**Type consistency:** `BuildResponseSchema()` → `JObject` 가 Task 1에서 정의되고 Task 2 Step 4에서 동일 시그니처로 소비됨. `ParseResponse`는 변경 없음 — 스키마의 키(message/assets/file/reason/palette.id/font.name)가 기존 ParseResponse가 읽는 키와 일치함(GeminiAiService.cs:124-161 확인).

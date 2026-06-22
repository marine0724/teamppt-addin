# A-1d · 벡터검색 읽기경로 + 추천 연결 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 사용자의 텍스트 질의("장점 3가지 나열할 에셋 추천해줘")를 임베딩 → Supabase `match_assets` 벡터검색(top 8, anon) → 후보를 Gemini Flash가 2~3개 선택·설명 → AI탭에 추천. Supabase 불통 시 마지막 캐시 → 번들 폴백. (A-3 ②"텍스트 질의 추천" = A의 첫 시연 깃발.)

**Architecture:** 순수 로직(row→HeaderAsset 매퍼, RPC 인자/결과 파싱, 추천 캐시 직렬화)은 xUnit TDD. 외부 HTTP(임베딩·RPC·Storage 다운로드)는 어댑터로 두고 수동 검증. 후보 선택·설명은 **기존 [GeminiAiService](../../../src/TeampptAddin/Services/GeminiAiService.cs)를 그대로 재사용**(top-8 후보만 넘기면 CatalogBuilder가 컴팩트 카탈로그로 줄여 저토큰 매칭). 새 `VectorRecommendService`가 `IAiService`를 구현해 기존 AI탭 흐름([AssetPanel.RecommendAsync](../../../src/TeampptAddin/UI/Wpf/AssetPanel.cs#L503))을 무변경으로 잇는다.

**Tech Stack:** .NET Framework 4.8, Newtonsoft.Json 13.0.3, HttpClient, Supabase REST/RPC/Storage, Gemini Flash + text-embedding-004, xUnit 2.9.

## Global Constraints

- Core/Connect.cs/Globals.cs 직접 수정 금지. (TaskPaneHost 배선만 최소 수정 — Task 6.)
- 의존성 추가 금지. REST는 HttpClient 직접. `IAiService` 인터페이스 고정(시그니처 변경 금지).
- 읽기는 **anon 키**(api-keys.json) + RLS 읽기전용. service-role 사용 금지(읽기 경로).
- top-N = **8**, 하드 유사도 컷 없음(점수는 로깅만). 데이터 늘면 컷 도입.
- 선행: A-1a(이해), A-1b/c(적재). Supabase `assets`에 행+벡터, `match_assets` RPC, `thumb`/`pptx` 버킷 존재.
- 단위테스트 절차/MSBuild = A-1a plan과 동일.

---

## File Structure

| 파일 | 책임 | 테스트 |
|---|---|---|
| `Services/SupabaseAssetMapper.cs` (신규) | `match_assets` row(JObject) → `HeaderAsset`(remote thumb/file는 Extra) (순수) | `SupabaseAssetMapperTest.cs` |
| `Services/MatchQuery.cs` (신규) | RPC 인자 빌더 + 결과 배열 파싱 → `List<HeaderAsset>` (순수) | `MatchQueryTest.cs` |
| `Services/RecommendationCache.cs` (신규) | 마지막 성공 추천 후보 저장/로드(오프라인 폴백) (순수, 파일) | `RecommendationCacheTest.cs` |
| `Services/RemoteAssetCache.cs` (신규) | Storage thumb/pptx → 로컬 캐시 다운로드, 로컬 경로 반환 (HTTP) | 수동 |
| `Services/VectorRecommendService.cs` (신규) | `IAiService`: 임베딩→RPC→후보→GeminiAiService 위임→캐시/폴백 (HTTP 오케스트레이션) | 수동 |
| `UI/TaskPaneHost.cs` (수정 165~179) | supabase 설정 있으면 `VectorRecommendService` 주입, 없으면 기존 폴백 | 수동 |

순수 3개(Mapper/MatchQuery/RecommendationCache)가 단위테스트 산출물. HTTP 2개 + 배선은 수동 검증.

---

### Task 1: row → HeaderAsset 매퍼

**Files:**
- Create: `src/TeampptAddin/Services/SupabaseAssetMapper.cs`
- Test: `src/TeampptAddin.Tests/SupabaseAssetMapperTest.cs`

**Interfaces:**
- Consumes: `HeaderAsset`, `AssetColor/Font/Slot`.
- Produces: `static HeaderAsset SupabaseAssetMapper.Map(JObject row)`. `match_assets` 결과 1행을 `HeaderAsset`로. `metadata.colors/fonts/slots` 복원. 원격 경로(`file`="pptx/..", `thumb`="thumb/..")는 삽입/표시용으로 `Extra["remote_file"]`/`Extra["remote_thumb"]`에 보관. `HeaderAsset.File`엔 파일명만(예: `표지_01.pptx`).

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using Newtonsoft.Json.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class SupabaseAssetMapperTest
    {
        private const string Row = @"{
          ""file"":""pptx/표지_01.pptx"", ""thumb"":""thumb/표지_01.png"",
          ""name"":""연도강조 표지"", ""category"":""표지"", ""kind"":""layout"", ""scope"":""slide"",
          ""tags"":[""표지"",""연도""], ""use_when"":""연도 강조"", ""content_fit"":[""표지""],
          ""metadata"":{ ""colors"":[{""role"":""main"",""value"":""#1A2B4C"",""locked"":false}],
                         ""fonts"":[{""role"":""heading"",""family"":""Pretendard""}],
                         ""slots"":[{""name"":""title"",""type"":""text"",""perSlide"":true}] },
          ""similarity"":0.83 }";

        [Fact]
        public void Map_Core_Fields_And_Filename()
        {
            var a = SupabaseAssetMapper.Map(JObject.Parse(Row));
            Assert.Equal("연도강조 표지", a.Name);
            Assert.Equal("layout", a.Kind);
            Assert.Equal("표지", a.Category);
            Assert.Equal("표지_01.pptx", a.File);   // 파일명만
        }

        [Fact]
        public void Map_Restores_Metadata_Structures()
        {
            var a = SupabaseAssetMapper.Map(JObject.Parse(Row));
            Assert.Equal("#1A2B4C", a.Colors[0].Value);
            Assert.Equal("Pretendard", a.Fonts[0].Family);
            Assert.Equal("title", a.Slots[0].Name);
        }

        [Fact]
        public void Map_Keeps_Remote_Paths_In_Extra()
        {
            var a = SupabaseAssetMapper.Map(JObject.Parse(Row));
            Assert.Equal("pptx/표지_01.pptx", a.Extra["remote_file"].ToString());
            Assert.Equal("thumb/표지_01.png", a.Extra["remote_thumb"].ToString());
        }
    }
}
```

- [ ] **Step 2: 실패 확인** — 테스트 실행. Expected: FAIL.

- [ ] **Step 3: 최소 구현**

```csharp
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>match_assets RPC 행 → HeaderAsset. 원격 경로는 Extra에 보관(삽입/표시용).</summary>
    public static class SupabaseAssetMapper
    {
        public static HeaderAsset Map(JObject row)
        {
            var meta = row["metadata"] as JObject ?? new JObject();
            var remoteFile = row["file"]?.ToString() ?? "";
            var remoteThumb = row["thumb"]?.ToString() ?? "";

            return new HeaderAsset
            {
                SchemaVersion = 2,
                File = Path.GetFileName(remoteFile),
                Name = row["name"]?.ToString(),
                Kind = row["kind"]?.ToString() ?? "component",
                Category = row["category"]?.ToString(),
                Scope = row["scope"]?.ToString() ?? "slide",
                UseWhen = row["use_when"]?.ToString(),
                Tags = StrList(row["tags"]),
                ContentFit = StrList(row["content_fit"]),
                Colors = (meta["colors"] as JArray)?.Select(c => new AssetColor
                {
                    Role = c["role"]?.ToString(), Value = c["value"]?.ToString(),
                    Locked = c["locked"]?.Value<bool>() ?? false
                }).ToList() ?? new List<AssetColor>(),
                Fonts = (meta["fonts"] as JArray)?.Select(f => new AssetFont
                {
                    Role = f["role"]?.ToString(), Family = f["family"]?.ToString(), Weight = f["weight"]?.ToString()
                }).ToList() ?? new List<AssetFont>(),
                Slots = (meta["slots"] as JArray)?.Select(s => new AssetSlot
                {
                    Name = s["name"]?.ToString(), Type = s["type"]?.ToString(),
                    PerSlide = s["perSlide"]?.Value<bool>() ?? false
                }).ToList() ?? new List<AssetSlot>(),
                Extra = new Dictionary<string, JToken>
                {
                    ["remote_file"] = remoteFile,
                    ["remote_thumb"] = remoteThumb
                }
            };
        }

        private static List<string> StrList(JToken t) =>
            (t as JArray)?.Select(x => x.ToString()).ToList() ?? new List<string>();
    }
}
```

- [ ] **Step 4: 통과 확인** — 테스트 실행. Expected: 3 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/SupabaseAssetMapper.cs src/TeampptAddin.Tests/SupabaseAssetMapperTest.cs
git commit -m "feat(read): SupabaseAssetMapper (RPC row→HeaderAsset, 원격경로 Extra)"
```

---

### Task 2: 벡터검색 쿼리 (인자 빌더 + 결과 파싱)

**Files:**
- Create: `src/TeampptAddin/Services/MatchQuery.cs`
- Test: `src/TeampptAddin.Tests/MatchQueryTest.cs`

**Interfaces:**
- Consumes: `SupabaseAssetMapper`.
- Produces: `static JObject MatchQuery.BuildArgs(float[] queryEmbedding, int matchCount)` → `{ query_embedding: "[...]", match_count: N }`. `static List<HeaderAsset> MatchQuery.ParseResults(string rpcJson)` → 행 배열을 매퍼로 변환(유사도 점수는 Logger로). 빈 결과는 빈 리스트.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using Xunit;

namespace TeampptAddin.Tests
{
    public class MatchQueryTest
    {
        [Fact]
        public void BuildArgs_Formats_Vector_And_Count()
        {
            var args = MatchQuery.BuildArgs(new float[] { 0.1f, 0.2f }, 8);
            Assert.Equal("[0.1,0.2]", args["query_embedding"]);
            Assert.Equal(8, args["match_count"]);
        }

        [Fact]
        public void ParseResults_Maps_Each_Row()
        {
            var json = @"[
              {""file"":""pptx/a.pptx"",""thumb"":""thumb/a.png"",""name"":""A"",""category"":""표지"",""kind"":""layout"",""metadata"":{},""similarity"":0.9},
              {""file"":""pptx/b.pptx"",""thumb"":""thumb/b.png"",""name"":""B"",""category"":""목차"",""kind"":""component"",""metadata"":{},""similarity"":0.7}
            ]";
            var list = MatchQuery.ParseResults(json);
            Assert.Equal(2, list.Count);
            Assert.Equal("A", list[0].Name);
            Assert.Equal("b.pptx", list[1].File);
        }

        [Fact]
        public void ParseResults_Empty_Returns_Empty()
        {
            Assert.Empty(MatchQuery.ParseResults("[]"));
        }
    }
}
```

- [ ] **Step 2: 실패 확인** — 테스트 실행. Expected: FAIL.

- [ ] **Step 3: 최소 구현**

```csharp
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>match_assets RPC 인자/결과 변환.</summary>
    public static class MatchQuery
    {
        public static JObject BuildArgs(float[] queryEmbedding, int matchCount)
        {
            var vec = "[" + string.Join(",", queryEmbedding.Select(v => v.ToString(CultureInfo.InvariantCulture))) + "]";
            return new JObject { ["query_embedding"] = vec, ["match_count"] = matchCount };
        }

        public static List<HeaderAsset> ParseResults(string rpcJson)
        {
            var arr = JArray.Parse(rpcJson);
            var result = new List<HeaderAsset>();
            foreach (var row in arr.OfType<JObject>())
            {
                var sim = row["similarity"]?.Value<double>() ?? 0;
                Logger.Log($"[Match] {row["name"]} sim={sim:F3}");
                result.Add(SupabaseAssetMapper.Map(row));
            }
            return result;
        }
    }
}
```

- [ ] **Step 4: 통과 확인** — 테스트 실행. Expected: 3 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/MatchQuery.cs src/TeampptAddin.Tests/MatchQueryTest.cs
git commit -m "feat(read): MatchQuery (RPC 인자/결과 파싱)"
```

---

### Task 3: 추천 캐시 (오프라인 폴백)

**Files:**
- Create: `src/TeampptAddin/Services/RecommendationCache.cs`
- Test: `src/TeampptAddin.Tests/RecommendationCacheTest.cs`

**Interfaces:**
- Produces: `class RecommendationCache { RecommendationCache(string path); void Save(List<HeaderAsset> candidates); List<HeaderAsset> Load(); static string DefaultPath }`. 마지막 성공 후보를 JSON으로 저장/로드. 파일 없거나 파싱 실패 시 빈 리스트(throw 안 함). `DefaultPath` = `%LOCALAPPDATA%\TeampptAddin\cache\last-candidates.json`.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace TeampptAddin.Tests
{
    public class RecommendationCacheTest
    {
        [Fact]
        public void Save_Then_Load_Roundtrips()
        {
            var tmp = Path.Combine(Path.GetTempPath(), "rc_" + System.Guid.NewGuid() + ".json");
            try
            {
                var cache = new RecommendationCache(tmp);
                cache.Save(new List<HeaderAsset> { new HeaderAsset { Name = "A", File = "a.pptx", Kind = "layout" } });
                var loaded = cache.Load();
                Assert.Single(loaded);
                Assert.Equal("A", loaded[0].Name);
            }
            finally { if (File.Exists(tmp)) File.Delete(tmp); }
        }

        [Fact]
        public void Load_Missing_File_Returns_Empty()
        {
            var cache = new RecommendationCache(Path.Combine(Path.GetTempPath(), "nope_" + System.Guid.NewGuid() + ".json"));
            Assert.Empty(cache.Load());
        }
    }
}
```

- [ ] **Step 2: 실패 확인** — 테스트 실행. Expected: FAIL.

- [ ] **Step 3: 최소 구현**

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace TeampptAddin
{
    /// <summary>마지막 성공 추천 후보 캐시. Supabase 불통 시 폴백 1순위(번들이 2순위).</summary>
    public class RecommendationCache
    {
        private readonly string _path;
        public RecommendationCache(string path = null) { _path = path ?? DefaultPath; }

        public static string DefaultPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TeampptAddin", "cache", "last-candidates.json");

        public void Save(List<HeaderAsset> candidates)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path));
                File.WriteAllText(_path, JsonConvert.SerializeObject(candidates));
            }
            catch (Exception ex) { Logger.Log($"[Cache] save 실패: {ex.Message}"); }
        }

        public List<HeaderAsset> Load()
        {
            try
            {
                if (!File.Exists(_path)) return new List<HeaderAsset>();
                return JsonConvert.DeserializeObject<List<HeaderAsset>>(File.ReadAllText(_path))
                       ?? new List<HeaderAsset>();
            }
            catch (Exception ex) { Logger.Log($"[Cache] load 실패: {ex.Message}"); return new List<HeaderAsset>(); }
        }
    }
}
```

- [ ] **Step 4: 통과 확인** — 테스트 실행. Expected: 2 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/RecommendationCache.cs src/TeampptAddin.Tests/RecommendationCacheTest.cs
git commit -m "feat(read): RecommendationCache (오프라인 폴백)"
```

---

### Task 4: 원격 에셋 캐시 (Storage 다운로드, 수동)

**Files:**
- Create: `src/TeampptAddin/Services/RemoteAssetCache.cs`

**Interfaces:**
- Produces: `class RemoteAssetCache { RemoteAssetCache(string baseUrl, string anonKey); Task<string> GetThumbAsync(string remoteThumb); Task<string> GetPptxAsync(string remoteFile); }`. public 버킷에서 `{baseUrl}/storage/v1/object/public/{path}` GET → `%LOCALAPPDATA%\TeampptAddin\cache\{thumb|pptx}\파일명`에 저장 후 로컬 경로 반환. 이미 있으면 재다운로드 생략.

- [ ] **Step 1: 구현 작성**

```csharp
using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace TeampptAddin
{
    /// <summary>Supabase Storage(public) → 로컬 캐시 다운로드. 표시=thumb, 삽입=pptx. 재다운로드 생략.</summary>
    public class RemoteAssetCache
    {
        private static readonly HttpClient Http = new HttpClient();
        private readonly string _baseUrl;
        private readonly string _anonKey;

        public RemoteAssetCache(string baseUrl, string anonKey)
        {
            _baseUrl = baseUrl.TrimEnd('/');
            _anonKey = anonKey;
        }

        public Task<string> GetThumbAsync(string remoteThumb) => GetAsync(remoteThumb, "thumb");
        public Task<string> GetPptxAsync(string remoteFile) => GetAsync(remoteFile, "pptx");

        private async Task<string> GetAsync(string remotePath, string subdir)
        {
            var fileName = Path.GetFileName(remotePath);
            var localDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TeampptAddin", "cache", subdir);
            Directory.CreateDirectory(localDir);
            var localPath = Path.Combine(localDir, fileName);
            if (File.Exists(localPath)) return localPath;

            var url = $"{_baseUrl}/storage/v1/object/public/{remotePath}";
            var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.TryAddWithoutValidation("apikey", _anonKey);
            var resp = await Http.SendAsync(req).ConfigureAwait(false);
            Logger.Log($"[RemoteCache] GET {remotePath}: HTTP {(int)resp.StatusCode}");
            resp.EnsureSuccessStatusCode();
            File.WriteAllBytes(localPath, await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(false));
            return localPath;
        }
    }
}
```

- [ ] **Step 2: 빌드 확인** — Run: A-1a MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Services/RemoteAssetCache.cs
git commit -m "feat(read): RemoteAssetCache (Storage→로컬 캐시 다운로드)"
```

---

### Task 5: 벡터 추천 서비스 (IAiService, 수동)

**Files:**
- Create: `src/TeampptAddin/Services/VectorRecommendService.cs`

**Interfaces:**
- Consumes: `EmbeddingService`, `SupabaseClient`(anon), `MatchQuery`, `RecommendationCache`, `GeminiAiService`(후보 선택 위임), `IAiService`.
- Produces: `class VectorRecommendService : IAiService`. 생성자 `(string supabaseUrl, string anonKey, string geminiKey)`. `RecommendAsync(intent, assets, palettes, fonts)`: ① intent 임베딩 ② `match_assets`(top 8) → 후보 `List<HeaderAsset>` ③ 성공 시 캐시 저장; 실패 시 캐시→인자 `assets`(번들) 폴백 ④ 내부 `GeminiAiService`에 **후보만** 넘겨 2~3개 선택·설명 위임. (인자 `assets`는 폴백 후보로만 사용 — 평소엔 Supabase 후보가 우선.)

- [ ] **Step 1: 구현 작성**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TeampptAddin
{
    /// <summary>
    /// A-3 ② 텍스트 질의 추천: 임베딩→Supabase 벡터검색(top8)→후보를 GeminiAiService가 2~3개 선택·설명.
    /// Supabase 불통 시 마지막 캐시 → 번들(인자 assets) 폴백 → "추천 빈손 없음".
    /// </summary>
    public class VectorRecommendService : IAiService
    {
        private const int TopN = 8;
        private readonly EmbeddingService _embed;
        private readonly SupabaseClient _supa;
        private readonly GeminiAiService _selector;   // 후보 선택·설명 + 세션 메모리(history) 보유
        private readonly RecommendationCache _cache = new RecommendationCache();

        public VectorRecommendService(string supabaseUrl, string anonKey, string geminiKey)
        {
            _embed = new EmbeddingService(geminiKey);
            _supa = new SupabaseClient(supabaseUrl, anonKey);
            _selector = new GeminiAiService(geminiKey);
        }

        public async Task<AiRecommendation> RecommendAsync(
            string userIntent,
            IEnumerable<HeaderAsset> assets,
            IEnumerable<StylePalette> palettes,
            IEnumerable<StyleFont> fonts)
        {
            List<HeaderAsset> candidates;
            try
            {
                var vector = await _embed.EmbedAsync(userIntent).ConfigureAwait(false);
                var rpcJson = await _supa.RpcAsync("match_assets", MatchQuery.BuildArgs(vector, TopN)).ConfigureAwait(false);
                candidates = MatchQuery.ParseResults(rpcJson);
                if (candidates.Count > 0) _cache.Save(candidates);
                Logger.Log($"[VectorRec] 후보 {candidates.Count}개");
            }
            catch (Exception ex)
            {
                Logger.Log($"[VectorRec] Supabase 실패 → 폴백: {ex.Message}");
                candidates = _cache.Load();
                if (candidates.Count == 0) candidates = (assets ?? Enumerable.Empty<HeaderAsset>()).ToList();
            }

            // 후보를 GeminiAiService에 위임 → 2~3개 선택·설명 (CatalogBuilder가 컴팩트화)
            return await _selector.RecommendAsync(userIntent, candidates, palettes, fonts).ConfigureAwait(false);
        }
    }
}
```

- [ ] **Step 2: 본프로젝트 빌드 확인** — Run: A-1a MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 전체 단위테스트 GREEN** — "테스트 실행 절차". Expected: 기존 + 신규(SupabaseAssetMapper 3, MatchQuery 3, RecommendationCache 2) PASS.

- [ ] **Step 4: 커밋**

```
git add src/TeampptAddin/Services/VectorRecommendService.cs
git commit -m "feat(read): VectorRecommendService (임베딩→벡터검색→Gemini 선택, 폴백)"
```

---

### Task 6: AI탭 배선 + PowerPoint 수동 검증

**Files:**
- Modify: `src/TeampptAddin/UI/TaskPaneHost.cs:165-179`

**Interfaces:**
- Consumes: `VectorRecommendService`, 기존 `GeminiAiService`/`MockAiService`.
- Produces: `LoadWpfCards`가 api-keys.json에 `supabaseUrl`+`supabaseAnonKey`+`gemini`가 있으면 `VectorRecommendService`를 주입, 없으면 기존 폴백 유지. (인터페이스 `IAiService` 무변경이라 AssetPanel 흐름은 그대로.)

- [ ] **Step 1: 배선 수정**

[TaskPaneHost.cs:167-176](../../../src/TeampptAddin/UI/TaskPaneHost.cs#L167)의 `IAiService ai; try {...} catch {...}` 블록을 아래로 교체:

```csharp
            IAiService ai;
            try
            {
                var keysPath = Path.Combine(assetsDir, "api-keys.json");
                var keys = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(keysPath));
                var gemini = keys["gemini"]?.ToString();
                var supaUrl = keys["supabaseUrl"]?.ToString();
                var supaAnon = keys["supabaseAnonKey"]?.ToString();

                if (!string.IsNullOrEmpty(supaUrl) && !string.IsNullOrEmpty(supaAnon) && !string.IsNullOrEmpty(gemini))
                {
                    ai = new VectorRecommendService(supaUrl, supaAnon, gemini);
                    Logger.Log("[AI] VectorRecommendService (Supabase 벡터검색) 사용");
                }
                else
                {
                    ai = GeminiAiService.FromAssetsDir(assetsDir);
                    Logger.Log("[AI] Supabase 설정 없음 → GeminiAiService(로컬 카탈로그) 사용");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[AI] 초기화 실패, MockAiService 사용: {ex.Message}");
                ai = new MockAiService();
            }
```

- [ ] **Step 2: 관리자 권한 빌드 (COM 등록 — 실제 구동)**

PowerPoint 완전 종료 후, 관리자 빌드로 애드인 갱신(이 변경은 UI 구동 경로라 실구동 검증 필요). 빌드 절차는 프로젝트 표준(HANDOFF §4.2~4.3) 따름. Expected: Build succeeded + 애드인 로드.

- [ ] **Step 3: PowerPoint 수동 검증 (읽기경로 깃발)**

전제: A-1b/c로 Supabase에 에셋 적재됨, api-keys.json에 supabase 설정 있음.
1. PowerPoint 실행 → TEAMPPT 패널 토글 → AI탭.
2. *"장점 3가지 나열할 에셋 추천해줘"* 입력.
3. 확인: 채팅에 추천 2~3개 + 각 reason 표시. `debug.log`에 `[VectorRec] 후보 8개`, `[Match] ... sim=...` 기록.
4. **오프라인 검증:** 네트워크 끊고 다시 질의 → 캐시 폴백으로 추천이 빈손 없이 나옴(`[VectorRec] Supabase 실패 → 폴백`).

- [ ] **Step 4: 커밋**

```
git add src/TeampptAddin/UI/TaskPaneHost.cs
git commit -m "feat(read): AI탭에 VectorRecommendService 배선 + 수동 검증"
```

---

## 테스트 실행 절차

A-1a plan과 동일.

## 완료 정의

- 순수 3개(SupabaseAssetMapper/MatchQuery/RecommendationCache) 단위테스트 GREEN.
- HTTP 2개(RemoteAssetCache/VectorRecommendService) 빌드 + 수동 검증.
- PowerPoint에서 텍스트 질의 → Supabase 벡터검색 → 추천 2~3개 표시(읽기경로 깃발).
- 오프라인 시 캐시→번들 폴백으로 빈손 없음.
- **A-1 전체(데이터 토대) 완성:** 인제스트→이해→임베딩→Supabase→벡터추천 읽기경로가 end-to-end로 동작.

## 다음 (A-1 밖)

- A-2 실시간 공유 엔진(현재/전체 슬라이드 온디맨드 + `[슬라이드 N 공유중]`).
- 원격 썸네일 카드 표시·삽입 시 pptx 다운로드(RemoteAssetCache) UI 연결 폴리싱.
- 발표 자산(docs/PITCH.md)에 "완성된 기능: LLM 인제스트→벡터추천" 도식 추가.

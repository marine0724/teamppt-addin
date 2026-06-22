# A-1b/c · Supabase 적재 경로 (인프라 + 임베딩 + 업로드 + 관리자 게이트) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A-1a의 `AssetUnderstanding`를 임베딩(text-embedding-004, 768)하고 Supabase(Postgres `assets` + Storage `pptx`/`thumb`)에 업로드하는 관리자 인제스트 경로를 만든다. 관리자 게이트는 `admin.json` 존재 여부를 `IAccessPolicy` seam 뒤에 두어 유료화(Auth/role) 때 교체만 하면 되게 한다.

**Architecture:** Supabase 프로젝트(테이블/RPC/Storage/RLS)는 인프라 셋업(SQL, 1회 수동). C# 쪽은 순수 로직(`IAccessPolicy` 파일게이트, Postgres row payload 조립)은 xUnit TDD, 외부 HTTP(임베딩, Supabase REST/Storage)는 기존 [GeminiAiService](../../../src/TeampptAddin/Services/GeminiAiService.cs) HttpClient 패턴 어댑터로 두고 수동/통합 검증. 오케스트레이터가 A-1a 이해 → 임베딩 → 업로드를 잇는다.

**Tech Stack:** .NET Framework 4.8, Newtonsoft.Json 13.0.3, System.Net.Http.HttpClient, Supabase(Postgres+pgvector+Storage, REST 직접 호출), Gemini text-embedding-004, xUnit 2.9.

## Global Constraints

- Core/Connect.cs/Globals.cs 직접 수정 금지 (신규 파일만 추가).
- 의존성 추가 금지 — `supabase-csharp` SDK 도입 금지(net48/COM 의존성 충돌). REST를 HttpClient로 직접.
- **시크릿 평문 커밋 금지.** service-role 키·gemini 키는 `%LOCALAPPDATA%\TeampptAddin\admin.json`(gitignore·배포 미포함). anon 키·URL은 `Assets/api-keys.json`(gitignore).
- 임베딩 = `text-embedding-004`, 768차원 고정. 인덱스 = pgvector `hnsw` cosine.
- 관리자 게이트 판단은 **반드시 `IAccessPolicy`를 거친다**(직접 `File.Exists(admin.json)` 호출 산재 금지) — 유료화 시 구현만 교체.
- 단위테스트 절차/MSBuild 경로는 A-1a plan과 동일.

---

## File Structure

| 파일 | 책임 | 테스트 |
|---|---|---|
| (Supabase 콘솔) | `assets` 테이블 + hnsw 인덱스 + `match_assets` RPC + RLS + `pptx`/`thumb` 버킷 | 수동 |
| `Assets/api-keys.json` (수정) | `supabaseUrl`, `supabaseAnonKey` 추가 (기존 `gemini` 유지) | — |
| `Services/AdminCredentials.cs` (신규) | `admin.json` 로딩(serviceKey/geminiKey) (순수 파싱) | `AdminCredentialsTest.cs` |
| `Services/IAccessPolicy.cs` (신규) | `IAccessPolicy` + `LocalFileAccessPolicy`(admin.json 존재=관리자) | `AccessPolicyTest.cs` |
| `Services/EmbeddingService.cs` (신규) | 문자열 → text-embedding-004 → float[768] (HTTP) | 수동 |
| `Services/AssetRowBuilder.cs` (신규) | `AssetUnderstanding`+embedding+경로 → Postgres insert JObject (순수) | `AssetRowBuilderTest.cs` |
| `Services/SupabaseClient.cs` (신규) | row insert(REST) + Storage 파일 업로드 (HTTP) | 수동 |
| `Services/AssetIngestUploader.cs` (신규) | 이해→임베딩→row+Storage 업로드 오케스트레이션 (수동) | 수동 |

순수 3개(AdminCredentials/IAccessPolicy/AssetRowBuilder)가 단위테스트 산출물. HTTP/오케스트레이션 4개는 빌드+수동 검증.

---

### Task 0: Supabase 인프라 셋업 (1회 수동, 사람)

**Files:** (코드 아님 — Supabase 대시보드 SQL Editor)

- [ ] **Step 1: 프로젝트 + pgvector 확장**

Supabase 새 프로젝트 생성 후 SQL Editor에서:
```sql
create extension if not exists vector;
```

- [ ] **Step 2: `assets` 테이블 + 인덱스**

```sql
create table assets (
  id          uuid primary key default gen_random_uuid(),
  file        text not null,          -- Storage 경로: "pptx/표지_01.pptx"
  thumb       text not null,          -- Storage 경로: "thumb/표지_01.png"
  name        text not null,
  category    text not null,          -- 섹션명
  kind        text not null,          -- layout / component
  scope       text,
  tags        text[],
  use_when    text,
  content_fit text[],
  metadata    jsonb,                  -- colors, fonts, slots + 미래 속성(탈출구)
  embed_text  text,                   -- 임베딩 원문(재현/디버깅)
  embedding   vector(768),
  source_deck text,
  created_at  timestamptz default now()
);
create index on assets using hnsw (embedding vector_cosine_ops);
```

- [ ] **Step 3: 벡터 검색 RPC** (A-1d가 호출)

```sql
create or replace function match_assets(query_embedding vector(768), match_count int)
returns table (
  id uuid, file text, thumb text, name text, category text, kind text,
  scope text, tags text[], use_when text, content_fit text[], metadata jsonb,
  similarity float
) language sql stable as $$
  select a.id, a.file, a.thumb, a.name, a.category, a.kind,
         a.scope, a.tags, a.use_when, a.content_fit, a.metadata,
         1 - (a.embedding <=> query_embedding) as similarity
  from assets a
  order by a.embedding <=> query_embedding
  limit match_count;
$$;
```

- [ ] **Step 4: RLS (anon 읽기전용, 쓰기는 service_role만)**

```sql
alter table assets enable row level security;
create policy "anon_read_assets" on assets for select to anon using (true);
-- INSERT/UPDATE/DELETE 정책 없음 → service_role(RLS 우회)만 쓰기 가능
```

- [ ] **Step 5: Storage 버킷**

대시보드 Storage에서 버킷 2개 생성, **public read 허용**(에셋은 비밀 아님; 썸네일 카드·pptx 다운로드용):
- `pptx`
- `thumb`

- [ ] **Step 6: 키/URL 확보**

Project Settings → API에서 **Project URL**, **anon public key**, **service_role key** 확보. (커밋 금지 — 다음 Task에서 로컬 파일에만.)

---

### Task 1: 설정 파일 형식 (api-keys.json + admin.json)

**Files:**
- Modify: `src/TeampptAddin/Assets/api-keys.json` (gitignore 확인)

**Interfaces:**
- Produces: `api-keys.json` = `{ "gemini": "...", "supabaseUrl": "...", "supabaseAnonKey": "..." }`. `admin.json`(관리자 PC 수동 생성, 코드/배포 미포함) = `{ "supabaseServiceKey": "...", "geminiKey": "..." }`.

- [ ] **Step 1: gitignore 확인**

`api-keys.json`과 `admin.json`이 무시되는지 확인:
```
git check-ignore src/TeampptAddin/Assets/api-keys.json
```
Expected: 경로가 출력됨(=무시됨). 안 되면 `.gitignore`에 `api-keys.json`, `admin.json` 추가 후 커밋.

- [ ] **Step 2: api-keys.json에 supabase 필드 추가**

`src/TeampptAddin/Assets/api-keys.json`을 열어 **기존 `gemini` 값은 그대로 두고**(이미 작동 확인된 키 — 절대 교체/수정 금지) 두 필드만 추가:
```json
{
  "gemini": "<기존 값 유지>",
  "supabaseUrl": "https://<project>.supabase.co",
  "supabaseAnonKey": "<anon public key>"
}
```

- [ ] **Step 3: admin.json 수동 생성 (관리자 PC)**

`%LOCALAPPDATA%\TeampptAddin\admin.json` 생성(이 PC가 관리자임을 의미):
```json
{
  "supabaseServiceKey": "<service_role key>",
  "geminiKey": "<api-keys.json의 gemini와 동일 값>"
}
```
(커밋 안 함. 일반 사용자 PC엔 이 파일이 없음 → 인제스트 버튼 숨김.)

- [ ] **Step 4: 커밋 (형식 문서만, 값 제외)**

```
git add .gitignore
git commit -m "chore(ingest): api-keys/admin.json gitignore 확인 (supabase 필드)"
```
(값이 든 파일은 커밋되지 않음 — gitignore.)

---

### Task 2: 관리자 자격증명 로더

**Files:**
- Create: `src/TeampptAddin/Services/AdminCredentials.cs`
- Test: `src/TeampptAddin.Tests/AdminCredentialsTest.cs`

**Interfaces:**
- Produces: `class AdminCredentials { string SupabaseServiceKey; string GeminiKey; static AdminCredentials Load(string path); static string DefaultPath }`. `Load`는 JSON 파싱. `DefaultPath` = `%LOCALAPPDATA%\TeampptAddin\admin.json`.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using System.IO;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AdminCredentialsTest
    {
        [Fact]
        public void Load_Reads_Service_And_Gemini_Keys()
        {
            var tmp = Path.GetTempFileName();
            File.WriteAllText(tmp, @"{""supabaseServiceKey"":""svc123"",""geminiKey"":""AIzaXYZ""}");
            try
            {
                var c = AdminCredentials.Load(tmp);
                Assert.Equal("svc123", c.SupabaseServiceKey);
                Assert.Equal("AIzaXYZ", c.GeminiKey);
            }
            finally { File.Delete(tmp); }
        }
    }
}
```

- [ ] **Step 2: 실패 확인** — 테스트 실행. Expected: FAIL (`AdminCredentials` 미정의).

- [ ] **Step 3: 최소 구현**

```csharp
using System;
using System.IO;
using System.Text;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>관리자 자격증명(admin.json) 로더. 이 파일 존재 = 관리자 PC.</summary>
    public class AdminCredentials
    {
        public string SupabaseServiceKey { get; set; }
        public string GeminiKey { get; set; }

        public static string DefaultPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TeampptAddin", "admin.json");

        public static AdminCredentials Load(string path)
        {
            var o = JObject.Parse(File.ReadAllText(path, Encoding.UTF8));
            return new AdminCredentials
            {
                SupabaseServiceKey = o["supabaseServiceKey"]?.ToString(),
                GeminiKey = o["geminiKey"]?.ToString()
            };
        }
    }
}
```

- [ ] **Step 4: 통과 확인** — 테스트 실행. Expected: 1 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/AdminCredentials.cs src/TeampptAddin.Tests/AdminCredentialsTest.cs
git commit -m "feat(ingest): AdminCredentials (admin.json 로더)"
```

---

### Task 3: 접근 권한 seam (IAccessPolicy)

**Files:**
- Create: `src/TeampptAddin/Services/IAccessPolicy.cs`
- Test: `src/TeampptAddin.Tests/AccessPolicyTest.cs`

**Interfaces:**
- Consumes: `AdminCredentials.DefaultPath`.
- Produces: `interface IAccessPolicy { bool IsAdmin { get; } bool CanIngest { get; } }` + `class LocalFileAccessPolicy : IAccessPolicy`(생성자에 admin.json 경로 주입, 기본 = `AdminCredentials.DefaultPath`). Stage 0 구현 = 파일 존재 여부. 유료화 시 `AuthAccessPolicy`로 교체.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using System.IO;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AccessPolicyTest
    {
        [Fact]
        public void IsAdmin_True_When_File_Exists()
        {
            var tmp = Path.GetTempFileName();
            try
            {
                var p = new LocalFileAccessPolicy(tmp);
                Assert.True(p.IsAdmin);
                Assert.True(p.CanIngest);
            }
            finally { File.Delete(tmp); }
        }

        [Fact]
        public void IsAdmin_False_When_File_Missing()
        {
            var p = new LocalFileAccessPolicy(Path.Combine(Path.GetTempPath(), "no_such_admin_" + System.Guid.NewGuid() + ".json"));
            Assert.False(p.IsAdmin);
            Assert.False(p.CanIngest);
        }
    }
}
```

- [ ] **Step 2: 실패 확인** — 테스트 실행. Expected: FAIL.

- [ ] **Step 3: 최소 구현**

```csharp
using System.IO;

namespace TeampptAddin
{
    /// <summary>권한 판단 seam. Stage 0 = admin.json 존재. 유료화(Stage 1) = 로그인 role로 교체.</summary>
    public interface IAccessPolicy
    {
        bool IsAdmin { get; }
        bool CanIngest { get; }
    }

    public class LocalFileAccessPolicy : IAccessPolicy
    {
        private readonly string _adminPath;
        public LocalFileAccessPolicy(string adminPath = null)
        {
            _adminPath = adminPath ?? AdminCredentials.DefaultPath;
        }
        public bool IsAdmin => File.Exists(_adminPath);
        public bool CanIngest => IsAdmin;
    }
}
```

- [ ] **Step 4: 통과 확인** — 테스트 실행. Expected: 2 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/IAccessPolicy.cs src/TeampptAddin.Tests/AccessPolicyTest.cs
git commit -m "feat(ingest): IAccessPolicy seam (admin.json 게이트, 유료화 교체지점)"
```

---

### Task 4: Postgres row 조립기

**Files:**
- Create: `src/TeampptAddin/Services/AssetRowBuilder.cs`
- Test: `src/TeampptAddin.Tests/AssetRowBuilderTest.cs`

**Interfaces:**
- Consumes: `AssetUnderstanding`, `HeaderAsset`.
- Produces: `static JObject AssetRowBuilder.Build(AssetUnderstanding u, float[] embedding, string embedText, string filePath, string thumbPath, string sourceDeck)`. `assets` 테이블 컬럼에 매핑. `colors/fonts/slots`는 `metadata` jsonb로. `embedding`은 pgvector 텍스트 형식 문자열 `"[v1,v2,...]"`로. `tags/content_fit`는 JSON 배열.

- [ ] **Step 1: 실패 테스트 작성**

```csharp
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetRowBuilderTest
    {
        private static AssetUnderstanding U() => new AssetUnderstanding
        {
            Asset = new HeaderAsset
            {
                Name = "연도강조 표지", Kind = "layout", Category = "표지", Scope = "slide",
                UseWhen = "연도 강조", Tags = new List<string> { "표지" },
                ContentFit = new List<string> { "표지" },
                Colors = new List<AssetColor> { new AssetColor { Role = "main", Value = "#1A2B4C", Locked = false } },
                Slots = new List<AssetSlot> { new AssetSlot { Name = "title", Type = "text", PerSlide = true } }
            },
            ExampleIntents = new List<string> { "IR 표지" }
        };

        [Fact]
        public void Build_Maps_Columns_And_Vector_String()
        {
            var row = AssetRowBuilder.Build(U(), new float[] { 0.1f, 0.2f }, "embed text",
                "pptx/표지_01.pptx", "thumb/표지_01.png", "bundle.pptx");

            Assert.Equal("연도강조 표지", row["name"]);
            Assert.Equal("layout", row["kind"]);
            Assert.Equal("표지", row["category"]);
            Assert.Equal("pptx/표지_01.pptx", row["file"]);
            Assert.Equal("thumb/표지_01.png", row["thumb"]);
            Assert.Equal("[0.1,0.2]", row["embedding"]);          // pgvector 텍스트 형식
            Assert.Equal("embed text", row["embed_text"]);
        }

        [Fact]
        public void Build_Puts_Structure_In_Metadata()
        {
            var row = AssetRowBuilder.Build(U(), new float[] { 0.1f }, "t", "p", "th", "d");
            var meta = (JObject)row["metadata"];
            Assert.NotNull(meta["colors"]);
            Assert.NotNull(meta["slots"]);
        }
    }
}
```

- [ ] **Step 2: 실패 확인** — 테스트 실행. Expected: FAIL.

- [ ] **Step 3: 최소 구현**

```csharp
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>AssetUnderstanding + 임베딩 → assets 테이블 insert JObject. 구조데이터는 metadata jsonb.</summary>
    public static class AssetRowBuilder
    {
        public static JObject Build(AssetUnderstanding u, float[] embedding, string embedText,
            string filePath, string thumbPath, string sourceDeck)
        {
            var a = u.Asset;
            var metadata = new JObject
            {
                ["colors"] = JArray.FromObject(a.Colors ?? new System.Collections.Generic.List<AssetColor>()),
                ["fonts"] = JArray.FromObject(a.Fonts ?? new System.Collections.Generic.List<AssetFont>()),
                ["slots"] = JArray.FromObject(a.Slots ?? new System.Collections.Generic.List<AssetSlot>())
            };

            var vec = "[" + string.Join(",", embedding.Select(v => v.ToString(CultureInfo.InvariantCulture))) + "]";

            return new JObject
            {
                ["file"] = filePath,
                ["thumb"] = thumbPath,
                ["name"] = a.Name,
                ["category"] = a.Category,
                ["kind"] = a.Kind,
                ["scope"] = a.Scope,
                ["tags"] = JArray.FromObject(a.Tags ?? new System.Collections.Generic.List<string>()),
                ["use_when"] = a.UseWhen,
                ["content_fit"] = JArray.FromObject(a.ContentFit ?? new System.Collections.Generic.List<string>()),
                ["metadata"] = metadata,
                ["embed_text"] = embedText,
                ["embedding"] = vec,
                ["source_deck"] = sourceDeck
            };
        }
    }
}
```

- [ ] **Step 4: 통과 확인** — 테스트 실행. Expected: 2 PASS.

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/AssetRowBuilder.cs src/TeampptAddin.Tests/AssetRowBuilderTest.cs
git commit -m "feat(ingest): AssetRowBuilder (Postgres row + pgvector 문자열)"
```

---

### Task 5: 임베딩 서비스 (HTTP 어댑터, 수동)

**Files:**
- Create: `src/TeampptAddin/Services/EmbeddingService.cs`

**Interfaces:**
- Produces: `class EmbeddingService { EmbeddingService(string apiKey); Task<float[]> EmbedAsync(string text); }`. `text-embedding-004:embedContent` 호출 → `embedding.values`(768) 파싱. 재시도(503/429/500) 동일 패턴.

- [ ] **Step 1: 구현 작성**

```csharp
using System;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>Gemini text-embedding-004 → 768차원 벡터. 검색 임베딩과 인제스트 임베딩 공용.</summary>
    public class EmbeddingService
    {
        private static readonly HttpClient Http = new HttpClient();
        private readonly string _apiKey;
        public EmbeddingService(string apiKey) { _apiKey = apiKey; }

        public async Task<float[]> EmbedAsync(string text)
        {
            var body = new JObject
            {
                ["model"] = "models/text-embedding-004",
                ["content"] = new JObject
                {
                    ["parts"] = new JArray { new JObject { ["text"] = text } }
                }
            }.ToString(Formatting.None);

            var url = $"https://generativelanguage.googleapis.com/v1beta/models/text-embedding-004:embedContent?key={_apiKey}";

            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var content = new StringContent(body, Encoding.UTF8, "application/json");
                Http.DefaultRequestHeaders.Authorization = null;
                var resp = await Http.PostAsync(url, content).ConfigureAwait(false);
                var respBody = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                Logger.Log($"[Embed] attempt {attempt}: HTTP {(int)resp.StatusCode}");

                if (resp.IsSuccessStatusCode)
                {
                    var values = JObject.Parse(respBody)["embedding"]?["values"] as JArray;
                    if (values == null) throw new InvalidOperationException("임베딩 응답에 values 없음.");
                    return values.Select(v => v.Value<float>()).ToArray();
                }

                var status = (int)resp.StatusCode;
                bool transient = status == 503 || status == 429 || status == 500;
                if (transient && attempt < maxAttempts)
                {
                    await Task.Delay(500 * (1 << (attempt - 1))).ConfigureAwait(false);
                    continue;
                }
                throw new HttpRequestException($"임베딩 API 오류 ({status}): {respBody}");
            }
            throw new InvalidOperationException("임베딩 재시도 소진.");
        }
    }
}
```

- [ ] **Step 2: 빌드 확인** — Run: A-1a의 MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Services/EmbeddingService.cs
git commit -m "feat(ingest): EmbeddingService (text-embedding-004 768)"
```

---

### Task 6: Supabase REST/Storage 클라이언트 (HTTP 어댑터, 수동)

**Files:**
- Create: `src/TeampptAddin/Services/SupabaseClient.cs`

**Interfaces:**
- Produces: `class SupabaseClient { SupabaseClient(string baseUrl, string key); Task InsertAssetAsync(JObject row); Task UploadObjectAsync(string bucket, string path, byte[] bytes, string contentType); Task<string> RpcAsync(string fn, JObject args); }`. `key`는 쓰기=service-role, 읽기=anon. 헤더 `apikey` + `Authorization: Bearer {key}`. Insert는 `Prefer: return=minimal`.

- [ ] **Step 1: 구현 작성**

```csharp
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    /// <summary>Supabase REST/RPC/Storage 직접 호출(SDK 없이). key: 쓰기=service-role, 읽기=anon.</summary>
    public class SupabaseClient
    {
        private static readonly HttpClient Http = new HttpClient();
        private readonly string _baseUrl;
        private readonly string _key;

        public SupabaseClient(string baseUrl, string key)
        {
            _baseUrl = baseUrl.TrimEnd('/');
            _key = key;
        }

        private void ApplyHeaders(HttpRequestMessage req)
        {
            req.Headers.TryAddWithoutValidation("apikey", _key);
            req.Headers.TryAddWithoutValidation("Authorization", "Bearer " + _key);
        }

        public async Task InsertAssetAsync(JObject row)
        {
            var req = new HttpRequestMessage(HttpMethod.Post, $"{_baseUrl}/rest/v1/assets");
            ApplyHeaders(req);
            req.Headers.TryAddWithoutValidation("Prefer", "return=minimal");
            req.Content = new StringContent(row.ToString(Formatting.None), Encoding.UTF8, "application/json");
            var resp = await Http.SendAsync(req).ConfigureAwait(false);
            var b = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            Logger.Log($"[Supabase] insert assets: HTTP {(int)resp.StatusCode}");
            if (!resp.IsSuccessStatusCode)
                throw new HttpRequestException($"Supabase insert 오류 ({(int)resp.StatusCode}): {b}");
        }

        public async Task UploadObjectAsync(string bucket, string path, byte[] bytes, string contentType)
        {
            var req = new HttpRequestMessage(HttpMethod.Post, $"{_baseUrl}/storage/v1/object/{bucket}/{Uri.EscapeDataString(path)}");
            ApplyHeaders(req);
            req.Headers.TryAddWithoutValidation("x-upsert", "true");
            req.Content = new ByteArrayContent(bytes);
            req.Content.Headers.TryAddWithoutValidation("Content-Type", contentType);
            var resp = await Http.SendAsync(req).ConfigureAwait(false);
            var b = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            Logger.Log($"[Supabase] upload {bucket}/{path}: HTTP {(int)resp.StatusCode}");
            if (!resp.IsSuccessStatusCode)
                throw new HttpRequestException($"Supabase storage 오류 ({(int)resp.StatusCode}): {b}");
        }

        public async Task<string> RpcAsync(string fn, JObject args)
        {
            var req = new HttpRequestMessage(HttpMethod.Post, $"{_baseUrl}/rest/v1/rpc/{fn}");
            ApplyHeaders(req);
            req.Content = new StringContent(args.ToString(Formatting.None), Encoding.UTF8, "application/json");
            var resp = await Http.SendAsync(req).ConfigureAwait(false);
            var b = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            Logger.Log($"[Supabase] rpc {fn}: HTTP {(int)resp.StatusCode}");
            if (!resp.IsSuccessStatusCode)
                throw new HttpRequestException($"Supabase rpc 오류 ({(int)resp.StatusCode}): {b}");
            return b;
        }
    }
}
```

- [ ] **Step 2: 빌드 확인** — Run: A-1a의 MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 커밋**

```
git add src/TeampptAddin/Services/SupabaseClient.cs
git commit -m "feat(ingest): SupabaseClient (REST insert/RPC/Storage upload)"
```

---

### Task 7: 인제스트 업로더 오케스트레이터 + 수동 검증 (마이그레이션 깃발)

**Files:**
- Create: `src/TeampptAddin/Services/AssetIngestUploader.cs`

**Interfaces:**
- Consumes: `IAccessPolicy`, `AdminCredentials`, `AssetUnderstandingService`, `EmbeddingService`, `EmbedTextBuilder`, `AssetRowBuilder`, `SupabaseClient`, `IngestRunner`(로컬 split 산출물), `Globals.AssetsDir`.
- Produces: `class AssetIngestUploader { Task<int> UploadDirectoryAsync(string splitDir, string sourceDeck); }`. `splitDir`(IngestRunner 산출 = `{AssetId}.pptx`/`.png` 쌍들)을 순회하며 각 에셋을 이해→임베딩→Storage 업로드(pptx,thumb)→row insert. 관리자 게이트 통과 못 하면 즉시 0 반환. 처리 개수 반환.

- [ ] **Step 1: 구현 작성**

```csharp
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace TeampptAddin
{
    /// <summary>
    /// 로컬 split 산출물(폴더 안 {id}.pptx/{id}.png 쌍)을 이해→임베딩→Supabase 업로드.
    /// 관리자 게이트(IAccessPolicy) 통과해야 동작. 기존 7+35 에셋 마이그레이션도 이 경로 1회 실행.
    /// </summary>
    public class AssetIngestUploader
    {
        private readonly IAccessPolicy _policy;
        public AssetIngestUploader(IAccessPolicy policy = null)
        {
            _policy = policy ?? new LocalFileAccessPolicy();
        }

        public async Task<int> UploadDirectoryAsync(string splitDir, string sourceDeck)
        {
            if (!_policy.CanIngest)
            {
                Logger.Log("[Upload] 관리자 아님 — 인제스트 거부");
                return 0;
            }

            var cred = AdminCredentials.Load(AdminCredentials.DefaultPath);
            var keys = JsonKeys();   // api-keys.json의 supabaseUrl
            var understand = new AssetUnderstandingService(cred.GeminiKey);
            var embed = new EmbeddingService(cred.GeminiKey);
            var supa = new SupabaseClient(keys.url, cred.SupabaseServiceKey);

            int count = 0;
            foreach (var pngPath in Directory.GetFiles(splitDir, "*.png"))
            {
                var id = Path.GetFileNameWithoutExtension(pngPath);
                var pptxPath = Path.Combine(splitDir, id + ".pptx");
                if (!File.Exists(pptxPath)) continue;

                var category = id.Contains("_") ? id.Substring(0, id.LastIndexOf('_')) : id;

                var u = await understand.UnderstandAsync(pngPath, category, "pptx/" + id + ".pptx").ConfigureAwait(false);
                var embedText = EmbedTextBuilder.Build(u);
                var vector = await embed.EmbedAsync(embedText).ConfigureAwait(false);

                await supa.UploadObjectAsync("thumb", id + ".png", File.ReadAllBytes(pngPath), "image/png").ConfigureAwait(false);
                await supa.UploadObjectAsync("pptx", id + ".pptx", File.ReadAllBytes(pptxPath),
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation").ConfigureAwait(false);

                var row = AssetRowBuilder.Build(u, vector, embedText, "pptx/" + id + ".pptx", "thumb/" + id + ".png", sourceDeck);
                await supa.InsertAssetAsync(row).ConfigureAwait(false);

                Logger.Log($"[Upload] {id} → Supabase OK (kind={u.Asset.Kind})");
                count++;
            }
            Logger.Log($"[Upload] 완료: {count}개");
            return count;
        }

        private (string url, string anon) JsonKeys()
        {
            var path = Path.Combine(Globals.AssetsDir, "api-keys.json");
            var o = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(path));
            return (o["supabaseUrl"]?.ToString(), o["supabaseAnonKey"]?.ToString());
        }
    }
}
```

- [ ] **Step 2: 본프로젝트 빌드 확인** — Run: A-1a MSBuild 명령. Expected: Build succeeded.

- [ ] **Step 3: 전체 단위테스트 GREEN** — "테스트 실행 절차". Expected: 기존 + 신규(AdminCredentials 1, AccessPolicy 2, AssetRowBuilder 2) 모두 PASS.

- [ ] **Step 4: 수동 검증 + 기존 에셋 마이그레이션 (깃발)**

전제: Task 0~1 완료(Supabase 인프라 + admin.json + api-keys.json). 로컬 인제스트 코어로 35개 split 산출물이 있는 폴더(예: `%LOCALAPPDATA%\TeampptAddin\ingest-test`)와 로컬 7개도 split(또는 직접 pptx/png 쌍 준비). 임시 실행 지점에서:
```csharp
var n = new AssetIngestUploader().UploadDirectoryAsync(@"%LOCALAPPDATA%\TeampptAddin\ingest-test", "experiment-bundle.pptx").GetAwaiter().GetResult();
Logger.Log($"[Upload] 업로드 개수={n}");
```
확인:
- Supabase `assets` 테이블에 행이 생성됨(Table Editor), `kind`가 layout/component, `embedding` 채워짐.
- Storage `thumb`/`pptx` 버킷에 파일 업로드됨.
- 일반 사용자 모드(admin.json 없는 상태)에선 `UploadDirectoryAsync`가 0 반환(게이트 동작).

- [ ] **Step 5: 커밋**

```
git add src/TeampptAddin/Services/AssetIngestUploader.cs
git commit -m "feat(ingest): AssetIngestUploader (이해→임베딩→Supabase 업로드) + 마이그레이션"
```

---

## 테스트 실행 절차

A-1a plan과 동일.

## 완료 정의

- Supabase 인프라(테이블/hnsw/RPC/RLS/버킷) 셋업 완료.
- 순수 3개(AdminCredentials/IAccessPolicy/AssetRowBuilder) 단위테스트 GREEN.
- HTTP 3개(EmbeddingService/SupabaseClient/AssetIngestUploader) 빌드 + 수동 업로드 검증.
- 기존 7+35 에셋이 Supabase에 행+벡터+Storage로 적재됨(마이그레이션).
- 관리자 게이트(IAccessPolicy) 동작: admin.json 없으면 인제스트 거부.

## 다음 plan (이 plan 밖)

- **A-1d:** anon 키로 `match_assets` RPC 벡터검색 → top-N → AI탭 텍스트 질의 추천 연결 + 오프라인 캐시/번들 폴백.

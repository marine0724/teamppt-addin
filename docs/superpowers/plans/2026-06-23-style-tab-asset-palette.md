# 스타일 탭 에셋 기반 팔레트 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 삽입한 에셋의 실제 색(`HeaderAsset.Colors`)에서 원본·반전 팔레트를 자동 생성해 스타일 탭에 띄우고, 클릭 시 활성 슬라이드 1장의 배경·도형·텍스트·폰트까지 그 컨셉으로 바꾼다.

**Architecture:** 색상 로직은 LLM 없이 순수 C# 색상수학(HSL 변환 + WCAG 대비). 세 개의 순수 함수 컴포넌트(`ColorHsl` → `PaletteRoleMapper` → `PaletteGenerator`)가 에셋 색을 정규화하고 팔레트를 생성한다. `SlideStyleApplier`(PPT Interop)가 슬라이드에 적용하고, `AssetPanel` UI가 삽입 시 앵커를 잡아 스타일 탭을 갱신한다.

**Tech Stack:** C# (.NET Framework 4.8), Newtonsoft.Json, Microsoft.Office.Interop.PowerPoint, xUnit (테스트), WPF (스타일 탭 UI).

## Global Constraints

- **타깃 프레임워크: net48.** 메인 프로젝트는 COM 등록(`RegisterForComInterop=true`)을 트리거하므로 테스트는 **항상 `-p:RegisterForComInterop=false`** 로 빌드한다.
- **테스트/빌드 전 PowerPoint를 종료한다.** 열려 있으면 DLL 잠금으로 빌드 실패.
- **테스트 러너(확정):** `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter <ClassName>`
- **메인 빌드(Interop/UI 검증용, 관리자 권한):** CLAUDE.md의 MSBuild 명령 사용 (`Start-Process ... -Verb RunAs`), 결과는 `build.log` 끝 5줄 확인.
- **JSON 직렬화는 Newtonsoft.Json**(`[JsonProperty]`). 표준 라이브러리 외 신규 의존성 추가 금지.
- **CoordinateConverter에 폴백 로직 추가 금지.** Core/Connect.cs/Globals.cs 구조 존중.
- **hex 색은 항상 `#RRGGBB` 6자리 대문자 문자열**로 표준화한다.
- 설계 문서: `docs/superpowers/specs/2026-06-23-style-tab-asset-palette-design.md`.

---

### Task 1: ColorHsl 색상수학 유틸

**Files:**
- Create: `src/TeampptAddin/Services/ColorHsl.cs`
- Test: `src/TeampptAddin.Tests/ColorHslTest.cs`

**Interfaces:**
- Consumes: 없음 (순수 함수, 입력은 hex 문자열).
- Produces:
  - `struct Hsl { double H; double S; double L; }` — H는 0~360, S/L은 0~1.
  - `static Hsl ColorHsl.FromHex(string hex)`
  - `static string ColorHsl.ToHex(Hsl hsl)` — `#RRGGBB` 대문자 반환
  - `static Hsl ColorHsl.WithLightness(Hsl hsl, double l)`
  - `static double ColorHsl.ContrastRatio(string hexA, string hexB)` — WCAG 1.0~21.0
  - `static string ColorHsl.AdjustForContrast(string fgHex, string bgHex, double targetRatio)` — H·S 유지, L만 조정해 대비 충족하는 fg hex 반환

- [ ] **Step 1: 테스트 작성**

Create `src/TeampptAddin.Tests/ColorHslTest.cs`:

```csharp
using TeampptAddin;
using Xunit;

namespace TeampptAddin.Tests
{
    public class ColorHslTest
    {
        [Theory]
        [InlineData("#2563EB")]
        [InlineData("#1E293B")]
        [InlineData("#FFFFFF")]
        [InlineData("#000000")]
        public void FromHex_ToHex_RoundTrips(string hex)
        {
            var hsl = ColorHsl.FromHex(hex);
            Assert.Equal(hex, ColorHsl.ToHex(hsl));
        }

        [Fact]
        public void FromHex_Accepts_Lowercase_And_NoHash()
        {
            Assert.Equal("#2563EB", ColorHsl.ToHex(ColorHsl.FromHex("2563eb")));
        }

        [Fact]
        public void ContrastRatio_BlackWhite_IsMax()
        {
            Assert.True(ColorHsl.ContrastRatio("#000000", "#FFFFFF") > 20.9);
        }

        [Fact]
        public void ContrastRatio_SameColor_IsOne()
        {
            Assert.Equal(1.0, ColorHsl.ContrastRatio("#2563EB", "#2563EB"), 2);
        }

        [Fact]
        public void WithLightness_Sets_L_Keeps_Hue()
        {
            var hsl = ColorHsl.FromHex("#2563EB");
            var lighter = ColorHsl.WithLightness(hsl, 0.9);
            Assert.Equal(0.9, lighter.L, 3);
            Assert.Equal(hsl.H, lighter.H, 1);
        }

        [Fact]
        public void AdjustForContrast_Darkens_On_Light_Background()
        {
            // 회색 글씨를 흰 배경에서 대비 4.5 이상으로
            var adjusted = ColorHsl.AdjustForContrast("#999999", "#FFFFFF", 4.5);
            Assert.True(ColorHsl.ContrastRatio(adjusted, "#FFFFFF") >= 4.5);
        }

        [Fact]
        public void AdjustForContrast_Lightens_On_Dark_Background()
        {
            var adjusted = ColorHsl.AdjustForContrast("#444444", "#0A1428", 4.5);
            Assert.True(ColorHsl.ContrastRatio(adjusted, "#0A1428") >= 4.5);
        }

        [Fact]
        public void AdjustForContrast_Returns_Original_When_Already_Sufficient()
        {
            var adjusted = ColorHsl.AdjustForContrast("#000000", "#FFFFFF", 4.5);
            Assert.Equal("#000000", adjusted);
        }
    }
}
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter ColorHslTest`
Expected: FAIL — `ColorHsl` 타입 없음 (컴파일 에러).

- [ ] **Step 3: 구현 작성**

Create `src/TeampptAddin/Services/ColorHsl.cs`:

```csharp
using System;

namespace TeampptAddin
{
    public struct Hsl
    {
        public double H; // 0..360
        public double S; // 0..1
        public double L; // 0..1
        public Hsl(double h, double s, double l) { H = h; S = s; L = l; }
    }

    public static class ColorHsl
    {
        public static Hsl FromHex(string hex)
        {
            var (r, g, b) = HexToRgb(hex);
            double rd = r / 255.0, gd = g / 255.0, bd = b / 255.0;
            double max = Math.Max(rd, Math.Max(gd, bd));
            double min = Math.Min(rd, Math.Min(gd, bd));
            double h = 0, s, l = (max + min) / 2.0;
            double d = max - min;
            if (d == 0) { s = 0; }
            else
            {
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
                if (max == rd) h = (gd - bd) / d + (gd < bd ? 6 : 0);
                else if (max == gd) h = (bd - rd) / d + 2;
                else h = (rd - gd) / d + 4;
                h *= 60;
            }
            return new Hsl(h, s, l);
        }

        public static string ToHex(Hsl hsl)
        {
            double h = hsl.H, s = Clamp01(hsl.S), l = Clamp01(hsl.L);
            double r, g, b;
            if (s == 0) { r = g = b = l; }
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                double p = 2 * l - q;
                double hk = h / 360.0;
                r = HueToRgb(p, q, hk + 1.0 / 3.0);
                g = HueToRgb(p, q, hk);
                b = HueToRgb(p, q, hk - 1.0 / 3.0);
            }
            return "#" + ToByteHex(r) + ToByteHex(g) + ToByteHex(b);
        }

        public static Hsl WithLightness(Hsl hsl, double l)
        {
            return new Hsl(hsl.H, hsl.S, Clamp01(l));
        }

        public static double ContrastRatio(string hexA, string hexB)
        {
            double la = RelativeLuminance(hexA);
            double lb = RelativeLuminance(hexB);
            double lighter = Math.Max(la, lb);
            double darker = Math.Min(la, lb);
            return (lighter + 0.05) / (darker + 0.05);
        }

        public static string AdjustForContrast(string fgHex, string bgHex, double targetRatio)
        {
            if (ContrastRatio(fgHex, bgHex) >= targetRatio) return fgHex;
            var fg = FromHex(fgHex);
            bool bgIsLight = RelativeLuminance(bgHex) > 0.5;
            // 밝은 배경 → fg를 어둡게(L↓), 어두운 배경 → fg를 밝게(L↑)
            for (int i = 0; i <= 50; i++)
            {
                double l = bgIsLight ? fg.L - i * 0.02 : fg.L + i * 0.02;
                if (l < 0 || l > 1) break;
                var candidate = ToHex(WithLightness(fg, l));
                if (ContrastRatio(candidate, bgHex) >= targetRatio) return candidate;
            }
            // 끝까지 못 맞추면 흑/백 폴백
            return bgIsLight ? "#000000" : "#FFFFFF";
        }

        // ── helpers ────────────────────────────────────────────────
        private static double RelativeLuminance(string hex)
        {
            var (r, g, b) = HexToRgb(hex);
            double rs = LinearChannel(r / 255.0);
            double gs = LinearChannel(g / 255.0);
            double bs = LinearChannel(b / 255.0);
            return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;
        }

        private static double LinearChannel(double c)
        {
            return c <= 0.03928 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6.0) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2.0) return q;
            if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6;
            return p;
        }

        private static (int, int, int) HexToRgb(string hex)
        {
            hex = (hex ?? "").Trim().TrimStart('#');
            if (hex.Length != 6) throw new ArgumentException($"잘못된 hex: {hex}");
            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return (r, g, b);
        }

        private static string ToByteHex(double channel01)
        {
            int v = (int)Math.Round(Clamp01(channel01) * 255.0);
            return v.ToString("X2");
        }

        private static double Clamp01(double v) => v < 0 ? 0 : (v > 1 ? 1 : v);
    }
}
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter ColorHslTest`
Expected: PASS (8개 케이스).

- [ ] **Step 5: 커밋**

```bash
git add src/TeampptAddin/Services/ColorHsl.cs src/TeampptAddin.Tests/ColorHslTest.cs
git commit -m "feat(style): ColorHsl 색상수학 유틸(HSL 변환·WCAG 대비·L조정)"
```

---

### Task 2: PaletteColors.Background + NormalizedPalette 모델

**Files:**
- Modify: `src/TeampptAddin/Models/StylePalette.cs:6-19` (`PaletteColors`에 `Background` 추가)
- Create: `src/TeampptAddin/Models/NormalizedPalette.cs`
- Test: `src/TeampptAddin.Tests/NormalizedPaletteTest.cs`

**Interfaces:**
- Consumes: 없음.
- Produces:
  - `PaletteColors.Background` (string, nullable, `[JsonProperty("background")]`).
  - `class NormalizedPalette { string Background; string Main; string Sub1; string Sub2; string Text; }` — 모두 `#RRGGBB` hex.

- [ ] **Step 1: 테스트 작성**

Create `src/TeampptAddin.Tests/NormalizedPaletteTest.cs`:

```csharp
using Newtonsoft.Json;
using TeampptAddin;
using Xunit;

namespace TeampptAddin.Tests
{
    public class NormalizedPaletteTest
    {
        [Fact]
        public void NormalizedPalette_Holds_Five_Roles()
        {
            var np = new NormalizedPalette
            {
                Background = "#FFFFFF", Main = "#2563EB",
                Sub1 = "#3B82F6", Sub2 = "#93C5FD", Text = "#1E293B"
            };
            Assert.Equal("#FFFFFF", np.Background);
            Assert.Equal("#1E293B", np.Text);
        }

        [Fact]
        public void PaletteColors_Background_Serializes_As_background()
        {
            var json = JsonConvert.SerializeObject(new PaletteColors { Background = "#0A1428" });
            Assert.Contains("\"background\":\"#0A1428\"", json);
        }

        [Fact]
        public void PaletteColors_Without_Background_Deserializes_Null()
        {
            var pc = JsonConvert.DeserializeObject<PaletteColors>("{\"main\":\"#2563EB\"}");
            Assert.Null(pc.Background);
            Assert.Equal("#2563EB", pc.Main);
        }
    }
}
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter NormalizedPaletteTest`
Expected: FAIL — `NormalizedPalette` 없음, `PaletteColors.Background` 없음.

- [ ] **Step 3: 구현 작성**

`src/TeampptAddin/Models/StylePalette.cs`의 `PaletteColors`에 `Background`를 **맨 앞에** 추가:

```csharp
    public class PaletteColors
    {
        [JsonProperty("background")]
        public string Background { get; set; }

        [JsonProperty("main")]
        public string Main { get; set; }

        [JsonProperty("sub1")]
        public string Sub1 { get; set; }

        [JsonProperty("sub2")]
        public string Sub2 { get; set; }

        [JsonProperty("text")]
        public string Text { get; set; }
    }
```

Create `src/TeampptAddin/Models/NormalizedPalette.cs`:

```csharp
namespace TeampptAddin
{
    /// 에셋 색을 5개 역할로 정규화한 결과. 모든 값은 #RRGGBB hex.
    public class NormalizedPalette
    {
        public string Background { get; set; }
        public string Main { get; set; }
        public string Sub1 { get; set; }
        public string Sub2 { get; set; }
        public string Text { get; set; }
    }
}
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter NormalizedPaletteTest`
Expected: PASS (3개).

- [ ] **Step 5: 커밋**

```bash
git add src/TeampptAddin/Models/StylePalette.cs src/TeampptAddin/Models/NormalizedPalette.cs src/TeampptAddin.Tests/NormalizedPaletteTest.cs
git commit -m "feat(style): PaletteColors.Background + NormalizedPalette 모델"
```

---

### Task 3: PaletteRoleMapper — 에셋 색 → 정규 5역할

**Files:**
- Create: `src/TeampptAddin/Services/PaletteRoleMapper.cs`
- Test: `src/TeampptAddin.Tests/PaletteRoleMapperTest.cs`

**Interfaces:**
- Consumes: `ColorHsl` (Task 1), `NormalizedPalette` (Task 2), `AssetColor`(`Role`,`Value`,`Locked` — 기존).
- Produces: `static NormalizedPalette PaletteRoleMapper.Map(System.Collections.Generic.List<AssetColor> colors)` — colors가 null/빈 배열이면 `null` 반환(호출부가 스킵). 빈 역할은 휴리스틱으로 보강.

규칙:
- **main**: role에 `main`/`primary` 포함 → 그 색. 없으면 채도(S) 최대 색. 그것도 없으면 첫 색.
- **text**: role에 `text`/`foreground` 포함 → 그 색. 없으면 명도(L) 최소(가장 어두운) 색. 없으면 main의 L=0.15.
- **background**: role에 `background`/`bg` 포함 → 그 색. 없으면 L≥0.85인 색. 없으면 main의 L=0.96.
- **sub1**: role에 `sub`/`accent`/`secondary` 포함된 색 중 첫째. 없으면 main의 L+0.18.
- **sub2**: 위 후보의 둘째. 없으면 main의 L−0.18.

- [ ] **Step 1: 테스트 작성**

Create `src/TeampptAddin.Tests/PaletteRoleMapperTest.cs`:

```csharp
using System.Collections.Generic;
using TeampptAddin;
using Xunit;

namespace TeampptAddin.Tests
{
    public class PaletteRoleMapperTest
    {
        private static AssetColor C(string role, string val) =>
            new AssetColor { Role = role, Value = val, Locked = false };

        [Fact]
        public void Null_Or_Empty_Returns_Null()
        {
            Assert.Null(PaletteRoleMapper.Map(null));
            Assert.Null(PaletteRoleMapper.Map(new List<AssetColor>()));
        }

        [Fact]
        public void Uses_Explicit_Roles_When_Present()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor>
            {
                C("main", "#2563EB"),
                C("text", "#1E293B"),
                C("background", "#FFFFFF"),
            });
            Assert.Equal("#2563EB", np.Main);
            Assert.Equal("#1E293B", np.Text);
            Assert.Equal("#FFFFFF", np.Background);
        }

        [Fact]
        public void Fills_Missing_Background_From_Main_Tint()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor> { C("main", "#2563EB") });
            Assert.Equal("#2563EB", np.Main);
            // 배경은 매우 밝게 보강됨
            Assert.True(ColorHsl.FromHex(np.Background).L > 0.85);
        }

        [Fact]
        public void Fills_Missing_Text_Dark()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor> { C("main", "#2563EB") });
            Assert.True(ColorHsl.FromHex(np.Text).L < 0.3);
        }

        [Fact]
        public void Picks_Most_Saturated_As_Main_When_No_Role()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor>
            {
                C("", "#808080"), // 무채색
                C("", "#2563EB"), // 채도 높음
            });
            Assert.Equal("#2563EB", np.Main);
        }

        [Fact]
        public void Sub1_Sub2_Always_Populated()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor> { C("main", "#2563EB") });
            Assert.False(string.IsNullOrEmpty(np.Sub1));
            Assert.False(string.IsNullOrEmpty(np.Sub2));
        }
    }
}
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter PaletteRoleMapperTest`
Expected: FAIL — `PaletteRoleMapper` 없음.

- [ ] **Step 3: 구현 작성**

Create `src/TeampptAddin/Services/PaletteRoleMapper.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    public static class PaletteRoleMapper
    {
        public static NormalizedPalette Map(List<AssetColor> colors)
        {
            var valid = (colors ?? new List<AssetColor>())
                .Where(c => c != null && !string.IsNullOrWhiteSpace(c.Value))
                .ToList();
            if (valid.Count == 0) return null;

            string main = FindByRole(valid, "main", "primary")
                          ?? MostSaturated(valid)
                          ?? valid[0].Value;

            string text = FindByRole(valid, "text", "foreground")
                          ?? Darkest(valid)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main), 0.15));

            string background = FindByRole(valid, "background", "bg")
                          ?? Lightest(valid, 0.85)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main), 0.96));

            var subCandidates = valid
                .Where(c => RoleContains(c.Role, "sub", "accent", "secondary"))
                .Select(c => c.Value)
                .ToList();

            string sub1 = subCandidates.ElementAtOrDefault(0)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main),
                                Clamp01(ColorHsl.FromHex(main).L + 0.18)));
            string sub2 = subCandidates.ElementAtOrDefault(1)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main),
                                Clamp01(ColorHsl.FromHex(main).L - 0.18)));

            return new NormalizedPalette
            {
                Background = background,
                Main = main,
                Sub1 = sub1,
                Sub2 = sub2,
                Text = text
            };
        }

        private static string FindByRole(List<AssetColor> colors, params string[] keys)
        {
            var hit = colors.FirstOrDefault(c => RoleContains(c.Role, keys));
            return hit?.Value;
        }

        private static bool RoleContains(string role, params string[] keys)
        {
            if (string.IsNullOrEmpty(role)) return false;
            var r = role.ToLowerInvariant();
            return keys.Any(k => r.Contains(k));
        }

        private static string MostSaturated(List<AssetColor> colors)
        {
            return colors
                .OrderByDescending(c => ColorHsl.FromHex(c.Value).S)
                .FirstOrDefault()?.Value;
        }

        private static string Darkest(List<AssetColor> colors)
        {
            var ordered = colors.OrderBy(c => ColorHsl.FromHex(c.Value).L).ToList();
            // 가장 어두운 색이 충분히 어두울 때만 텍스트로 인정
            if (ordered.Count > 0 && ColorHsl.FromHex(ordered[0].Value).L < 0.3)
                return ordered[0].Value;
            return null;
        }

        private static string Lightest(List<AssetColor> colors, double minL)
        {
            var hit = colors
                .OrderByDescending(c => ColorHsl.FromHex(c.Value).L)
                .FirstOrDefault(c => ColorHsl.FromHex(c.Value).L >= minL);
            return hit?.Value;
        }

        private static double Clamp01(double v) => v < 0 ? 0 : (v > 1 ? 1 : v);
    }
}
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter PaletteRoleMapperTest`
Expected: PASS (6개).

- [ ] **Step 5: 커밋**

```bash
git add src/TeampptAddin/Services/PaletteRoleMapper.cs src/TeampptAddin.Tests/PaletteRoleMapperTest.cs
git commit -m "feat(style): PaletteRoleMapper — 에셋 색을 5역할로 정규화(빈 역할 보강)"
```

---

### Task 4: PaletteGenerator — 원본 + 반전 생성

**Files:**
- Create: `src/TeampptAddin/Services/PaletteGenerator.cs`
- Test: `src/TeampptAddin.Tests/PaletteGeneratorTest.cs`

**Interfaces:**
- Consumes: `ColorHsl` (Task 1), `NormalizedPalette` (Task 2), `StylePalette`/`PaletteColors` (기존+Task 2).
- Produces: `static System.Collections.Generic.List<StylePalette> PaletteGenerator.Generate(NormalizedPalette np)` — 항상 2개: `[0]` id `"asset-original"` 이름 `"원본"`, `[1]` id `"asset-inverted"` 이름 `"반전"`. np가 null이면 빈 리스트.
- 반전 규칙: Hue 고정. Background.L ↔ Text.L 스왑. Main/Sub1/Sub2는 H·S 유지하고 새 배경 대비 WCAG ≥ 4.5가 되도록 `ColorHsl.AdjustForContrast`로 L 조정.

- [ ] **Step 1: 테스트 작성**

Create `src/TeampptAddin.Tests/PaletteGeneratorTest.cs`:

```csharp
using TeampptAddin;
using Xunit;

namespace TeampptAddin.Tests
{
    public class PaletteGeneratorTest
    {
        private static NormalizedPalette DarkSample() => new NormalizedPalette
        {
            Background = "#0A1428", // 다크
            Main = "#5DBEE0",
            Sub1 = "#3B82F6",
            Sub2 = "#93C5FD",
            Text = "#FFFFFF"        // 밝음
        };

        [Fact]
        public void Null_Returns_Empty()
        {
            Assert.Empty(PaletteGenerator.Generate(null));
        }

        [Fact]
        public void Produces_Original_And_Inverted()
        {
            var list = PaletteGenerator.Generate(DarkSample());
            Assert.Equal(2, list.Count);
            Assert.Equal("asset-original", list[0].Id);
            Assert.Equal("asset-inverted", list[1].Id);
            Assert.Equal("원본", list[0].Name);
            Assert.Equal("반전", list[1].Name);
        }

        [Fact]
        public void Original_Mirrors_Input()
        {
            var orig = PaletteGenerator.Generate(DarkSample())[0];
            Assert.Equal("#0A1428", orig.Colors.Background);
            Assert.Equal("#5DBEE0", orig.Colors.Main);
            Assert.Equal("#FFFFFF", orig.Colors.Text);
        }

        [Fact]
        public void Inverted_Flips_Background_And_Text_Lightness()
        {
            var list = PaletteGenerator.Generate(DarkSample());
            double origBgL = ColorHsl.FromHex(list[0].Colors.Background).L;
            double invBgL = ColorHsl.FromHex(list[1].Colors.Background).L;
            double invTextL = ColorHsl.FromHex(list[1].Colors.Text).L;
            Assert.True(invBgL > origBgL);     // 배경이 밝아짐
            Assert.True(invTextL < invBgL);    // 텍스트는 배경보다 어두움
        }

        [Fact]
        public void Inverted_Text_Meets_Contrast()
        {
            var inv = PaletteGenerator.Generate(DarkSample())[1];
            Assert.True(ColorHsl.ContrastRatio(inv.Colors.Text, inv.Colors.Background) >= 4.5);
        }

        [Fact]
        public void Inverted_Main_Meets_Contrast_And_Keeps_Hue()
        {
            var inv = PaletteGenerator.Generate(DarkSample())[1];
            Assert.True(ColorHsl.ContrastRatio(inv.Colors.Main, inv.Colors.Background) >= 4.5);
            double origH = ColorHsl.FromHex("#5DBEE0").H;
            double invH = ColorHsl.FromHex(inv.Colors.Main).H;
            Assert.True(System.Math.Abs(origH - invH) < 2.0); // Hue 유지
        }
    }
}
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter PaletteGeneratorTest`
Expected: FAIL — `PaletteGenerator` 없음.

- [ ] **Step 3: 구현 작성**

Create `src/TeampptAddin/Services/PaletteGenerator.cs`:

```csharp
using System.Collections.Generic;

namespace TeampptAddin
{
    public static class PaletteGenerator
    {
        private const double TargetContrast = 4.5;

        public static List<StylePalette> Generate(NormalizedPalette np)
        {
            var result = new List<StylePalette>();
            if (np == null) return result;

            result.Add(Original(np));
            result.Add(Inverted(np));
            return result;
        }

        private static StylePalette Original(NormalizedPalette np)
        {
            return new StylePalette
            {
                Id = "asset-original",
                Name = "원본",
                Mood = new List<string> { "에셋 원본" },
                UseWhen = "에셋의 원래 색감을 유지할 때",
                Colors = new PaletteColors
                {
                    Background = np.Background,
                    Main = np.Main,
                    Sub1 = np.Sub1,
                    Sub2 = np.Sub2,
                    Text = np.Text
                }
            };
        }

        private static StylePalette Inverted(NormalizedPalette np)
        {
            var bg = ColorHsl.FromHex(np.Background);
            var text = ColorHsl.FromHex(np.Text);

            // 배경 ↔ 텍스트 명도 스왑 (Hue·S 유지)
            string newBg = ColorHsl.ToHex(ColorHsl.WithLightness(bg, text.L));
            string newText = ColorHsl.ToHex(ColorHsl.WithLightness(text, bg.L));

            // 텍스트가 새 배경에서 대비 부족하면 보정
            newText = ColorHsl.AdjustForContrast(newText, newBg, TargetContrast);

            // 강조색은 H·S 유지, 새 배경 대비 충족하도록 L 조정
            string newMain = ColorHsl.AdjustForContrast(np.Main, newBg, TargetContrast);
            string newSub1 = ColorHsl.AdjustForContrast(np.Sub1, newBg, TargetContrast);
            string newSub2 = ColorHsl.AdjustForContrast(np.Sub2, newBg, TargetContrast);

            return new StylePalette
            {
                Id = "asset-inverted",
                Name = "반전",
                Mood = new List<string> { "반전", "대비 전환" },
                UseWhen = "밝기/배경을 뒤집어 다른 컨셉을 줄 때",
                Colors = new PaletteColors
                {
                    Background = newBg,
                    Main = newMain,
                    Sub1 = newSub1,
                    Sub2 = newSub2,
                    Text = newText
                }
            };
        }
    }
}
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false --filter PaletteGeneratorTest`
Expected: PASS (6개).

- [ ] **Step 5: 전체 단위 테스트 회귀 확인**

Run: `dotnet test src/TeampptAddin.Tests/TeampptAddin.Tests.csproj -p:RegisterForComInterop=false`
Expected: PASS (기존 테스트 + 신규 전부).

- [ ] **Step 6: 커밋**

```bash
git add src/TeampptAddin/Services/PaletteGenerator.cs src/TeampptAddin.Tests/PaletteGeneratorTest.cs
git commit -m "feat(style): PaletteGenerator — 원본+반전 팔레트 생성(명도반전·대비보정)"
```

---

### Task 5: SlideStyleApplier — 활성 슬라이드 적용

> **참고:** PPT Interop은 단위 테스트가 불가하다. 이 태스크는 빌드(관리자 MSBuild) + PowerPoint 수동 검증으로 확인한다.

**Files:**
- Create: `src/TeampptAddin/Core/SlideStyleApplier.cs`
- Modify: `src/TeampptAddin/UI/TaskPaneHost.cs:269-305` (`OnStyleApply`가 `SlideStyleApplier.Apply` 호출로 위임)

**Interfaces:**
- Consumes: `ColorHsl`(Task 1), `StylePalette`/`PaletteColors`(Task 2), `StyleFont`(기존), `Microsoft.Office.Interop.PowerPoint`.
- Produces: `static void SlideStyleApplier.Apply(PowerPoint.Slide slide, StylePalette palette, StyleFont font)`.

적용 규칙(설계 §5):
- 슬라이드 배경: `palette.Colors.Background` 가 null이 아니면 solid로 채움. null이면 배경 미변경.
- 각 shape: `msoPicture`/`msoPlaceholder`/`msoMedia`/`msoOLEControlObject`는 **스킵**(이미지·로고 보호).
- fill이 보이는 도형 → `Main`. line이 보이는 도형 → `Sub1`. 텍스트 → 단락별 폰트(선택 폰트) + 색 `Text`.
- 모든 shape 처리는 try/catch로 감싸 개별 실패가 전체를 막지 않게(기존 패턴 계승).

- [ ] **Step 1: 구현 작성**

Create `src/TeampptAddin/Core/SlideStyleApplier.cs`:

```csharp
using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public static class SlideStyleApplier
    {
        public static void Apply(PowerPoint.Slide slide, StylePalette palette, StyleFont font)
        {
            if (slide == null || palette?.Colors == null) return;
            var c = palette.Colors;

            // ① 슬라이드 배경
            if (!string.IsNullOrEmpty(c.Background))
            {
                try
                {
                    slide.FollowMasterBackground = MsoTriState.msoFalse;
                    slide.Background.Fill.Visible = MsoTriState.msoTrue;
                    slide.Background.Fill.Solid();
                    slide.Background.Fill.ForeColor.RGB = Ole(c.Background);
                }
                catch { }
            }

            // ② 도형들
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                try
                {
                    if (IsProtected(shape)) continue;

                    // 도형 채우기
                    if (!string.IsNullOrEmpty(c.Main))
                    {
                        try
                        {
                            if (shape.Fill.Visible == MsoTriState.msoTrue)
                                shape.Fill.ForeColor.RGB = Ole(c.Main);
                        }
                        catch { }
                    }

                    // 도형 선
                    if (!string.IsNullOrEmpty(c.Sub1))
                    {
                        try
                        {
                            if (shape.Line.Visible == MsoTriState.msoTrue)
                                shape.Line.ForeColor.RGB = Ole(c.Sub1);
                        }
                        catch { }
                    }

                    // 텍스트
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        var tr = shape.TextFrame.TextRange;
                        int count = tr.Paragraphs().Count;
                        for (int i = 1; i <= count; i++)
                        {
                            var para = tr.Paragraphs(i);
                            if (font != null && !string.IsNullOrEmpty(font.Name))
                                para.Font.Name = font.Name;
                            if (!string.IsNullOrEmpty(c.Text))
                                para.Font.Color.RGB = Ole(c.Text);
                        }
                    }
                }
                catch { }
            }
        }

        private static bool IsProtected(PowerPoint.Shape shape)
        {
            switch (shape.Type)
            {
                case MsoShapeType.msoPicture:
                case MsoShapeType.msoPlaceholder:
                case MsoShapeType.msoMedia:
                case MsoShapeType.msoOLEControlObject:
                case MsoShapeType.msoEmbeddedOLEObject:
                    return true;
                default:
                    return false;
            }
        }

        private static int Ole(string hex)
        {
            var h = ColorHsl.ToHex(ColorHsl.FromHex(hex)); // 정규화
            int r = Convert.ToInt32(h.Substring(1, 2), 16);
            int g = Convert.ToInt32(h.Substring(3, 2), 16);
            int b = Convert.ToInt32(h.Substring(5, 2), 16);
            return ColorTranslator.ToOle(Color.FromArgb(r, g, b));
        }
    }
}
```

- [ ] **Step 2: OnStyleApply 위임으로 교체**

`src/TeampptAddin/UI/TaskPaneHost.cs`의 `OnStyleApply`(269~305) 본문을 아래로 교체. 기존 `ColorFromHex`(307~314)는 다른 곳에서 안 쓰면 함께 제거(쓰이면 유지):

```csharp
        private void OnStyleApply(StylePalette palette, StyleFont font)
        {
            try
            {
                var slide = (PowerPoint.Slide)Globals.Application.ActiveWindow.View.Slide;
                SlideStyleApplier.Apply(slide, palette, font);

                _wpfPanel.SetStatus(
                    $"✓ {palette?.Name ?? "팔레트"} · {font?.Name ?? "폰트"} 적용 완료",
                    ThemeResources.StatusSuccess.Color);
            }
            catch (Exception ex)
            {
                _wpfPanel.SetStatus($"적용 실패: {ex.Message}", ThemeResources.StatusError.Color);
                Logger.Log($"StyleApply fail: {ex}");
            }
        }
```

> `ColorFromHex` 제거 여부: `src/TeampptAddin/UI/TaskPaneHost.cs`에서 `ColorFromHex` 참조를 검색해 다른 사용처가 없으면 메서드를 지운다. 사용처가 있으면 그대로 둔다.

- [ ] **Step 3: 메인 프로젝트 빌드 (관리자 MSBuild)**

PowerPoint를 닫고, CLAUDE.md의 관리자 빌드 명령 실행 후 `build.log` 끝 5줄 확인:

```
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

- [ ] **Step 4: 수동 검증 (PowerPoint)**

1. PowerPoint 실행 → 애드인 패널 표시.
2. 도형·텍스트가 있는 슬라이드에서 스타일 탭 → 아무 팔레트나 클릭.
3. 확인: 슬라이드 배경·도형 채우기·도형 선·텍스트 색이 팔레트대로 바뀐다. 폰트도 바뀐다.
4. 슬라이드에 그림/이미지가 있으면 **그 그림은 색이 안 변하고 그대로**인지 확인.
5. "✓ … 적용 완료" 상태 메시지 표시.

- [ ] **Step 5: 커밋**

```bash
git add src/TeampptAddin/Core/SlideStyleApplier.cs src/TeampptAddin/UI/TaskPaneHost.cs
git commit -m "feat(style): SlideStyleApplier — 배경·도형·선·텍스트·폰트 적용(이미지 보호)"
```

---

### Task 6: UI 연결 — 앵커 캡처 + 생성 팔레트 표시

> **참고:** WPF UI라 단위 테스트 없음. 빌드 + PowerPoint 수동 검증.

**Files:**
- Modify: `src/TeampptAddin/UI/Wpf/AssetPanel.cs` — 앵커 필드, `SetStyleAnchorByFile`, `BuildEffectiveStyleConfig`, `PopulateStylePanel` 갱신, 원격 삽입 경로 연결
- Modify: `src/TeampptAddin/UI/TaskPaneHost.cs:225-238` (`OnWpfClickInsert`에서 앵커 설정)

**Interfaces:**
- Consumes: `PaletteRoleMapper`(Task 3), `PaletteGenerator`(Task 4), `HeaderAsset`/`AssetColor`/`AssetFont`(기존), `StyleConfig`/`StylePalette`/`StyleFont`(기존).
- Produces:
  - `AssetPanel.SetStyleAnchorByFile(string pathOrFile)` (public) — `_allAssets`에서 파일명 매칭으로 앵커 에셋을 찾아 스타일 탭을 다시 그린다.
  - `AssetPanel` 내부: `_anchorAsset`(HeaderAsset), `BuildEffectiveStyleConfig()` → 생성 팔레트(원본·반전) + 기존 프리셋, 앵커 폰트 + 기존 폰트 병합한 `StyleConfig` 반환.

- [ ] **Step 1: AssetPanel에 앵커 필드 + 메서드 추가**

`src/TeampptAddin/UI/Wpf/AssetPanel.cs`의 필드 영역(예: `_styleConfig` 선언 부근, 29행 근처)에 추가:

```csharp
        private HeaderAsset _anchorAsset;
```

같은 파일의 `SetAssets`/`InitAi` 부근(1596행 근처)에 public 메서드 추가:

```csharp
        public void SetStyleAnchorByFile(string pathOrFile)
        {
            if (string.IsNullOrEmpty(pathOrFile)) return;
            var key = System.IO.Path.GetFileName(pathOrFile);
            _anchorAsset = (_allAssets ?? new List<HeaderAsset>())
                .FirstOrDefault(a =>
                    string.Equals(System.IO.Path.GetFileName(a.File ?? ""), key,
                        StringComparison.OrdinalIgnoreCase)
                    || (a.Extra != null && a.Extra.TryGetValue("remote_file", out var rf)
                        && string.Equals(System.IO.Path.GetFileName(rf.ToString()), key,
                            StringComparison.OrdinalIgnoreCase)));
            PopulateStylePanel();
        }

        private StyleConfig BuildEffectiveStyleConfig()
        {
            var palettes = new List<StylePalette>();
            var fonts = new List<StyleFont>();

            if (_anchorAsset != null)
            {
                var np = PaletteRoleMapper.Map(_anchorAsset.Colors);
                palettes.AddRange(PaletteGenerator.Generate(np)); // 원본 + 반전

                if (_anchorAsset.Fonts != null)
                {
                    foreach (var f in _anchorAsset.Fonts)
                    {
                        if (f == null || string.IsNullOrWhiteSpace(f.Family)) continue;
                        if (fonts.Any(x => string.Equals(x.Name, f.Family,
                                StringComparison.OrdinalIgnoreCase))) continue;
                        fonts.Add(new StyleFont { Name = f.Family, Mood = new List<string>(),
                            UseWhen = "에셋 폰트" });
                    }
                }
            }

            // 대체 컨셉 프리셋(기존 styles.json) 뒤에 붙임
            if (_styleConfig?.Palettes != null) palettes.AddRange(_styleConfig.Palettes);
            if (_styleConfig?.Fonts != null)
            {
                foreach (var f in _styleConfig.Fonts)
                {
                    if (f == null || string.IsNullOrWhiteSpace(f.Name)) continue;
                    if (fonts.Any(x => string.Equals(x.Name, f.Name,
                            StringComparison.OrdinalIgnoreCase))) continue;
                    fonts.Add(f);
                }
            }

            return new StyleConfig { Palettes = palettes, Fonts = fonts };
        }
```

> 파일 상단 using에 `System`, `System.Linq`, `System.Collections.Generic`가 이미 있는지 확인하고 없으면 추가. (`System.IO`는 기존 사용 중.)

- [ ] **Step 2: PopulateStylePanel이 BuildEffectiveStyleConfig를 쓰도록 변경**

`PopulateStylePanel`(1315행) 시작부의 가드와 소스를 교체:

```csharp
        private void PopulateStylePanel()
        {
            if (_styleStack == null) return;
            var effective = BuildEffectiveStyleConfig();
            _styleStack.Children.Clear();

            var palettes = effective.Palettes ?? new List<StylePalette>();
            var fonts    = effective.Fonts    ?? new List<StyleFont>();

            if (palettes.Count == 0 && fonts.Count == 0) return;

            _selectedPalette = palettes.Count > 0 ? palettes[0] : null;
            _selectedFont    = fonts.Count    > 0 ? fonts[0]    : null;
            // 이하 기존 코드(컬러 팔레트 라벨부터) 그대로 유지
```

> 나머지(팔레트 wrap, 폰트 패널 등 1326행 이후)는 그대로 둔다. 변경점은 소스가 `_styleConfig` → `effective`로 바뀐 것과 상단 가드뿐이다. 클릭 시에만 `StyleApplyRequested`가 발생하므로 선택만으로 슬라이드가 바뀌지 않는다(수동 적용 유지).

- [ ] **Step 3: 로컬 삽입 경로에 앵커 연결**

`src/TeampptAddin/UI/TaskPaneHost.cs`의 `OnWpfClickInsert`(225~238)에서 삽입 성공 직후 앵커 설정:

```csharp
            try
            {
                ShapeInserter.InsertToActiveSlide(card.PptxPath);
                _wpfPanel.SetStyleAnchorByFile(card.PptxPath);
                _wpfPanel.SetStatus($"✓ {card.Title} 삽입 완료",
                    ThemeResources.StatusSuccess.Color);
            }
```

- [ ] **Step 4: 원격 삽입 경로에 앵커 연결**

`src/TeampptAddin/UI/Wpf/AssetPanel.cs`의 `InsertRemoteAssetAsync`(1164)에서 삽입 성공 직후:

```csharp
                        ShapeInserter.InsertToActiveSlide(localPath);
                        SetStyleAnchorByFile(remotePath);
                        SetStatus($"✓ {assetName} 삽입 완료", ThemeResources.StatusSuccess.Color);
```

- [ ] **Step 5: 메인 프로젝트 빌드 (관리자 MSBuild)**

PowerPoint를 닫고 관리자 빌드 → `build.log` 끝 5줄에 `0 Error(s)` 확인.

- [ ] **Step 6: 수동 검증 (PowerPoint)**

1. 애드인 패널에서 에셋 하나를 클릭 삽입.
2. 스타일 탭으로 이동 → **맨 앞에 "원본"·"반전" 팔레트**가 보이고, 그 뒤에 기존 프리셋들이 보인다.
3. 폰트 목록 맨 앞에 에셋 폰트가 보인다.
4. "반전" 클릭 → 슬라이드가 라이트(또는 반대) 컨셉으로 바뀐다(목업 2번처럼).
5. "원본" 클릭 → 에셋 원래 색으로 돌아온다.
6. 다른 에셋을 삽입하면 스타일 탭의 원본·반전이 그 에셋 색으로 갱신된다.
7. 앵커가 없을 때(에셋 미삽입 최초 상태)는 기존 프리셋만 보인다(빈 화면 아님).

- [ ] **Step 7: 커밋**

```bash
git add src/TeampptAddin/UI/Wpf/AssetPanel.cs src/TeampptAddin/UI/TaskPaneHost.cs
git commit -m "feat(style): 삽입 에셋 앵커로 원본+반전 팔레트 스타일 탭 연결"
```

---

## 완료 후

- PROGRESS-BOARD.md의 현재 잎(Task)을 갱신: 스타일 탭 v1 구현 완료, 수동 검증 결과 기록. **다음 단계 = 덱 전체 통일**(spec §8 경고 박스)을 잎으로 올림.
- 설계 §8의 안전 변주(단색·유사색)는 v1 범위 밖. 원본+반전이 핵심이고 검증 후 필요하면 추가.

## Notes / 미해결

- **앵커 폴백(슬라이드 도형 색 샘플링) 의도적 보류**: 설계 §2 결정7은 앵커 출처를 ①삽입 에셋 → ②활성 슬라이드 도형 샘플링 → ③프리셋만 순으로 둔다. v1 plan은 ①과 ③만 구현하고 ②(슬라이드에서 색을 직접 읽어 앵커 생성)는 보류한다. 이유: 삽입 시 매번 앵커가 잡히므로 주 흐름은 ①로 커버되고, ②는 Interop 색 샘플링 복잡도를 더한다. 기존 덱을 열고 삽입 없이 스타일 탭을 여는 경우엔 ③(프리셋만)으로 떨어진다. 필요해지면 별도 태스크로 추가.
- **반전 시 슬라이드 배경 채우기**: `slide.Background.Fill`은 마스터 배경을 따르는 경우가 있어 `FollowMasterBackground=msoFalse` 후 solid로 덮는다. 일부 템플릿에서 효과가 약하면 Task 5 수동 검증 단계에서 확인하고, 슬라이드 크기 사각형 배경 도형을 까는 방식(설계 §5 규약 7과 동일 발상)으로 폴백 가능 — 단 v1에서는 배경 Fill 우선 시도.

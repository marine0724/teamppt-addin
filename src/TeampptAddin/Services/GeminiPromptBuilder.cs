using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace TeampptAddin
{
    public static class GeminiPromptBuilder
    {
        public static string BuildSystemPrompt(
            List<CatalogEntry> catalog,
            IEnumerable<StylePalette> palettes,
            IEnumerable<StyleFont> fonts)
        {
            var catalogJson = JsonConvert.SerializeObject(catalog, Formatting.Indented);

            var paletteSummaries = palettes.Select(p => new
            {
                p.Id, p.Name, p.Mood, p.UseWhen
            });
            var palettesJson = JsonConvert.SerializeObject(paletteSummaries, Formatting.Indented);

            var fontSummaries = fonts.Select(f => new
            {
                f.Name, f.Mood, f.UseWhen
            });
            var fontsJson = JsonConvert.SerializeObject(fontSummaries, Formatting.Indented);

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
- 적합한 에셋이 없으면 솔직하게 ""현재 보유한 에셋 중에는 딱 맞는 것이 없다""고 말해.
- 사용자의 요청이 모호하면 바로 추천하지 말고, 먼저 질문해서 의도를 파악해.

## 응답 형식
반드시 아래 JSON으로 응답해.

**추천할 에셋이 있을 때:**
```json
{{
  ""message"": ""추천 설명 (한국어, 1~2문장)"",
  ""assets"": [
    {{ ""file"": ""header_N.pptx"", ""reason"": ""구체적 추천 이유"" }}
  ],
  ""palette"": {{ ""id"": ""팔레트id"", ""reason"": ""이유"" }},
  ""font"": {{ ""name"": ""폰트이름"", ""reason"": ""이유"" }}
}}
```

**질문이 필요하거나 적합한 에셋이 없을 때:**
```json
{{
  ""message"": ""질문 또는 안내 메시지 (한국어)"",
  ""assets"": [],
  ""palette"": null,
  ""font"": null
}}
```";
        }

        public static string BuildUserPrompt(string userIntent)
        {
            return userIntent;
        }
    }
}

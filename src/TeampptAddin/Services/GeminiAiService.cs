using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public class GeminiAiService : IAiService
    {
        private static readonly HttpClient Http = new HttpClient();
        private readonly string _apiKey;

        public GeminiAiService(string apiKey)
        {
            _apiKey = apiKey;
        }

        public static GeminiAiService FromAssetsDir(string assetsDir)
        {
            var path = Path.Combine(assetsDir, "api-keys.json");
            var json = File.ReadAllText(path, Encoding.UTF8);
            var obj = JObject.Parse(json);
            var key = obj["gemini"]?.ToString()
                ?? throw new InvalidOperationException("api-keys.json에 'gemini' 키가 없습니다.");
            return new GeminiAiService(key);
        }

        public async Task<AiRecommendation> RecommendAsync(
            string userIntent,
            IEnumerable<HeaderAsset> assets,
            IEnumerable<StylePalette> palettes,
            IEnumerable<StyleFont> fonts)
        {
            var assetList = assets.ToList();
            var paletteList = palettes.ToList();
            var fontList = fonts.ToList();

            var catalog = CatalogBuilder.Build(assetList);

            var systemPrompt = GeminiPromptBuilder.BuildSystemPrompt(
                catalog, paletteList, fontList);
            var userPrompt = GeminiPromptBuilder.BuildUserPrompt(userIntent);

            var requestBody = new
            {
                contents = new[]
                {
                    new
                    {
                        role = "user",
                        parts = new[] { new { text = userPrompt } }
                    }
                },
                systemInstruction = new
                {
                    parts = new[] { new { text = systemPrompt } }
                },
                generationConfig = new
                {
                    temperature = 0.7,
                    responseMimeType = "application/json",
                    thinkingConfig = new { thinkingBudget = 1024 }
                }
            };

            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={_apiKey}";
            var content = new StringContent(
                JsonConvert.SerializeObject(requestBody),
                Encoding.UTF8, "application/json");

            var response = await Http.PostAsync(url, content).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException(
                    $"Gemini API 오류 ({(int)response.StatusCode}): {body}");

            var root = JObject.Parse(body);

            LogTokenUsage(root);

            var text = root["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();
            if (string.IsNullOrEmpty(text))
                throw new InvalidOperationException("Gemini 응답에 텍스트가 없습니다.");

            return ParseResponse(text, assetList, paletteList, fontList);
        }

        private void LogTokenUsage(JObject root)
        {
            var usage = root["usageMetadata"];
            if (usage == null) return;

            var input = usage["promptTokenCount"]?.Value<int>() ?? 0;
            var output = usage["candidatesTokenCount"]?.Value<int>() ?? 0;
            var total = usage["totalTokenCount"]?.Value<int>() ?? 0;

            Logger.Log($"[Gemini] 토큰 사용: input={input}, output={output}, total={total}");
        }

        private AiRecommendation ParseResponse(
            string json,
            List<HeaderAsset> assets,
            List<StylePalette> palettes,
            List<StyleFont> fonts)
        {
            var obj = JObject.Parse(json);

            var message = obj["message"]?.ToString() ?? "";

            var assetSuggestions = new List<AssetSuggestion>();
            var assetArray = obj["assets"] as JArray;
            if (assetArray != null)
            {
                foreach (var item in assetArray)
                {
                    var file = item["file"]?.ToString();
                    var reason = item["reason"]?.ToString() ?? "";
                    var match = assets.FirstOrDefault(a =>
                        string.Equals(a.File, file, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                        assetSuggestions.Add(new AssetSuggestion { Asset = match, Reason = reason });
                }
            }

            StylePalette matchedPalette = null;
            string paletteReason = "";
            var paletteObj = obj["palette"];
            if (paletteObj != null && paletteObj.Type != JTokenType.Null)
            {
                var pid = paletteObj["id"]?.ToString();
                paletteReason = paletteObj["reason"]?.ToString() ?? "";
                matchedPalette = palettes.FirstOrDefault(p =>
                    string.Equals(p.Id, pid, StringComparison.OrdinalIgnoreCase));
            }

            StyleFont matchedFont = null;
            string fontReason = "";
            var fontObj = obj["font"];
            if (fontObj != null && fontObj.Type != JTokenType.Null)
            {
                var fname = fontObj["name"]?.ToString();
                fontReason = fontObj["reason"]?.ToString() ?? "";
                matchedFont = fonts.FirstOrDefault(f =>
                    string.Equals(f.Name, fname, StringComparison.OrdinalIgnoreCase));
            }

            return new AiRecommendation
            {
                Message = message,
                Assets = assetSuggestions,
                Style = new StyleSuggestion
                {
                    Palette = matchedPalette,
                    Font = matchedFont,
                    Reason = string.Join("; ", new[] { paletteReason, fontReason }
                        .Where(r => !string.IsNullOrEmpty(r)))
                }
            };
        }
    }
}

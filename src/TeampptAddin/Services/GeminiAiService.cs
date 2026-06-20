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
        private readonly List<JObject> _history = new List<JObject>();
        private const int MaxHistoryTurns = 10;

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

            _history.Add(new JObject
            {
                ["role"] = "user",
                ["parts"] = new JArray { new JObject { ["text"] = userPrompt } }
            });

            while (_history.Count > MaxHistoryTurns * 2)
                _history.RemoveAt(0);

            var requestBody = new JObject
            {
                ["contents"] = new JArray(_history.ToArray()),
                ["systemInstruction"] = new JObject
                {
                    ["parts"] = new JArray { new JObject { ["text"] = systemPrompt } }
                },
                ["generationConfig"] = new JObject
                {
                    ["temperature"] = 0.7,
                    ["responseMimeType"] = "application/json",
                    ["responseSchema"] = GeminiPromptBuilder.BuildResponseSchema(),
                    ["thinkingConfig"] = new JObject { ["thinkingBudget"] = 1024 }
                }
            };

            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={_apiKey}";
            var bodyString = requestBody.ToString(Formatting.None);

            const int maxAttempts = 3;
            HttpResponseMessage response = null;
            string body = null;

            Logger.Log($"[Gemini] API Key prefix: {_apiKey.Substring(0, Math.Min(6, _apiKey.Length))}..., length={_apiKey.Length}");
            Logger.Log($"[Gemini] URL host: generativelanguage.googleapis.com, model: gemini-2.5-flash");

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                // StringContent는 전송 후 재사용 불가하므로 시도마다 새로 만든다
                var content = new StringContent(bodyString, Encoding.UTF8, "application/json");

                // 기존 인증 헤더가 있으면 제거 (프록시 등에 의한 오염 방지)
                Http.DefaultRequestHeaders.Authorization = null;

                response = await Http.PostAsync(url, content).ConfigureAwait(false);
                body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                Logger.Log($"[Gemini] Attempt {attempt}: HTTP {(int)response.StatusCode}");

                if (response.IsSuccessStatusCode)
                    break;

                var status = (int)response.StatusCode;
                // 503(과부하)·429(쿼터)·500은 일시적 → 백오프 후 재시도. 403 등은 즉시 실패.
                bool transient = status == 503 || status == 429 || status == 500;
                if (transient && attempt < maxAttempts)
                {
                    var delayMs = 500 * (1 << (attempt - 1)); // 500, 1000, 2000ms
                    Logger.Log($"[Gemini] 일시 오류 {status}, {delayMs}ms 후 재시도 ({attempt}/{maxAttempts})");
                    await Task.Delay(delayMs).ConfigureAwait(false);
                    continue;
                }

                throw new HttpRequestException($"Gemini API 오류 ({status}): {body}");
            }

            var root = JObject.Parse(body);

            LogTokenUsage(root);

            var text = root["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();
            if (string.IsNullOrEmpty(text))
                throw new InvalidOperationException("Gemini 응답에 텍스트가 없습니다.");

            _history.Add(new JObject
            {
                ["role"] = "model",
                ["parts"] = new JArray { new JObject { ["text"] = text } }
            });

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

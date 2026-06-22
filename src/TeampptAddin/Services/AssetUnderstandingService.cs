using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
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

            const int maxAttempts = 5;
            HttpResponseMessage response = null;
            string body = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var content = new StringContent(bodyString, Encoding.UTF8, "application/json");

                response = await Http.PostAsync(url, content).ConfigureAwait(false);
                body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                Logger.Log($"[Understand] {Path.GetFileName(pngPath)} attempt {attempt}: HTTP {(int)response.StatusCode}");

                if (response.IsSuccessStatusCode) break;

                var status = (int)response.StatusCode;
                bool transient = status == 503 || status == 429 || status == 500;
                if (transient && attempt < maxAttempts)
                {
                    var delay = 1000 * (1 << (attempt - 1));
                    Logger.Log($"[Understand] 재시도 대기 {delay}ms...");
                    await Task.Delay(delay).ConfigureAwait(false);
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

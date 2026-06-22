using System;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
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

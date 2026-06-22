using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public class SupabaseClient
    {
        private static readonly HttpClient Http;
        static SupabaseClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            Http = new HttpClient();
        }
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
            var req = new HttpRequestMessage(HttpMethod.Post, $"{_baseUrl}/rest/v1/assets?on_conflict=file");
            ApplyHeaders(req);
            req.Headers.TryAddWithoutValidation("Prefer", "return=minimal,resolution=merge-duplicates");
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

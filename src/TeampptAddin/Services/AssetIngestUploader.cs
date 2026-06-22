using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public class AssetIngestUploader
    {
        private readonly IAccessPolicy _policy;
        public AssetIngestUploader(IAccessPolicy policy = null)
        {
            _policy = policy ?? new LocalFileAccessPolicy();
        }

        public async Task<int> UploadDirectoryAsync(string splitDir, string sourceDeck, System.Action<IngestProgress> onProgress = null, int startFrom = 0)
        {
            if (!_policy.CanIngest)
            {
                Logger.Log("[Upload] 관리자 아님 — 인제스트 거부");
                return 0;
            }

            var cred = AdminCredentials.Load(AdminCredentials.DefaultPath);
            var keys = LoadApiKeys();
            var understand = new AssetUnderstandingService(cred.GeminiKey);
            var embed = new EmbeddingService(cred.GeminiKey);
            var supa = new SupabaseClient(keys.url, cred.SupabaseServiceKey);

            var pngFiles = Directory.GetFiles(splitDir, "*.png").OrderBy(f => f).ToArray();
            int total = 0;
            foreach (var p in pngFiles)
            {
                var pptx = Path.Combine(splitDir, Path.GetFileNameWithoutExtension(p) + ".pptx");
                if (File.Exists(pptx)) total++;
            }

            int count = 0;
            foreach (var pngPath in pngFiles)
            {
                var id = Path.GetFileNameWithoutExtension(pngPath);
                var pptxPath = Path.Combine(splitDir, id + ".pptx");
                if (!File.Exists(pptxPath)) continue;

                count++;
                if (count <= startFrom) continue;

                var category = id.Contains("_") ? id.Substring(0, id.LastIndexOf('_')) : id;
                var prog = new IngestProgress { Index = count, Total = total, AssetId = id, PngPath = pngPath };

                prog.Stage = IngestStage.Understanding;
                onProgress?.Invoke(prog);
                var u = await understand.UnderstandAsync(pngPath, category, "pptx/" + id + ".pptx").ConfigureAwait(false);

                prog.Name = u.Asset?.Name;
                prog.Stage = IngestStage.Embedding;
                onProgress?.Invoke(prog);
                var embedText = EmbedTextBuilder.Build(u);
                var vector = await embed.EmbedAsync(embedText).ConfigureAwait(false);

                var storageKey = ToStorageKey(id);
                prog.Stage = IngestStage.Uploading;
                onProgress?.Invoke(prog);
                await supa.UploadObjectAsync("thumb", storageKey + ".png", File.ReadAllBytes(pngPath), "image/png").ConfigureAwait(false);
                await supa.UploadObjectAsync("pptx", storageKey + ".pptx", File.ReadAllBytes(pptxPath),
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation").ConfigureAwait(false);

                var row = AssetRowBuilder.Build(u, vector, embedText, "pptx/" + storageKey + ".pptx", "thumb/" + storageKey + ".png", sourceDeck);
                await supa.InsertAssetAsync(row).ConfigureAwait(false);

                prog.Stage = IngestStage.AssetDone;
                prog.Kind = u.Asset?.Kind ?? "unknown";
                onProgress?.Invoke(prog);
                Logger.Log($"[Upload] {id} → Supabase OK (kind={prog.Kind})");
            }
            Logger.Log($"[Upload] 완료: {count}개");
            return count;
        }

        private static string ToStorageKey(string id)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(id));
                return string.Concat(bytes.Take(8).Select(b => b.ToString("x2")));
            }
        }

        private (string url, string anon) LoadApiKeys()
        {
            var path = Path.Combine(Globals.AssetsDir, "api-keys.json");
            var o = JObject.Parse(File.ReadAllText(path, Encoding.UTF8));
            return (o["supabaseUrl"]?.ToString(), o["supabaseAnonKey"]?.ToString());
        }
    }
}

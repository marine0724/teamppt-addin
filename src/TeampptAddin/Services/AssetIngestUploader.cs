using System.IO;
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

        public async Task<int> UploadDirectoryAsync(string splitDir, string sourceDeck)
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

        private (string url, string anon) LoadApiKeys()
        {
            var path = Path.Combine(Globals.AssetsDir, "api-keys.json");
            var o = JObject.Parse(File.ReadAllText(path, Encoding.UTF8));
            return (o["supabaseUrl"]?.ToString(), o["supabaseAnonKey"]?.ToString());
        }
    }
}

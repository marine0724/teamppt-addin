using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public static class AssetLoader
    {
        public static List<HeaderAsset> Load(string assetsDir)
        {
            var jsonPath = Path.Combine(assetsDir, "assets.json");

            var assets = File.Exists(jsonPath)
                ? LoadFromJson(jsonPath)
                : ScanFolder(assetsDir);

            return assets
                .Where(a => File.Exists(Path.Combine(assetsDir, a.File)))
                .ToList();
        }

        private static List<HeaderAsset> LoadFromJson(string jsonPath)
        {
            var json = File.ReadAllText(jsonPath);
            var arr = JArray.Parse(json);
            var result = new List<HeaderAsset>();
            foreach (var token in arr)
            {
                if (!(token is JObject obj)) continue;
                var migrated = AssetSchemaMigrator.Migrate(obj);
                result.Add(migrated.ToObject<HeaderAsset>());
            }
            return result;
        }

        private static List<HeaderAsset> ScanFolder(string assetsDir)
        {
            if (!Directory.Exists(assetsDir))
                return new List<HeaderAsset>();

            return Directory.GetFiles(assetsDir, "header_*.pptx")
                .OrderBy(f => f)
                .Select(f => new HeaderAsset
                {
                    SchemaVersion = 2,
                    File = Path.GetFileName(f),
                    Name = Path.GetFileNameWithoutExtension(f),
                    Kind = "component",
                    Category = "헤더",
                    Scope = "slide",
                    ContentFit = new List<string>(),
                    UseWhen = "",
                    GridColumns = 1
                })
                .ToList();
        }
    }
}

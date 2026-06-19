using System.IO;
using System.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetsJsonIntegrationTest
    {
        private static string FindAssetsJson()
        {
            var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "src", "TeampptAddin", "Assets", "assets.json")))
                dir = dir.Parent;
            Assert.NotNull(dir);
            return Path.Combine(dir.FullName, "src", "TeampptAddin", "Assets", "assets.json");
        }

        [Fact]
        public void Real_AssetsJson_Loads_As_V2_With_Header_Scope()
        {
            var assetsJsonPath = FindAssetsJson();
            var tmpDir = Path.Combine(Path.GetTempPath(), "teamppt_integ_" + System.Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tmpDir);
            try
            {
                File.Copy(assetsJsonPath, Path.Combine(tmpDir, "assets.json"));
                for (int i = 1; i <= 7; i++)
                    File.WriteAllText(Path.Combine(tmpDir, $"header_{i}.pptx"), "dummy");

                var assets = AssetLoader.Load(tmpDir);

                Assert.Equal(7, assets.Count);
                Assert.All(assets, a => Assert.Equal(2, a.SchemaVersion));
                var hero = assets.First(a => a.File == "header_1.pptx");
                Assert.Equal("deck", hero.Scope);
                Assert.NotEmpty(hero.Slots);

                var catalog = CatalogBuilder.Build(assets);
                Assert.Equal(7, catalog.Count);
            }
            finally { Directory.Delete(tmpDir, true); }
        }
    }
}

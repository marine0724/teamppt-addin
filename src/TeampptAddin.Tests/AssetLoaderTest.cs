using System.IO;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetLoaderTest
    {
        [Fact]
        public void Load_Migrates_V1_File_To_Role_Colors()
        {
            var dir = Path.Combine(Path.GetTempPath(), "teamppt_test_" + System.Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                File.WriteAllText(Path.Combine(dir, "header_1.pptx"), "dummy");
                File.WriteAllText(Path.Combine(dir, "assets.json"), @"[
                  { ""file"": ""header_1.pptx"", ""name"": ""타이틀"", ""category"": ""헤더"",
                    ""colors"": { ""main"": ""#2563EB"", ""text"": ""#FFFFFF"" } }
                ]");

                var assets = AssetLoader.Load(dir);

                Assert.Single(assets);
                Assert.Equal(2, assets[0].SchemaVersion);
                Assert.NotNull(assets[0].Colors);
                Assert.Contains(assets[0].Colors, c => c.Role == "main" && c.Value == "#2563EB");
            }
            finally { Directory.Delete(dir, true); }
        }
    }
}

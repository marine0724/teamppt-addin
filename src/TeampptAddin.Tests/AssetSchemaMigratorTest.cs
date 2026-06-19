using Newtonsoft.Json.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AssetSchemaMigratorTest
    {
        [Fact]
        public void Migrates_V1_Object_Colors_To_Role_Array()
        {
            var v1 = JObject.Parse(@"{
              ""file"": ""header_1.pptx"", ""name"": ""x"", ""category"": ""헤더"",
              ""colors"": { ""main"": ""#2563EB"", ""sub1"": ""#3B82F6"", ""sub2"": ""#93C5FD"", ""text"": ""#1E293B"" }
            }");

            var v2 = AssetSchemaMigrator.Migrate(v1);

            Assert.Equal(2, (int)v2["schemaVersion"]);
            Assert.Equal("slide", (string)v2["scope"]);
            var colors = (JArray)v2["colors"];
            Assert.Equal(4, colors.Count);
            Assert.Equal("main", (string)colors[0]["role"]);
            Assert.Equal("#2563EB", (string)colors[0]["value"]);
            Assert.False((bool)colors[0]["locked"]);
        }

        [Fact]
        public void Passes_Through_V2_Array_Colors()
        {
            var v2in = JObject.Parse(@"{
              ""schemaVersion"": 2, ""scope"": ""deck"",
              ""colors"": [ { ""role"": ""main"", ""value"": ""#000000"", ""locked"": true } ]
            }");

            var v2 = AssetSchemaMigrator.Migrate(v2in);

            Assert.Equal(2, (int)v2["schemaVersion"]);
            Assert.Equal("deck", (string)v2["scope"]);
            var colors = (JArray)v2["colors"];
            Assert.Single(colors);
            Assert.True((bool)colors[0]["locked"]);
        }
    }
}

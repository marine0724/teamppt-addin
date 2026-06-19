using Newtonsoft.Json;
using Xunit;

namespace TeampptAddin.Tests
{
    public class HeaderAssetDeserializeTest
    {
        [Fact]
        public void Deserializes_V2_Asset_With_Roles_Fonts_Slots()
        {
            const string json = @"{
              ""schemaVersion"": 2,
              ""file"": ""header_3.pptx"",
              ""name"": ""장점 나열"",
              ""kind"": ""layout"",
              ""category"": ""헤더"",
              ""scope"": ""deck"",
              ""provenance"": ""IR 덱"",
              ""colors"": [ { ""role"": ""main"", ""value"": ""#2563EB"", ""locked"": false } ],
              ""fonts"": [ { ""role"": ""heading"", ""family"": ""Pretendard"", ""fallback"": ""맑은 고딕"", ""weight"": ""Bold"", ""source"": ""bundled"" } ],
              ""slots"": [ { ""name"": ""title"", ""type"": ""text"", ""perSlide"": true } ]
            }";

            var a = JsonConvert.DeserializeObject<HeaderAsset>(json);

            Assert.Equal(2, a.SchemaVersion);
            Assert.Equal("layout", a.Kind);
            Assert.Equal("deck", a.Scope);
            Assert.Equal("IR 덱", a.Provenance);
            Assert.Single(a.Colors);
            Assert.Equal("main", a.Colors[0].Role);
            Assert.False(a.Colors[0].Locked);
            Assert.Equal("Pretendard", a.Fonts[0].Family);
            Assert.True(a.Slots[0].PerSlide);
        }
    }
}

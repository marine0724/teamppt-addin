using Xunit;

namespace TeampptAddin.Tests
{
    public class UnderstandingParserTest
    {
        private const string Sample = @"{
          ""name"": ""우측정렬 연도강조 표지"",
          ""kind"": ""layout"",
          ""use_when"": ""연도를 강조하는 표지가 필요할 때"",
          ""content_fit"": [""표지"", ""연도 강조""],
          ""tags"": [""표지"", ""연도"", ""미니멀""],
          ""example_intents"": [""투자 유치 IR 표지"", ""회사 소개 첫 장""],
          ""slots"": [{""name"":""title"",""type"":""text"",""perSlide"":true}],
          ""colors"": [{""role"":""main"",""value"":""#1A2B4C"",""locked"":false}],
          ""fonts"": [{""role"":""heading"",""family"":""Pretendard""}]
        }";

        [Fact]
        public void Parse_Injects_Category_And_File_Not_From_Llm()
        {
            var u = UnderstandingParser.Parse(Sample, "표지", "표지_01.pptx");
            Assert.Equal("표지", u.Asset.Category);
            Assert.Equal("표지_01.pptx", u.Asset.File);
            Assert.Equal(2, u.Asset.SchemaVersion);
            Assert.Equal("slide", u.Asset.Scope);
        }

        [Fact]
        public void Parse_Maps_Core_Llm_Fields()
        {
            var u = UnderstandingParser.Parse(Sample, "표지", "표지_01.pptx");
            Assert.Equal("우측정렬 연도강조 표지", u.Asset.Name);
            Assert.Equal("layout", u.Asset.Kind);
            Assert.Equal("연도를 강조하는 표지가 필요할 때", u.Asset.UseWhen);
            Assert.Single(u.Asset.Slots);
            Assert.Equal("title", u.Asset.Slots[0].Name);
            Assert.Equal("#1A2B4C", u.Asset.Colors[0].Value);
            Assert.Equal("Pretendard", u.Asset.Fonts[0].Family);
        }

        [Fact]
        public void Parse_Extracts_Example_Intents_Separately()
        {
            var u = UnderstandingParser.Parse(Sample, "표지", "표지_01.pptx");
            Assert.Equal(2, u.ExampleIntents.Count);
            Assert.Contains("투자 유치 IR 표지", u.ExampleIntents);
        }

        [Fact]
        public void Parse_Missing_Arrays_Default_To_Empty()
        {
            var u = UnderstandingParser.Parse(@"{""name"":""x"",""kind"":""component""}", "표", "x.pptx");
            Assert.Empty(u.Asset.Tags);
            Assert.Empty(u.Asset.Slots);
            Assert.Empty(u.ExampleIntents);
        }
    }
}

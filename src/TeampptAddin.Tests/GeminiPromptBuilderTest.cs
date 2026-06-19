using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class GeminiPromptBuilderTest
    {
        private List<CatalogEntry> MakeCatalog()
        {
            return new List<CatalogEntry>
            {
                new CatalogEntry
                {
                    File = "header_1.pptx",
                    Name = "깔끔한 제목",
                    Kind = "component",
                    Category = "헤더",
                    Scope = "slide",
                    Tags = new List<string> { "심플", "제목" },
                    UseWhen = "간결한 제목 슬라이드가 필요할 때",
                    SlotNames = new List<string> { "title", "subtitle" },
                    ColorRoles = new List<string> { "main", "text" },
                    FontRoles = new List<string> { "heading", "body" }
                }
            };
        }

        private List<StylePalette> MakePalettes()
        {
            return new List<StylePalette>
            {
                new StylePalette
                {
                    Id = "blue-professional",
                    Name = "블루 프로페셔널",
                    Mood = new List<string> { "신뢰", "전문성" },
                    UseWhen = "B2B 제안서"
                }
            };
        }

        private List<StyleFont> MakeFonts()
        {
            return new List<StyleFont>
            {
                new StyleFont
                {
                    Name = "Pretendard",
                    Mood = new List<string> { "모던" },
                    UseWhen = "범용"
                }
            };
        }

        [Fact]
        public void SystemPrompt_Contains_Catalog_Json()
        {
            var prompt = GeminiPromptBuilder.BuildSystemPrompt(
                MakeCatalog(), MakePalettes(), MakeFonts());

            Assert.Contains("header_1.pptx", prompt);
            Assert.Contains("깔끔한 제목", prompt);
            Assert.Contains("blue-professional", prompt);
            Assert.Contains("Pretendard", prompt);
        }

        [Fact]
        public void SystemPrompt_Excludes_Hex_Values()
        {
            var palettes = MakePalettes();
            palettes[0].Colors = new PaletteColors
            {
                Main = "#2563EB", Sub1 = "#3B82F6",
                Sub2 = "#93C5FD", Text = "#1E293B"
            };

            var prompt = GeminiPromptBuilder.BuildSystemPrompt(
                MakeCatalog(), palettes, MakeFonts());

            Assert.DoesNotContain("#2563EB", prompt);
            Assert.DoesNotContain("#3B82F6", prompt);
        }

        [Fact]
        public void SystemPrompt_Contains_Json_Response_Schema()
        {
            var prompt = GeminiPromptBuilder.BuildSystemPrompt(
                MakeCatalog(), MakePalettes(), MakeFonts());

            Assert.Contains("\"file\"", prompt);
            Assert.Contains("\"reason\"", prompt);
            Assert.Contains("\"palette\"", prompt);
            Assert.Contains("\"font\"", prompt);
        }

        [Fact]
        public void UserPrompt_Contains_Intent()
        {
            var result = GeminiPromptBuilder.BuildUserPrompt("깔끔한 발표");
            Assert.Contains("깔끔한 발표", result);
        }

        [Fact]
        public void ResponseSchema_Has_Expected_Top_Level_Properties()
        {
            var schema = GeminiPromptBuilder.BuildResponseSchema();

            Assert.Equal("object", schema["type"]?.ToString());
            var props = schema["properties"];
            Assert.NotNull(props["message"]);
            Assert.NotNull(props["assets"]);
            Assert.NotNull(props["palette"]);
            Assert.NotNull(props["font"]);
        }

        [Fact]
        public void ResponseSchema_Assets_Is_Array_Of_File_Reason()
        {
            var schema = GeminiPromptBuilder.BuildResponseSchema();
            var assets = schema["properties"]["assets"];

            Assert.Equal("array", assets["type"]?.ToString());
            var itemProps = assets["items"]["properties"];
            Assert.NotNull(itemProps["file"]);
            Assert.NotNull(itemProps["reason"]);
        }

        [Fact]
        public void ResponseSchema_Palette_And_Font_Are_Nullable()
        {
            var schema = GeminiPromptBuilder.BuildResponseSchema();

            Assert.True(schema["properties"]["palette"]["nullable"].Value<bool>());
            Assert.True(schema["properties"]["font"]["nullable"].Value<bool>());
        }
    }
}

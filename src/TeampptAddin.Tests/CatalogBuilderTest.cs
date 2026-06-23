using System.Collections.Generic;
using Xunit;

namespace TeampptAddin.Tests
{
    public class CatalogBuilderTest
    {
        [Fact]
        public void Build_Projects_Compact_Entry_Without_Hex_Values()
        {
            var asset = new HeaderAsset
            {
                File = "header_3.pptx", Name = "장점 나열", Category = "헤더",
                Tags = new List<string> { "장점", "나열" },
                UseWhen = "장점 3개",
                Colors = new List<AssetColor> { new AssetColor { Role = "main", Value = "#2563EB" } },
                Fonts = new List<AssetFont> { new AssetFont { Role = "heading", Family = "Pretendard" } },
                Slots = new List<AssetSlot> { new AssetSlot { Name = "title" }, new AssetSlot { Name = "body" } }
            };

            var entries = CatalogBuilder.Build(new[] { asset });

            Assert.Single(entries);
            var e = entries[0];
            Assert.Equal("header_3.pptx", e.File);
        }
    }
}

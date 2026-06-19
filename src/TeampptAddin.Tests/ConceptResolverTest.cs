using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace TeampptAddin.Tests
{
    public class ConceptResolverTest
    {
        private static HeaderAsset Asset() => new HeaderAsset
        {
            Colors = new List<AssetColor>
            {
                new AssetColor { Role = "main",  Value = "#2563EB", Locked = false },
                new AssetColor { Role = "logo",  Value = "#FF0000", Locked = true },
                new AssetColor { Role = "text",  Value = "#1E293B", Locked = false }
            },
            Fonts = new List<AssetFont>
            {
                new AssetFont { Role = "heading", Family = "Pretendard" }
            }
        };

        private static DesignConcept Concept() => new DesignConcept
        {
            Colors = new Dictionary<string, string> { ["main"] = "#111827", ["text"] = "#374151" },
            Fonts = new Dictionary<string, string> { ["heading"] = "Noto Sans KR" }
        };

        [Fact]
        public void Unlocked_Role_With_Concept_Value_Is_Replaced()
        {
            var r = ConceptResolver.ResolveColors(Asset(), Concept());
            Assert.Equal("#111827", r.First(c => c.Role == "main").Value);
        }

        [Fact]
        public void Locked_Role_Keeps_Original()
        {
            var r = ConceptResolver.ResolveColors(Asset(), Concept());
            Assert.Equal("#FF0000", r.First(c => c.Role == "logo").Value);
        }

        [Fact]
        public void Role_Missing_From_Concept_Keeps_Original()
        {
            var concept = new DesignConcept { Colors = new Dictionary<string, string> { ["main"] = "#111827" } };
            var r = ConceptResolver.ResolveColors(Asset(), concept);
            Assert.Equal("#1E293B", r.First(c => c.Role == "text").Value);
        }

        [Fact]
        public void Null_Concept_Keeps_All_Original()
        {
            var r = ConceptResolver.ResolveColors(Asset(), null);
            Assert.Equal("#2563EB", r.First(c => c.Role == "main").Value);
        }

        [Fact]
        public void Fonts_Replaced_By_Concept_Family()
        {
            var r = ConceptResolver.ResolveFonts(Asset(), Concept());
            Assert.Equal("Noto Sans KR", r.First(f => f.Role == "heading").Family);
        }
    }
}

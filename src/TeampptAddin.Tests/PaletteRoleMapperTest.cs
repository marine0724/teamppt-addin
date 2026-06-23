using System.Collections.Generic;
using TeampptAddin;
using Xunit;

namespace TeampptAddin.Tests
{
    public class PaletteRoleMapperTest
    {
        private static AssetColor C(string role, string val) =>
            new AssetColor { Role = role, Value = val, Locked = false };

        [Fact]
        public void Null_Or_Empty_Returns_Null()
        {
            Assert.Null(PaletteRoleMapper.Map(null));
            Assert.Null(PaletteRoleMapper.Map(new List<AssetColor>()));
        }

        [Fact]
        public void Uses_Explicit_Roles_When_Present()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor>
            {
                C("main", "#2563EB"),
                C("text", "#1E293B"),
                C("background", "#FFFFFF"),
            });
            Assert.Equal("#2563EB", np.Main);
            Assert.Equal("#1E293B", np.Text);
            Assert.Equal("#FFFFFF", np.Background);
        }

        [Fact]
        public void Fills_Missing_Background_From_Main_Tint()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor> { C("main", "#2563EB") });
            Assert.Equal("#2563EB", np.Main);
            Assert.True(ColorHsl.FromHex(np.Background).L > 0.85);
        }

        [Fact]
        public void Fills_Missing_Text_Dark()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor> { C("main", "#2563EB") });
            Assert.True(ColorHsl.FromHex(np.Text).L < 0.3);
        }

        [Fact]
        public void Picks_Most_Saturated_As_Main_When_No_Role()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor>
            {
                C("", "#808080"),
                C("", "#2563EB"),
            });
            Assert.Equal("#2563EB", np.Main);
        }

        [Fact]
        public void Sub1_Sub2_Always_Populated()
        {
            var np = PaletteRoleMapper.Map(new List<AssetColor> { C("main", "#2563EB") });
            Assert.False(string.IsNullOrEmpty(np.Sub1));
            Assert.False(string.IsNullOrEmpty(np.Sub2));
        }
    }
}

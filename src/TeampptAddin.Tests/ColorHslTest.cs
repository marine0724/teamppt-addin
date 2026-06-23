using TeampptAddin;
using Xunit;

namespace TeampptAddin.Tests
{
    public class ColorHslTest
    {
        [Theory]
        [InlineData("#2563EB")]
        [InlineData("#1E293B")]
        [InlineData("#FFFFFF")]
        [InlineData("#000000")]
        public void FromHex_ToHex_RoundTrips(string hex)
        {
            var hsl = ColorHsl.FromHex(hex);
            Assert.Equal(hex, ColorHsl.ToHex(hsl));
        }

        [Fact]
        public void FromHex_Accepts_Lowercase_And_NoHash()
        {
            Assert.Equal("#2563EB", ColorHsl.ToHex(ColorHsl.FromHex("2563eb")));
        }

        [Fact]
        public void ContrastRatio_BlackWhite_IsMax()
        {
            Assert.True(ColorHsl.ContrastRatio("#000000", "#FFFFFF") > 20.9);
        }

        [Fact]
        public void ContrastRatio_SameColor_IsOne()
        {
            Assert.Equal(1.0, ColorHsl.ContrastRatio("#2563EB", "#2563EB"), 2);
        }

        [Fact]
        public void WithLightness_Sets_L_Keeps_Hue()
        {
            var hsl = ColorHsl.FromHex("#2563EB");
            var lighter = ColorHsl.WithLightness(hsl, 0.9);
            Assert.Equal(0.9, lighter.L, 3);
            Assert.Equal(hsl.H, lighter.H, 1);
        }

        [Fact]
        public void AdjustForContrast_Darkens_On_Light_Background()
        {
            var adjusted = ColorHsl.AdjustForContrast("#999999", "#FFFFFF", 4.5);
            Assert.True(ColorHsl.ContrastRatio(adjusted, "#FFFFFF") >= 4.5);
        }

        [Fact]
        public void AdjustForContrast_Lightens_On_Dark_Background()
        {
            var adjusted = ColorHsl.AdjustForContrast("#444444", "#0A1428", 4.5);
            Assert.True(ColorHsl.ContrastRatio(adjusted, "#0A1428") >= 4.5);
        }

        [Fact]
        public void AdjustForContrast_Returns_Original_When_Already_Sufficient()
        {
            var adjusted = ColorHsl.AdjustForContrast("#000000", "#FFFFFF", 4.5);
            Assert.Equal("#000000", adjusted);
        }
    }
}

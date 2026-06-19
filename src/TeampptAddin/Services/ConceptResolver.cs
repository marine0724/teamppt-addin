using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    public class ResolvedColor
    {
        public string Role { get; set; }
        public string Value { get; set; }
    }

    public class ResolvedFont
    {
        public string Role { get; set; }
        public string Family { get; set; }
    }

    public static class ConceptResolver
    {
        public static List<ResolvedColor> ResolveColors(HeaderAsset asset, DesignConcept concept)
        {
            var colors = asset?.Colors ?? new List<AssetColor>();
            return colors.Select(c =>
            {
                var value = c.Value;
                if (!c.Locked && concept?.Colors != null &&
                    concept.Colors.TryGetValue(c.Role, out var conceptValue))
                {
                    value = conceptValue;
                }
                return new ResolvedColor { Role = c.Role, Value = value };
            }).ToList();
        }

        public static List<ResolvedFont> ResolveFonts(HeaderAsset asset, DesignConcept concept)
        {
            var fonts = asset?.Fonts ?? new List<AssetFont>();
            return fonts.Select(f =>
            {
                var family = f.Family;
                if (concept?.Fonts != null &&
                    concept.Fonts.TryGetValue(f.Role, out var conceptFamily))
                {
                    family = conceptFamily;
                }
                return new ResolvedFont { Role = f.Role, Family = family };
            }).ToList();
        }
    }
}

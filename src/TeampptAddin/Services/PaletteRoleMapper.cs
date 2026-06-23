using System;
using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    public static class PaletteRoleMapper
    {
        public static NormalizedPalette Map(List<AssetColor> colors)
        {
            var valid = (colors ?? new List<AssetColor>())
                .Where(c => c != null && !string.IsNullOrWhiteSpace(c.Value))
                .ToList();
            if (valid.Count == 0) return null;

            string main = FindByRole(valid, "main", "primary")
                          ?? MostSaturated(valid)
                          ?? valid[0].Value;

            string text = FindByRole(valid, "text", "foreground")
                          ?? Darkest(valid)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main), 0.15));

            string background = FindByRole(valid, "background", "bg")
                          ?? Lightest(valid, 0.85)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main), 0.96));

            var subCandidates = valid
                .Where(c => RoleContains(c.Role, "sub", "accent", "secondary"))
                .Select(c => c.Value)
                .ToList();

            string sub1 = subCandidates.ElementAtOrDefault(0)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main),
                                Clamp01(ColorHsl.FromHex(main).L + 0.18)));
            string sub2 = subCandidates.ElementAtOrDefault(1)
                          ?? ColorHsl.ToHex(ColorHsl.WithLightness(ColorHsl.FromHex(main),
                                Clamp01(ColorHsl.FromHex(main).L - 0.18)));

            return new NormalizedPalette
            {
                Background = background,
                Main = main,
                Sub1 = sub1,
                Sub2 = sub2,
                Text = text
            };
        }

        private static string FindByRole(List<AssetColor> colors, params string[] keys)
        {
            var hit = colors.FirstOrDefault(c => RoleContains(c.Role, keys));
            return hit?.Value;
        }

        private static bool RoleContains(string role, params string[] keys)
        {
            if (string.IsNullOrEmpty(role)) return false;
            var r = role.ToLowerInvariant();
            return keys.Any(k => r.Contains(k));
        }

        private static string MostSaturated(List<AssetColor> colors)
        {
            return colors
                .OrderByDescending(c => ColorHsl.FromHex(c.Value).S)
                .FirstOrDefault()?.Value;
        }

        private static string Darkest(List<AssetColor> colors)
        {
            var ordered = colors.OrderBy(c => ColorHsl.FromHex(c.Value).L).ToList();
            if (ordered.Count > 0 && ColorHsl.FromHex(ordered[0].Value).L < 0.3)
                return ordered[0].Value;
            return null;
        }

        private static string Lightest(List<AssetColor> colors, double minL)
        {
            var hit = colors
                .OrderByDescending(c => ColorHsl.FromHex(c.Value).L)
                .FirstOrDefault(c => ColorHsl.FromHex(c.Value).L >= minL);
            return hit?.Value;
        }

        private static double Clamp01(double v) => v < 0 ? 0 : (v > 1 ? 1 : v);
    }
}

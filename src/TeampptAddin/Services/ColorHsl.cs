using System;

namespace TeampptAddin
{
    public struct Hsl
    {
        public double H;
        public double S;
        public double L;
        public Hsl(double h, double s, double l) { H = h; S = s; L = l; }
    }

    public static class ColorHsl
    {
        public static Hsl FromHex(string hex)
        {
            var (r, g, b) = HexToRgb(hex);
            double rd = r / 255.0, gd = g / 255.0, bd = b / 255.0;
            double max = Math.Max(rd, Math.Max(gd, bd));
            double min = Math.Min(rd, Math.Min(gd, bd));
            double h = 0, s, l = (max + min) / 2.0;
            double d = max - min;
            if (d == 0) { s = 0; }
            else
            {
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
                if (max == rd) h = (gd - bd) / d + (gd < bd ? 6 : 0);
                else if (max == gd) h = (bd - rd) / d + 2;
                else h = (rd - gd) / d + 4;
                h *= 60;
            }
            return new Hsl(h, s, l);
        }

        public static string ToHex(Hsl hsl)
        {
            double h = hsl.H, s = Clamp01(hsl.S), l = Clamp01(hsl.L);
            double r, g, b;
            if (s == 0) { r = g = b = l; }
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                double p = 2 * l - q;
                double hk = h / 360.0;
                r = HueToRgb(p, q, hk + 1.0 / 3.0);
                g = HueToRgb(p, q, hk);
                b = HueToRgb(p, q, hk - 1.0 / 3.0);
            }
            return "#" + ToByteHex(r) + ToByteHex(g) + ToByteHex(b);
        }

        public static Hsl WithLightness(Hsl hsl, double l)
        {
            return new Hsl(hsl.H, hsl.S, Clamp01(l));
        }

        public static double ContrastRatio(string hexA, string hexB)
        {
            double la = RelativeLuminance(hexA);
            double lb = RelativeLuminance(hexB);
            double lighter = Math.Max(la, lb);
            double darker = Math.Min(la, lb);
            return (lighter + 0.05) / (darker + 0.05);
        }

        public static string AdjustForContrast(string fgHex, string bgHex, double targetRatio)
        {
            if (ContrastRatio(fgHex, bgHex) >= targetRatio) return fgHex;
            var fg = FromHex(fgHex);
            bool bgIsLight = RelativeLuminance(bgHex) > 0.5;
            for (int i = 0; i <= 50; i++)
            {
                double l = bgIsLight ? fg.L - i * 0.02 : fg.L + i * 0.02;
                if (l < 0 || l > 1) break;
                var candidate = ToHex(WithLightness(fg, l));
                if (ContrastRatio(candidate, bgHex) >= targetRatio) return candidate;
            }
            return bgIsLight ? "#000000" : "#FFFFFF";
        }

        private static double RelativeLuminance(string hex)
        {
            var (r, g, b) = HexToRgb(hex);
            double rs = LinearChannel(r / 255.0);
            double gs = LinearChannel(g / 255.0);
            double bs = LinearChannel(b / 255.0);
            return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;
        }

        private static double LinearChannel(double c)
        {
            return c <= 0.03928 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6.0) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2.0) return q;
            if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6;
            return p;
        }

        private static (int, int, int) HexToRgb(string hex)
        {
            hex = (hex ?? "").Trim().TrimStart('#');
            if (hex.Length != 6) throw new ArgumentException($"잘못된 hex: {hex}");
            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return (r, g, b);
        }

        private static string ToByteHex(double channel01)
        {
            int v = (int)Math.Round(Clamp01(channel01) * 255.0);
            return v.ToString("X2");
        }

        private static double Clamp01(double v) => v < 0 ? 0 : (v > 1 ? 1 : v);
    }
}

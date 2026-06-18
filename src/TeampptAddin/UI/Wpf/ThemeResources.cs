using System.Windows;
using System.Windows.Media;

namespace TeampptAddin
{
    /// <summary>
    /// Runtime mirror of Views/Styles/Theme.xaml.
    /// Keep both files in sync when updating design tokens.
    /// </summary>
    internal static class ThemeResources
    {
        // ── Backgrounds ─────────────────────────────────────────────
        public static readonly SolidColorBrush BgBase           = F(0xFF, 0xFF, 0xFF);
        public static readonly SolidColorBrush BgSurface        = F(0xFA, 0xFA, 0xFA);
        public static readonly SolidColorBrush BgCard           = F(0xFF, 0xFF, 0xFF);
        public static readonly SolidColorBrush BgCardHover      = F(0xF8, 0xF9, 0xFF);
        public static readonly SolidColorBrush BgThumb          = F(0xF8, 0xF9, 0xFA);
        public static readonly SolidColorBrush BgInput          = F(0xF4, 0xF5, 0xFE);
        public static readonly SolidColorBrush BgChip           = F(0xF8, 0xF9, 0xFA);
        public static readonly SolidColorBrush BgCategoryActive = F(0xEE, 0xF0, 0xFE);
        public static readonly SolidColorBrush BgBadge          = F(0xEE, 0xF0, 0xFE);
        public static readonly SolidColorBrush BgAiResponse     = F(0xF8, 0xF9, 0xFA);
        public static readonly SolidColorBrush BgUserBubble     = F(0x4F, 0x5C, 0xF5);

        // ── Accent ──────────────────────────────────────────────────
        public static readonly SolidColorBrush Accent           = F(0x4F, 0x5C, 0xF5);
        public static readonly SolidColorBrush AccentBorder     = F(0xC7, 0xCA, 0xFC);

        // ── Text ────────────────────────────────────────────────────
        public static readonly SolidColorBrush TextMain         = F(0x19, 0x1F, 0x28);
        public static readonly SolidColorBrush TextSub          = F(0x8B, 0x95, 0xA1);
        public static readonly SolidColorBrush TextDim          = F(0xB0, 0xB8, 0xC1);
        public static readonly SolidColorBrush TextAccent       = F(0x4F, 0x5C, 0xF5);

        // ── Border ──────────────────────────────────────────────────
        public static readonly SolidColorBrush BorderBase       = F(0xF2, 0xF4, 0xF6);
        public static readonly SolidColorBrush BorderCard       = F(0xE5, 0xE8, 0xEB);
        public static readonly SolidColorBrush BorderCardHover  = F(0x4F, 0x5C, 0xF5);

        // ── Status ──────────────────────────────────────────────────
        public static readonly SolidColorBrush StatusSuccess    = F(0x10, 0xB9, 0x81);
        public static readonly SolidColorBrush StatusError      = F(0xF8, 0x71, 0x71);

        // ── Typography ──────────────────────────────────────────────
        public static readonly FontFamily FontBase =
            new FontFamily("Pretendard, Segoe UI");

        // ── Corner Radius ────────────────────────────────────────────
        public static readonly CornerRadius RadiusPane    = new CornerRadius(18);
        public static readonly CornerRadius RadiusCard    = new CornerRadius(13);
        public static readonly CornerRadius RadiusInput   = new CornerRadius(14);
        public static readonly CornerRadius RadiusBadge   = new CornerRadius(5);
        public static readonly CornerRadius RadiusChip    = new CornerRadius(99);
        public static readonly CornerRadius RadiusBubble  = new CornerRadius(14);

        public static void ApplyRoundedClip(FrameworkElement element, double radius)
        {
            ApplyRoundedClip(element, new CornerRadius(radius));
        }

        public static void ApplyRoundedClip(FrameworkElement element, CornerRadius radius)
        {
            element.SizeChanged += (s, e) =>
            {
                element.Clip = BuildRoundedGeometry(
                    element.ActualWidth, element.ActualHeight, radius);
            };
        }

        private static Geometry BuildRoundedGeometry(double w, double h, CornerRadius r)
        {
            if (r.TopLeft == r.TopRight && r.TopRight == r.BottomRight && r.BottomRight == r.BottomLeft)
            {
                var rg = new RectangleGeometry(new Rect(0, 0, w, h), r.TopLeft, r.TopLeft);
                rg.Freeze();
                return rg;
            }

            double tl = r.TopLeft, tr = r.TopRight;
            double br = r.BottomRight, bl = r.BottomLeft;

            var geo = new StreamGeometry();
            using (var ctx = geo.Open())
            {
                ctx.BeginFigure(new Point(tl, 0), true, true);
                ctx.LineTo(new Point(w - tr, 0), true, false);
                if (tr > 0)
                    ctx.ArcTo(new Point(w, tr), new Size(tr, tr), 0, false, SweepDirection.Clockwise, true, false);
                ctx.LineTo(new Point(w, h - br), true, false);
                if (br > 0)
                    ctx.ArcTo(new Point(w - br, h), new Size(br, br), 0, false, SweepDirection.Clockwise, true, false);
                ctx.LineTo(new Point(bl, h), true, false);
                if (bl > 0)
                    ctx.ArcTo(new Point(0, h - bl), new Size(bl, bl), 0, false, SweepDirection.Clockwise, true, false);
                ctx.LineTo(new Point(0, tl), true, false);
                if (tl > 0)
                    ctx.ArcTo(new Point(tl, 0), new Size(tl, tl), 0, false, SweepDirection.Clockwise, true, false);
            }
            geo.Freeze();
            return geo;
        }

        private static SolidColorBrush F(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }
    }
}

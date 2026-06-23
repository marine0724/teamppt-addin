using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public static class SlideStyleApplier
    {
        private static readonly HashSet<string> _installedFonts = LoadInstalledFonts();

        private static HashSet<string> LoadInstalledFonts()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var ifc = new InstalledFontCollection())
            {
                foreach (var f in ifc.Families)
                    set.Add(f.Name);
            }
            Logger.Log($"[StyleApply] InstalledFonts loaded: {set.Count} families");
            return set;
        }

        internal static string ResolveFontName(StyleFont font)
        {
            if (font == null || string.IsNullOrEmpty(font.Name)) return null;
            if (_installedFonts.Contains(font.Name)) return font.Name;
            Logger.Log($"[StyleApply] font '{font.Name}' not installed, trying fallback '{font.Fallback}'");
            if (!string.IsNullOrEmpty(font.Fallback) && _installedFonts.Contains(font.Fallback))
                return font.Fallback;
            Logger.Log($"[StyleApply] fallback '{font.Fallback}' also not installed, skipping font");
            return null;
        }

        public static void Apply(PowerPoint.Slide slide, StylePalette palette, StyleFont font)
        {
            var resolvedFontName = ResolveFontName(font);
            Logger.Log($"[StyleApply] palette={palette?.Name ?? "NULL"}, font={font?.Name ?? "NULL"}, resolved={resolvedFontName ?? "NULL"}");
            if (slide == null) return;
            var c = palette?.Colors;

            if (c != null && !string.IsNullOrEmpty(c.Background))
            {
                try
                {
                    slide.FollowMasterBackground = MsoTriState.msoFalse;
                    slide.Background.Fill.Visible = MsoTriState.msoTrue;
                    slide.Background.Fill.Solid();
                    slide.Background.Fill.ForeColor.RGB = Ole(c.Background);
                }
                catch { }
            }

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                try
                {
                    Logger.Log($"[StyleApply] shape={shape.Name}, type={shape.Type}, protected={IsProtected(shape)}, hasText={shape.HasTextFrame}");
                    if (IsProtected(shape)) continue;

                    if (c != null && !string.IsNullOrEmpty(c.Main))
                    {
                        try
                        {
                            if (shape.Fill.Visible == MsoTriState.msoTrue)
                                shape.Fill.ForeColor.RGB = Ole(c.Main);
                        }
                        catch { }
                    }

                    if (c != null && !string.IsNullOrEmpty(c.Sub1))
                    {
                        try
                        {
                            if (shape.Line.Visible == MsoTriState.msoTrue)
                                shape.Line.ForeColor.RGB = Ole(c.Sub1);
                        }
                        catch { }
                    }

                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        var tr = shape.TextFrame.TextRange;
                        int count = tr.Paragraphs().Count;
                        Logger.Log($"[StyleApply]   textParagraphs={count}, fontToApply={resolvedFontName ?? "NULL"}");
                        for (int i = 1; i <= count; i++)
                        {
                            var para = tr.Paragraphs(i);
                            if (!string.IsNullOrEmpty(resolvedFontName))
                            {
                                Logger.Log($"[StyleApply]   para[{i}] before={para.Font.Name}, setting={resolvedFontName}");
                                para.Font.Name = resolvedFontName;
                                Logger.Log($"[StyleApply]   para[{i}] after={para.Font.Name}");
                            }
                            if (c != null && !string.IsNullOrEmpty(c.Text))
                                para.Font.Color.RGB = Ole(c.Text);
                        }
                    }
                }
                catch { }
            }
        }

        private static bool IsProtected(PowerPoint.Shape shape)
        {
            switch (shape.Type)
            {
                case MsoShapeType.msoPicture:
                case MsoShapeType.msoPlaceholder:
                case MsoShapeType.msoMedia:
                case MsoShapeType.msoOLEControlObject:
                case MsoShapeType.msoEmbeddedOLEObject:
                    return true;
                default:
                    return false;
            }
        }

        private static int Ole(string hex)
        {
            var h = ColorHsl.ToHex(ColorHsl.FromHex(hex));
            int r = Convert.ToInt32(h.Substring(1, 2), 16);
            int g = Convert.ToInt32(h.Substring(3, 2), 16);
            int b = Convert.ToInt32(h.Substring(5, 2), 16);
            return ColorTranslator.ToOle(Color.FromArgb(r, g, b));
        }
    }
}

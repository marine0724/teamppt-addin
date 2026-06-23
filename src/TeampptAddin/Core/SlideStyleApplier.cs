using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public static class SlideStyleApplier
    {
        public static void Apply(PowerPoint.Slide slide, StylePalette palette, StyleFont font)
        {
            if (slide == null || palette?.Colors == null) return;
            var c = palette.Colors;

            if (!string.IsNullOrEmpty(c.Background))
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
                    if (IsProtected(shape)) continue;

                    if (!string.IsNullOrEmpty(c.Main))
                    {
                        try
                        {
                            if (shape.Fill.Visible == MsoTriState.msoTrue)
                                shape.Fill.ForeColor.RGB = Ole(c.Main);
                        }
                        catch { }
                    }

                    if (!string.IsNullOrEmpty(c.Sub1))
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
                        for (int i = 1; i <= count; i++)
                        {
                            var para = tr.Paragraphs(i);
                            if (font != null && !string.IsNullOrEmpty(font.Name))
                                para.Font.Name = font.Name;
                            if (!string.IsNullOrEmpty(c.Text))
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

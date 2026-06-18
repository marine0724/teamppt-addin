using System;
using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public static class CoordinateConverter
    {
        public static PointF? ScreenToSlide(Point screenPos, PowerPoint.DocumentWindow window)
        {
            float refPt = 500f;
            int ox = window.PointsToScreenPixelsX(0f);
            int oy = window.PointsToScreenPixelsY(0f);
            int rx = window.PointsToScreenPixelsX(refPt);
            int ry = window.PointsToScreenPixelsY(refPt);

            float scaleX = (rx - ox) / refPt;
            float scaleY = (ry - oy) / refPt;

            Logger.Log($"Coord: o=({ox},{oy}) r=({rx},{ry}) scale=({scaleX:F3},{scaleY:F3}) mouse=({screenPos.X},{screenPos.Y})");

            if (Math.Abs(scaleX) < 0.01f || Math.Abs(scaleY) < 0.01f)
                return null;

            float slideX = (screenPos.X - ox) / scaleX;
            float slideY = (screenPos.Y - oy) / scaleY;
            Logger.Log($"Raw slide: ({slideX:F1},{slideY:F1})");

            float slideW = window.Presentation.PageSetup.SlideWidth;
            float slideH = window.Presentation.PageSetup.SlideHeight;

            slideX = Math.Max(0, Math.Min(slideX, slideW));
            slideY = Math.Max(0, Math.Min(slideY, slideH));

            return new PointF(slideX, slideY);
        }

        public static void CenterShapesAt(PowerPoint.ShapeRange shapes, PointF slidePos)
        {
            int cnt = shapes.Count;
            Logger.Log($"Shape count: {cnt}");
            if (cnt == 0) return;

            float minL = float.MaxValue, minT = float.MaxValue;
            float maxR = float.MinValue, maxB = float.MinValue;
            for (int i = 1; i <= cnt; i++)
            {
                var s = shapes[i];
                float l = s.Left, t = s.Top, w = s.Width, h = s.Height;
                Logger.Log($"  shape[{i}]: L={l:F1} T={t:F1} W={w:F1} H={h:F1}");
                if (l < minL) minL = l;
                if (t < minT) minT = t;
                if (l + w > maxR) maxR = l + w;
                if (t + h > maxB) maxB = t + h;
            }

            float cx = (minL + maxR) / 2f;
            float cy = (minT + maxB) / 2f;
            float dx = slidePos.X - cx;
            float dy = slidePos.Y - cy;
            Logger.Log($"Center: ({cx:F1},{cy:F1}), delta: ({dx:F1},{dy:F1})");

            for (int i = 1; i <= cnt; i++)
            {
                shapes[i].Left += dx;
                shapes[i].Top += dy;
            }

            Logger.Log($"Positioned at slide({slidePos.X:F0},{slidePos.Y:F0})");
        }

        public static void PositionShapesAtCursor(
            PowerPoint.ShapeRange shapes, Point screenPos, PowerPoint.DocumentWindow window)
        {
            try
            {
                var slidePos = ScreenToSlide(screenPos, window);
                if (!slidePos.HasValue) return;
                CenterShapesAt(shapes, slidePos.Value);
            }
            catch (Exception ex)
            {
                Logger.Log($"PositionShapes (non-fatal): {ex.Message}");
            }
        }
    }
}

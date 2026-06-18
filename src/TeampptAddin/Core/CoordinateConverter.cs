using System;
using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 스크린 픽셀 좌표 ↔ PowerPoint 슬라이드 포인트 좌표 변환기.
    ///
    /// 변환 원리:
    /// PowerPoint의 PointsToScreenPixelsX/Y를 사용하여 슬라이드 원점(0,0)과
    /// 참조점(500,500)의 스크린 좌표를 얻고, 그 차이로 scale factor를 계산.
    /// 마우스 스크린 좌표에서 원점을 빼고 scale로 나누면 슬라이드 좌표가 됨.
    ///
    /// 사용처: 드래그앤드롭 시 마우스 드롭 위치를 슬라이드 좌표로 변환하여
    /// 붙여넣은 Shape의 중심을 해당 위치로 이동.
    /// </summary>
    public static class CoordinateConverter
    {
        /// <summary>
        /// 스크린 픽셀 좌표 → 슬라이드 포인트 좌표 변환.
        /// 슬라이드 범위(0~slideW, 0~slideH)로 클램핑됨.
        /// scale이 너무 작으면(줌 문제) null 반환.
        /// </summary>
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

        /// <summary>
        /// ShapeRange의 바운딩 박스 중심을 slidePos로 이동.
        /// 전체 Shape 그룹의 min/max를 계산하여 중심점을 구하고, 그 delta만큼 이동.
        /// </summary>
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

        /// <summary>
        /// 통합 메서드: 스크린 좌표 → 슬라이드 좌표 변환 + Shape 중심 이동.
        /// DragHandler.EndDrag에서 호출됨. 예외 발생 시 로그만 남기고 무시(non-fatal).
        /// </summary>
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

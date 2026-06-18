using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace TeampptAddin
{
    internal class CardControl : Control
    {
        private const int ThumbH = 120;
        private const int LabelH = 30;
        private const int Radius = 10;

        static readonly Color CardBg = Color.FromArgb(39, 39, 42);
        static readonly Color CardHover = Color.FromArgb(52, 52, 56);
        static readonly Color CardDrag = Color.FromArgb(45, 45, 50);
        static readonly Color Accent = Color.FromArgb(99, 102, 241);
        static readonly Color TextMain = Color.FromArgb(228, 228, 231);
        static readonly Color TextSub = Color.FromArgb(113, 113, 122);
        static readonly Color BorderNormal = Color.FromArgb(50, 50, 55);
        static readonly Color ThumbBg = Color.FromArgb(28, 28, 32);

        static readonly Font TitleFont = new Font("Segoe UI", 10f, FontStyle.Bold);
        static readonly Font BadgeFont = new Font("Segoe UI", 7f, FontStyle.Bold);
        static readonly Font PlaceholderFont = new Font("Segoe UI", 13f, FontStyle.Bold);

        private readonly Image _thumb;
        private readonly string _title;
        private readonly DragHandler _drag;
        private bool _hovered;

        public CardControl(Image thumb, string title, string pptxPath,
            Action<string, Color> setStatus, Action resetStatus, Func<Rectangle> getHostBounds)
        {
            _thumb = thumb;
            _title = title;
            Height = ThumbH + LabelH;
            Cursor = Cursors.Hand;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint
                   | ControlStyles.DoubleBuffer | ControlStyles.ResizeRedraw, true);

            _drag = new DragHandler(this, pptxPath, title, thumb,
                setStatus, resetStatus, getHostBounds);
        }

        #region Painting

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = MakeRoundRect(rect, Radius))
            {
                using (var brush = new SolidBrush(
                    _drag.IsDragging ? CardDrag : (_hovered ? CardHover : CardBg)))
                    g.FillPath(brush, path);

                g.SetClip(path);

                using (var b = new SolidBrush(ThumbBg))
                    g.FillRectangle(b, 0, 0, Width, ThumbH);

                if (_thumb != null)
                {
                    float scaleW = (float)(Width - 16) / _thumb.Width;
                    float scaleH = (float)(ThumbH - 16) / _thumb.Height;
                    float scale = Math.Min(scaleW, scaleH);
                    int dw = (int)(_thumb.Width * scale);
                    int dh = (int)(_thumb.Height * scale);
                    g.DrawImage(_thumb, (Width - dw) / 2, (ThumbH - dh) / 2, dw, dh);
                }
                else
                {
                    TextRenderer.DrawText(g, _title, PlaceholderFont,
                        new Rectangle(0, 0, Width, ThumbH), TextSub,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }

                using (var pen = new Pen(BorderNormal))
                    g.DrawLine(pen, 0, ThumbH, Width, ThumbH);

                g.ResetClip();

                int ty = ThumbH + (LabelH - 16) / 2;

                TextRenderer.DrawText(g, _title, TitleFont,
                    new Rectangle(12, ty, Width - 70, 16), TextMain,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                var badge = "DRAG";
                var bs = TextRenderer.MeasureText(badge, BadgeFont);
                var br = new Rectangle(Width - bs.Width - 10, ty, bs.Width + 6, 16);
                using (var bp = MakeRoundRect(br, 3))
                using (var bb = new SolidBrush(Color.FromArgb(30, 99, 102, 241)))
                    g.FillPath(bb, bp);
                TextRenderer.DrawText(g, badge, BadgeFont, br, Accent,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

                using (var pen = new Pen(
                    _drag.IsDragging ? Accent : (_hovered ? Accent : BorderNormal),
                    (_hovered || _drag.IsDragging) ? 1.5f : 0.5f))
                    g.DrawPath(pen, path);
            }
        }

        private static GraphicsPath MakeRoundRect(Rectangle r, int rad)
        {
            int d = rad * 2;
            var gp = new GraphicsPath();
            gp.AddArc(r.X, r.Y, d, d, 180, 90);
            gp.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            gp.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            gp.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            gp.CloseFigure();
            return gp;
        }

        #endregion

        #region Mouse Events

        protected override void OnMouseEnter(EventArgs e)
        {
            _hovered = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            if (!_drag.IsDragging) { _hovered = false; Invalidate(); }
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            _drag.HandleMouseDown(e);
            base.OnMouseDown(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            _drag.HandleMouseMove(e);
            base.OnMouseMove(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            _drag.HandleMouseUp(e);
            _hovered = ClientRectangle.Contains(PointToClient(Cursor.Position));
            Cursor = Cursors.Hand;
            Invalidate();
            base.OnMouseUp(e);
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            _drag.HandleCaptureChanged();
            Cursor = Cursors.Hand;
            Invalidate();
            base.OnMouseCaptureChanged(e);
        }

        #endregion
    }
}

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    [ComVisible(true)]
    [Guid("2D4E6F8A-1B3C-5D7E-9F0A-4C6E8D2B1A3F")]
    [ProgId("TeampptAddin.TaskPaneHost")]
    public class TaskPaneHost : UserControl, IObjectSafety
    {
        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;

        static readonly Color BgColor = Color.FromArgb(24, 24, 27);
        static readonly Color HeaderBg = Color.FromArgb(30, 30, 34);
        static readonly Color AccentColor = Color.FromArgb(99, 102, 241);
        static readonly Color TextDim = Color.FromArgb(113, 113, 122);

        internal Label StatusLabel;
        private Panel _scrollPanel;
        private bool _loaded;
        private int _assetCount;

        public TaskPaneHost()
        {
            try
            {
                LogDebug($"Constructor. STA={Thread.CurrentThread.GetApartmentState()}");
                InitUI();
            }
            catch (Exception ex)
            {
                LogDebug($"Constructor FAILED: {ex}");
            }
        }

        private void InitUI()
        {
            BackColor = BgColor;

            var header = new Panel { Dock = DockStyle.Top, Height = 52, BackColor = HeaderBg };
            header.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(50, 50, 55)))
                    e.Graphics.DrawLine(pen, 0, 51, header.Width, 51);
            };
            header.Controls.Add(new Label
            {
                Text = "TEAMPPT",
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = AccentColor,
                Location = new Point(16, 6),
                AutoSize = true
            });
            header.Controls.Add(new Label
            {
                Text = "헤더 에셋",
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextDim,
                Location = new Point(16, 30),
                AutoSize = true
            });

            StatusLabel = new Label
            {
                Text = "로딩 중...",
                ForeColor = TextDim,
                Font = new Font("Segoe UI", 9f),
                Dock = DockStyle.Bottom,
                Height = 28,
                Padding = new Padding(14, 6, 0, 0),
                BackColor = HeaderBg
            };

            _scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = BgColor
            };

            Controls.Add(_scrollPanel);
            Controls.Add(StatusLabel);
            Controls.Add(header);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (!_loaded && Width > 0 && _scrollPanel != null)
            {
                _loaded = true;
                LoadCards();
            }
        }

        private void LoadCards()
        {
            var assetsDir = Globals.AssetsDir;
            var thumbDir = Globals.ThumbnailDir;
            int y = 10;

            for (int i = 1; i <= 7; i++)
            {
                var pptxPath = Path.Combine(assetsDir, $"header_{i}.pptx");
                if (!File.Exists(pptxPath)) continue;

                var thumbPath = Path.Combine(thumbDir, $"header_{i}.png");
                var thumb = LoadThumbnail(pptxPath, thumbPath);

                var card = new CardControl(thumb, $"Header {i}", pptxPath, this);
                card.Location = new Point(10, y);
                card.Width = _scrollPanel.ClientSize.Width - 20 - SystemInformation.VerticalScrollBarWidth;
                card.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
                _scrollPanel.Controls.Add(card);

                y += card.Height + 10;
                _assetCount++;
            }

            ResetStatus();
        }

        internal void ResetStatus()
        {
            StatusLabel.Text = _assetCount > 0
                ? $"{_assetCount}개 에셋 · 클릭 또는 드래그하여 삽입"
                : "Assets 폴더에 header_N.pptx 파일을 넣으세요";
            StatusLabel.ForeColor = TextDim;
        }

        #region Thumbnail Loading

        private Image LoadThumbnail(string pptxPath, string cachePath)
        {
            if (File.Exists(cachePath))
            {
                try
                {
                    if (File.GetLastWriteTime(cachePath) >= File.GetLastWriteTime(pptxPath))
                        return LoadImageNoLock(cachePath);
                    File.Delete(cachePath);
                }
                catch { }
            }

            try
            {
                ShapeInserter.GenerateThumbnail(pptxPath, cachePath);
                if (File.Exists(cachePath))
                    return LoadImageNoLock(cachePath);
            }
            catch (Exception ex)
            {
                LogDebug($"COM thumb fail [{Path.GetFileName(pptxPath)}]: {ex.Message}");
            }

            try
            {
                using (var zip = ZipFile.OpenRead(pptxPath))
                {
                    foreach (var entry in zip.Entries)
                    {
                        if (!entry.FullName.StartsWith("docProps/thumbnail", StringComparison.OrdinalIgnoreCase))
                            continue;
                        using (var stream = entry.Open())
                        using (var ms = new MemoryStream())
                        {
                            stream.CopyTo(ms);
                            return Image.FromStream(new MemoryStream(ms.ToArray()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"ZIP thumb fail [{Path.GetFileName(pptxPath)}]: {ex.Message}");
            }

            return null;
        }

        private static Image LoadImageNoLock(string path)
        {
            return Image.FromStream(new MemoryStream(File.ReadAllBytes(path)));
        }

        #endregion

        #region Logging

        internal static void LogDebug(string msg)
        {
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TeampptAddin");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "debug.log"), $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\r\n");
            }
            catch { }
        }

        #endregion

        #region IObjectSafety

        public int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            return S_OK;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            return S_OK;
        }

        #endregion

        #region GhostWindow — per-pixel alpha, actual size

        private class GhostWindow : Form
        {
            [DllImport("user32.dll", SetLastError = true)]
            static extern bool UpdateLayeredWindow(
                IntPtr hwnd, IntPtr hdcDst,
                ref POINT pptDst, ref SIZE psize,
                IntPtr hdcSrc, ref POINT pptSrc,
                int crKey, ref BLENDFUNCTION pblend, int dwFlags);

            [DllImport("user32.dll")]
            static extern IntPtr GetDC(IntPtr hWnd);

            [DllImport("user32.dll")]
            static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

            [DllImport("gdi32.dll")]
            static extern IntPtr CreateCompatibleDC(IntPtr hDC);

            [DllImport("gdi32.dll")]
            static extern bool DeleteDC(IntPtr hDC);

            [DllImport("gdi32.dll")]
            static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObj);

            [DllImport("gdi32.dll")]
            static extern bool DeleteObject(IntPtr hObj);

            [StructLayout(LayoutKind.Sequential)]
            struct POINT { public int X, Y; }

            [StructLayout(LayoutKind.Sequential)]
            struct SIZE { public int cx, cy; }

            [StructLayout(LayoutKind.Sequential, Pack = 1)]
            struct BLENDFUNCTION
            {
                public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat;
            }

            const int ULW_ALPHA = 2;
            const int WS_EX_LAYERED = 0x00080000;
            const int WS_EX_TOOLWINDOW = 0x00000080;
            const int WS_EX_TRANSPARENT = 0x00000020;
            const int WS_EX_NOACTIVATE = 0x08000000;
            const int WS_EX_TOPMOST = 0x00000008;

            private Bitmap _alphaBmp;

            public GhostWindow(Image thumb)
            {
                FormBorderStyle = FormBorderStyle.None;
                ShowInTaskbar = false;
                TopMost = true;
                StartPosition = FormStartPosition.Manual;

                _alphaBmp = BuildAlphaBitmap(thumb);
                Size = _alphaBmp.Size;
            }

            protected override bool ShowWithoutActivation => true;

            protected override CreateParams CreateParams
            {
                get
                {
                    var cp = base.CreateParams;
                    cp.ExStyle |= WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT
                                | WS_EX_NOACTIVATE | WS_EX_TOPMOST;
                    return cp;
                }
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);
                ApplyAlpha();
            }

            public void MoveTo(Point screenPos)
            {
                Location = new Point(
                    screenPos.X - Width / 2,
                    screenPos.Y - Height / 2);
            }

            private void ApplyAlpha()
            {
                if (_alphaBmp == null || !IsHandleCreated) return;

                var screenDC = GetDC(IntPtr.Zero);
                var memDC = CreateCompatibleDC(screenDC);
                var hBmp = _alphaBmp.GetHbitmap(Color.FromArgb(0, 0, 0, 0));
                var prev = SelectObject(memDC, hBmp);

                var ptDst = new POINT { X = Left, Y = Top };
                var ptSrc = new POINT { X = 0, Y = 0 };
                var sz = new SIZE { cx = _alphaBmp.Width, cy = _alphaBmp.Height };
                var blend = new BLENDFUNCTION
                {
                    BlendOp = 0,
                    BlendFlags = 0,
                    SourceConstantAlpha = 180,
                    AlphaFormat = 1
                };

                UpdateLayeredWindow(Handle, screenDC, ref ptDst, ref sz,
                    memDC, ref ptSrc, 0, ref blend, ULW_ALPHA);

                SelectObject(memDC, prev);
                DeleteObject(hBmp);
                DeleteDC(memDC);
                ReleaseDC(IntPtr.Zero, screenDC);
            }

            private static Bitmap BuildAlphaBitmap(Image src)
            {
                if (src == null)
                {
                    var fb = new Bitmap(200, 30, PixelFormat.Format32bppArgb);
                    using (var g = Graphics.FromImage(fb))
                        g.Clear(Color.FromArgb(200, 39, 39, 42));
                    return fb;
                }

                int w = src.Width;
                int h = src.Height;

                var screenW = Screen.PrimaryScreen.WorkingArea.Width;
                if (w > screenW * 0.85)
                {
                    float s = (screenW * 0.85f) / w;
                    w = (int)(w * s);
                    h = Math.Max(20, (int)(h * s));
                }

                var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.CompositingMode = CompositingMode.SourceOver;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.Clear(Color.Transparent);
                    g.DrawImage(src, 0, 0, w, h);
                }
                return bmp;
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing) { _alphaBmp?.Dispose(); _alphaBmp = null; }
                base.Dispose(disposing);
            }
        }

        #endregion

        #region CardControl

        private class CardControl : Control
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
            static readonly Color SuccessColor = Color.FromArgb(134, 239, 172);
            static readonly Color ErrorColor = Color.FromArgb(252, 165, 165);

            static readonly Font TitleFont = new Font("Segoe UI", 10f, FontStyle.Bold);
            static readonly Font BadgeFont = new Font("Segoe UI", 7f, FontStyle.Bold);
            static readonly Font PlaceholderFont = new Font("Segoe UI", 13f, FontStyle.Bold);

            private readonly Image _thumb;
            private readonly string _title;
            private readonly string _pptxPath;
            private readonly TaskPaneHost _host;

            private bool _hovered;
            private bool _mousePressed;
            private bool _dragging;
            private Point _dragStart;
            private GhostWindow _ghost;

            public CardControl(Image thumb, string title, string pptxPath, TaskPaneHost host)
            {
                _thumb = thumb;
                _title = title;
                _pptxPath = pptxPath;
                _host = host;
                Height = ThumbH + LabelH;
                Cursor = Cursors.Hand;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint
                       | ControlStyles.DoubleBuffer | ControlStyles.ResizeRedraw, true);
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
                    // Card background
                    using (var brush = new SolidBrush(_dragging ? CardDrag : (_hovered ? CardHover : CardBg)))
                        g.FillPath(brush, path);

                    g.SetClip(path);

                    // ── Thumbnail area (0 ~ ThumbH) ──
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

                    // ── Separator ──
                    using (var pen = new Pen(BorderNormal))
                        g.DrawLine(pen, 0, ThumbH, Width, ThumbH);

                    g.ResetClip();

                    // ── Label area (ThumbH ~ ThumbH+LabelH) ──
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

                    // ── Card border ──
                    using (var pen = new Pen(_dragging ? Accent : (_hovered ? Accent : BorderNormal),
                        (_hovered || _dragging) ? 1.5f : 0.5f))
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

            #region Mouse / Drag

            protected override void OnMouseEnter(EventArgs e)
            {
                _hovered = true;
                Invalidate();
                base.OnMouseEnter(e);
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                if (!_dragging) { _hovered = false; Invalidate(); }
                base.OnMouseLeave(e);
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    _dragStart = e.Location;
                    _mousePressed = true;
                }
                base.OnMouseDown(e);
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                if (_mousePressed && !_dragging && e.Button == MouseButtons.Left)
                {
                    if (Math.Abs(e.X - _dragStart.X) > SystemInformation.DragSize.Width / 2 ||
                        Math.Abs(e.Y - _dragStart.Y) > SystemInformation.DragSize.Height / 2)
                        BeginDrag();
                }

                if (_dragging && _ghost != null)
                    _ghost.MoveTo(PointToScreen(e.Location));

                base.OnMouseMove(e);
            }

            private void BeginDrag()
            {
                try
                {
                    LogDebug($"BeginDrag: {_title}");
                    ShapeInserter.CopyShapesToClipboard(_pptxPath);

                    _ghost = new GhostWindow(_thumb);
                    _ghost.MoveTo(PointToScreen(_dragStart));
                    _ghost.Show();

                    _dragging = true;
                    Capture = true;
                    Cursor.Current = Cursors.Cross;
                    _host.StatusLabel.Text = $"{_title} → 슬라이드에 놓으세요";
                    _host.StatusLabel.ForeColor = Accent;
                    Invalidate();
                }
                catch (Exception ex)
                {
                    _mousePressed = false;
                    DisposeGhost();
                    _host.StatusLabel.Text = $"드래그 실패: {ex.Message}";
                    _host.StatusLabel.ForeColor = ErrorColor;
                    LogDebug($"BeginDrag fail: {ex}");
                }
            }

            protected override void OnMouseUp(MouseEventArgs e)
            {
                if (_dragging)
                    EndDrag(e);
                else if (_mousePressed)
                    DoClickInsert();

                _mousePressed = false;
                _dragging = false;
                _hovered = ClientRectangle.Contains(PointToClient(Cursor.Position));
                Cursor = Cursors.Hand;
                Invalidate();
                base.OnMouseUp(e);
            }

            protected override void OnMouseCaptureChanged(EventArgs e)
            {
                if (_dragging)
                {
                    _dragging = false;
                    _mousePressed = false;
                    DisposeGhost();
                    Cursor = Cursors.Hand;
                    _host.ResetStatus();
                    Invalidate();
                }
                base.OnMouseCaptureChanged(e);
            }

            private void EndDrag(MouseEventArgs e)
            {
                Capture = false;
                DisposeGhost();

                var screenPos = PointToScreen(e.Location);
                var hostRect = _host.RectangleToScreen(_host.ClientRectangle);

                if (!hostRect.Contains(screenPos))
                {
                    try
                    {
                        var app = Globals.Application;
                        var window = app.ActiveWindow;
                        var slide = (PowerPoint.Slide)window.View.Slide;
                        var shapes = slide.Shapes.Paste();

                        PositionShapesAtCursor(shapes, screenPos, window);

                        _host.StatusLabel.Text = $"✓ {_title} 삽입 완료";
                        _host.StatusLabel.ForeColor = SuccessColor;
                    }
                    catch (Exception ex)
                    {
                        _host.StatusLabel.Text = $"삽입 실패: {ex.Message}";
                        _host.StatusLabel.ForeColor = ErrorColor;
                        LogDebug($"EndDrag paste fail: {ex}");
                    }
                }
                else
                {
                    _host.ResetStatus();
                }
            }

            private static void PositionShapesAtCursor(
                PowerPoint.ShapeRange shapes, Point screenPos, PowerPoint.DocumentWindow window)
            {
                try
                {
                    float refPt = 500f;
                    int ox = window.PointsToScreenPixelsX(0f);
                    int oy = window.PointsToScreenPixelsY(0f);
                    int rx = window.PointsToScreenPixelsX(refPt);
                    int ry = window.PointsToScreenPixelsY(refPt);

                    float scaleX = (rx - ox) / refPt;
                    float scaleY = (ry - oy) / refPt;

                    LogDebug($"Coord: o=({ox},{oy}) r=({rx},{ry}) scale=({scaleX:F3},{scaleY:F3}) mouse=({screenPos.X},{screenPos.Y})");

                    if (Math.Abs(scaleX) < 0.01f || Math.Abs(scaleY) < 0.01f) return;

                    float slideX = (screenPos.X - ox) / scaleX;
                    float slideY = (screenPos.Y - oy) / scaleY;
                    LogDebug($"Raw slide: ({slideX:F1},{slideY:F1})");

                    float slideW = window.Presentation.PageSetup.SlideWidth;
                    float slideH = window.Presentation.PageSetup.SlideHeight;
                    LogDebug($"Slide bounds: {slideW}x{slideH}");

                    slideX = Math.Max(0, Math.Min(slideX, slideW));
                    slideY = Math.Max(0, Math.Min(slideY, slideH));

                    int cnt = shapes.Count;
                    LogDebug($"Shape count: {cnt}");
                    if (cnt == 0) return;

                    float minL = float.MaxValue, minT = float.MaxValue;
                    float maxR = float.MinValue, maxB = float.MinValue;
                    for (int i = 1; i <= cnt; i++)
                    {
                        var s = shapes[i];
                        float l = s.Left, t = s.Top, w = s.Width, h = s.Height;
                        LogDebug($"  shape[{i}]: L={l:F1} T={t:F1} W={w:F1} H={h:F1}");
                        if (l < minL) minL = l;
                        if (t < minT) minT = t;
                        if (l + w > maxR) maxR = l + w;
                        if (t + h > maxB) maxB = t + h;
                    }

                    float cx = (minL + maxR) / 2f;
                    float cy = (minT + maxB) / 2f;
                    float dx = slideX - cx;
                    float dy = slideY - cy;
                    LogDebug($"Center: ({cx:F1},{cy:F1}), delta: ({dx:F1},{dy:F1})");

                    for (int i = 1; i <= cnt; i++)
                    {
                        shapes[i].Left += dx;
                        shapes[i].Top += dy;
                    }

                    LogDebug($"Positioned at slide({slideX:F0},{slideY:F0})");
                }
                catch (Exception ex)
                {
                    LogDebug($"PositionShapes (non-fatal): {ex.Message}");
                }
            }

            private void DoClickInsert()
            {
                try
                {
                    ShapeInserter.InsertToActiveSlide(_pptxPath);
                    _host.StatusLabel.Text = $"✓ {_title} 삽입 완료";
                    _host.StatusLabel.ForeColor = SuccessColor;
                }
                catch (Exception ex)
                {
                    _host.StatusLabel.Text = $"삽입 실패: {ex.Message}";
                    _host.StatusLabel.ForeColor = ErrorColor;
                    LogDebug($"ClickInsert fail: {ex}");
                }
            }

            private void DisposeGhost()
            {
                if (_ghost != null)
                {
                    _ghost.Close();
                    _ghost.Dispose();
                    _ghost = null;
                }
            }

            #endregion
        }

        #endregion
    }

    [ComImport]
    [Guid("CB5BDC81-93C1-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectSafety
    {
        [PreserveSig] int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions);
        [PreserveSig] int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions);
    }
}

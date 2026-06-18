using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    internal class DragHandler
    {
        static readonly Color Accent = Color.FromArgb(99, 102, 241);
        static readonly Color SuccessColor = Color.FromArgb(134, 239, 172);
        static readonly Color ErrorColor = Color.FromArgb(252, 165, 165);

        private readonly Control _owner;
        private readonly string _pptxPath;
        private readonly string _title;
        private readonly Image _thumb;
        private readonly Action<string, Color> _setStatus;
        private readonly Action _resetStatus;
        private readonly Func<Rectangle> _getHostBounds;

        private bool _mousePressed;
        private Point _dragStart;
        private GhostWindow _ghost;

        public bool IsDragging { get; private set; }

        public DragHandler(Control owner, string pptxPath, string title, Image thumb,
            Action<string, Color> setStatus, Action resetStatus, Func<Rectangle> getHostBounds)
        {
            _owner = owner;
            _pptxPath = pptxPath;
            _title = title;
            _thumb = thumb;
            _setStatus = setStatus;
            _resetStatus = resetStatus;
            _getHostBounds = getHostBounds;
        }

        public void HandleMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _dragStart = e.Location;
                _mousePressed = true;
            }
        }

        public void HandleMouseMove(MouseEventArgs e)
        {
            if (_mousePressed && !IsDragging && e.Button == MouseButtons.Left)
            {
                if (Math.Abs(e.X - _dragStart.X) > SystemInformation.DragSize.Width / 2 ||
                    Math.Abs(e.Y - _dragStart.Y) > SystemInformation.DragSize.Height / 2)
                    BeginDrag();
            }

            if (IsDragging && _ghost != null)
                _ghost.MoveTo(_owner.PointToScreen(e.Location));
        }

        public void HandleMouseUp(MouseEventArgs e)
        {
            if (IsDragging)
                EndDrag(e);
            else if (_mousePressed)
                DoClickInsert();

            _mousePressed = false;
            IsDragging = false;
        }

        public void HandleCaptureChanged()
        {
            if (IsDragging)
            {
                IsDragging = false;
                _mousePressed = false;
                DisposeGhost();
                _resetStatus();
            }
        }

        private void BeginDrag()
        {
            try
            {
                Logger.Log($"BeginDrag: {_title}");
                ShapeInserter.CopyShapesToClipboard(_pptxPath);

                _ghost = new GhostWindow(_thumb);
                _ghost.MoveTo(_owner.PointToScreen(_dragStart));
                _ghost.Show();

                IsDragging = true;
                _owner.Capture = true;
                Cursor.Current = Cursors.Cross;
                _setStatus($"{_title} → 슬라이드에 놓으세요", Accent);
                _owner.Invalidate();
            }
            catch (Exception ex)
            {
                _mousePressed = false;
                DisposeGhost();
                _setStatus($"드래그 실패: {ex.Message}", ErrorColor);
                Logger.Log($"BeginDrag fail: {ex}");
            }
        }

        private void EndDrag(MouseEventArgs e)
        {
            _owner.Capture = false;
            DisposeGhost();

            var screenPos = _owner.PointToScreen(e.Location);
            var hostRect = _getHostBounds();

            if (!hostRect.Contains(screenPos))
            {
                try
                {
                    var app = Globals.Application;
                    var window = app.ActiveWindow;
                    var slide = (PowerPoint.Slide)window.View.Slide;
                    var shapes = slide.Shapes.Paste();

                    CoordinateConverter.PositionShapesAtCursor(shapes, screenPos, window);

                    _setStatus($"✓ {_title} 삽입 완료", SuccessColor);
                }
                catch (Exception ex)
                {
                    _setStatus($"삽입 실패: {ex.Message}", ErrorColor);
                    Logger.Log($"EndDrag paste fail: {ex}");
                }
            }
            else
            {
                _resetStatus();
            }
        }

        private void DoClickInsert()
        {
            try
            {
                ShapeInserter.InsertToActiveSlide(_pptxPath);
                _setStatus($"✓ {_title} 삽입 완료", SuccessColor);
            }
            catch (Exception ex)
            {
                _setStatus($"삽입 실패: {ex.Message}", ErrorColor);
                Logger.Log($"ClickInsert fail: {ex}");
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
    }
}

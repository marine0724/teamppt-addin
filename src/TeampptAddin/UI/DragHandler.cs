using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 카드의 드래그앤드롭 + 클릭 삽입 로직을 관리.
    /// OLE DragDrop이 아닌 Win32 마우스 캡처 방식 (PowerMockup 스타일).
    ///
    /// 드래그 흐름:
    /// 1. MouseDown → 드래그 시작 위치 기록
    /// 2. MouseMove → 임계값(SystemInformation.DragSize) 초과 시 BeginDrag:
    ///    - ShapeInserter.CopyShapesToClipboard()로 Shape을 클립보드에 복사
    ///    - GhostWindow 생성 (썸네일 기반 반투명 윈도우)
    ///    - Capture = true로 Task Pane 밖의 마우스 이벤트도 수신
    /// 3. MouseMove (드래그 중) → GhostWindow를 커서 중앙에 이동
    /// 4. MouseUp → EndDrag:
    ///    - Task Pane 밖이면: slide.Shapes.Paste() + CoordinateConverter로 드롭 위치에 배치
    ///    - Task Pane 안이면: 취소 (ResetStatus)
    ///
    /// 클릭 삽입: 드래그 임계값 미달 시 DoClickInsert → ShapeInserter.InsertToActiveSlide
    ///
    /// 상태 업데이트는 setStatus/resetStatus 콜백으로 TaskPaneHost에 전달.
    /// </summary>
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

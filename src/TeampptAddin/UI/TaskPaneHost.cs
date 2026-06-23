using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
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

        // WinForms fallback controls
        private Label _statusLabel;
        private Panel _scrollPanel;
        private Panel _headerPanel;
        private int _assetCount;

        // WPF controls
        private ElementHost _elementHost;
        private AssetPanel _wpfPanel;

        // WPF drag state
        private GhostWindow _wpfDragGhost;
        private AssetCard _wpfDragCard;
        private bool _wpfDragging;

        private bool _loaded;

        public TaskPaneHost()
        {
            try
            {
                Logger.Log($"Constructor. STA={Thread.CurrentThread.GetApartmentState()}");
                InitUI();
            }
            catch (Exception ex)
            {
                Logger.Log($"Constructor FAILED: {ex}");
            }
        }

        private void InitUI()
        {
            BackColor = BgColor;

            _headerPanel = new Panel { Dock = DockStyle.Top, Height = 52, BackColor = HeaderBg };
            _headerPanel.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(50, 50, 55)))
                    e.Graphics.DrawLine(pen, 0, 51, _headerPanel.Width, 51);
            };
            _headerPanel.Controls.Add(new Label
            {
                Text = "TEAMPPT",
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = AccentColor,
                Location = new Point(16, 6),
                AutoSize = true
            });
            _headerPanel.Controls.Add(new Label
            {
                Text = "헤더 에셋",
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextDim,
                Location = new Point(16, 30),
                AutoSize = true
            });

            _statusLabel = new Label
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
            Controls.Add(_statusLabel);
            Controls.Add(_headerPanel);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (!_loaded && Width > 0 && _scrollPanel != null)
            {
                _loaded = true;
                InitWpfUI();
            }
        }

        #region WPF Initialization

        private void InitWpfUI()
        {
            try
            {
                _wpfPanel = new AssetPanel();
                LoadWpfCards();

                _elementHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _wpfPanel
                };

                _wpfPanel.CardClickInsert += OnWpfClickInsert;
                _wpfPanel.CardDragStart += OnWpfDragStart;
                _wpfPanel.StyleApplyRequested += OnStyleApply;

                _headerPanel.Visible = false;
                _statusLabel.Visible = false;
                _scrollPanel.Visible = false;
                Controls.Add(_elementHost);

                Logger.Log("WPF UI initialized successfully");
            }
            catch (Exception ex)
            {
                Logger.Log($"WPF init failed, falling back to WinForms: {ex.Message}");
                if (_elementHost != null)
                {
                    Controls.Remove(_elementHost);
                    _elementHost.Dispose();
                    _elementHost = null;
                }
                _wpfPanel = null;

                _headerPanel.Visible = true;
                _statusLabel.Visible = true;
                _scrollPanel.Visible = true;
                LoadCards();
            }
        }

        private void LoadWpfCards()
        {
            var assetsDir = Globals.AssetsDir;

            // Load metadata + styles
            var assets = AssetLoader.Load(assetsDir);
            var styles = StyleLoader.Load(assetsDir);
            IAiService ai;
            RemoteAssetCache remoteCache = null;
            try
            {
                var keysPath = Path.Combine(assetsDir, "api-keys.json");
                var keys = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(keysPath));
                var gemini = keys["gemini"]?.ToString();
                var supaUrl = keys["supabaseUrl"]?.ToString();
                var supaAnon = keys["supabaseAnonKey"]?.ToString();

                if (!string.IsNullOrEmpty(supaUrl) && !string.IsNullOrEmpty(supaAnon) && !string.IsNullOrEmpty(gemini))
                {
                    ai = new VectorRecommendService(supaUrl, supaAnon, gemini);
                    remoteCache = new RemoteAssetCache(supaUrl, supaAnon);
                    Logger.Log("[AI] VectorRecommendService (Supabase 벡터검색) 사용");
                }
                else
                {
                    ai = GeminiAiService.FromAssetsDir(assetsDir);
                    Logger.Log("[AI] Supabase 설정 없음 → GeminiAiService(로컬 카탈로그) 사용");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[AI] 초기화 실패, MockAiService 사용: {ex.Message}");
                ai = new MockAiService();
            }

            _wpfPanel.SetAssets(assets);
            _wpfPanel.InitAi(ai, styles, remoteCache);

            // Create cards with thumbnails
            foreach (var asset in assets)
            {
                var pptxPath  = Path.Combine(assetsDir, asset.File);
                var thumbPath = Path.Combine(Globals.ThumbnailDir,
                    Path.GetFileNameWithoutExtension(asset.File) + ".png");
                var thumb = ThumbnailService.LoadThumbnail(pptxPath, thumbPath);
                _wpfPanel.AddAssetCard(
                    new AssetCard(thumb, asset.Name, pptxPath, asset.Category, asset.UseWhen),
                    asset);
            }

            _wpfPanel.ResetStatus();

            var policy = new LocalFileAccessPolicy();
            if (policy.CanIngest)
            {
                _wpfPanel.ShowIngestButton();
                Logger.Log("[TaskPane] 관리자 모드 — 인제스트 버튼 표시");
            }
        }

        #endregion

        #region WPF Drag Handling

        private void OnWpfClickInsert(AssetCard card)
        {
            try
            {
                ShapeInserter.InsertToActiveSlide(card.PptxPath);
                _wpfPanel.SetStyleAnchorByFile(card.PptxPath);
                _wpfPanel.SetStatus($"✓ {card.Title} 삽입 완료",
                    ThemeResources.StatusSuccess.Color);
            }
            catch (Exception ex)
            {
                _wpfPanel.SetStatus($"삽입 실패: {ex.Message}",
                    ThemeResources.StatusError.Color);
                Logger.Log($"WPF ClickInsert fail: {ex}");
            }
        }

        private void OnWpfDragStart(AssetCard card)
        {
            try
            {
                Logger.Log($"WPF BeginDrag: {card.Title}");
                ShapeInserter.CopyShapesToClipboard(card.PptxPath);

                _wpfDragGhost = new GhostWindow(card.DrawingThumbnail);
                _wpfDragGhost.MoveTo(Cursor.Position);
                _wpfDragGhost.Show();

                _wpfDragCard = card;
                _wpfDragging = true;
                Capture = true;
                Cursor.Current = Cursors.Cross;

                _wpfPanel.SetStatus($"{card.Title} → 슬라이드에 놓으세요",
                    ThemeResources.Accent.Color);
            }
            catch (Exception ex)
            {
                _wpfDragging = false;
                DisposeWpfDragGhost();
                _wpfPanel.SetStatus($"드래그 실패: {ex.Message}",
                    ThemeResources.StatusError.Color);
                Logger.Log($"WPF BeginDrag fail: {ex}");
            }
        }

        private void OnStyleApply(StylePalette palette, StyleFont font)
        {
            try
            {
                var slide = (PowerPoint.Slide)Globals.Application.ActiveWindow.View.Slide;
                SlideStyleApplier.Apply(slide, palette, font);

                _wpfPanel.SetStatus(
                    $"✓ {palette?.Name ?? "팔레트"} · {font?.Name ?? "폰트"} 적용 완료",
                    ThemeResources.StatusSuccess.Color);
            }
            catch (Exception ex)
            {
                _wpfPanel.SetStatus($"적용 실패: {ex.Message}", ThemeResources.StatusError.Color);
                Logger.Log($"StyleApply fail: {ex}");
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_wpfDragging && _wpfDragGhost != null)
                _wpfDragGhost.MoveTo(PointToScreen(e.Location));
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (!_wpfDragging) return;

            var card = _wpfDragCard;
            _wpfDragging = false;
            _wpfDragCard = null;
            Capture = false;
            DisposeWpfDragGhost();

            var screenPos = PointToScreen(e.Location);
            var hostRect = RectangleToScreen(ClientRectangle);

            if (!hostRect.Contains(screenPos))
            {
                try
                {
                    var app = Globals.Application;
                    var window = app.ActiveWindow;
                    var slide = (PowerPoint.Slide)window.View.Slide;
                    var shapes = slide.Shapes.Paste();
                    CoordinateConverter.PositionShapesAtCursor(shapes, screenPos, window);

                    _wpfPanel.SetStatus($"✓ {card.Title} 삽입 완료",
                        ThemeResources.StatusSuccess.Color);
                }
                catch (Exception ex)
                {
                    _wpfPanel.SetStatus($"삽입 실패: {ex.Message}",
                        ThemeResources.StatusError.Color);
                    Logger.Log($"WPF EndDrag paste fail: {ex}");
                }
            }
            else
            {
                _wpfPanel.ResetStatus();
            }

            Cursor.Current = Cursors.Default;
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            base.OnMouseCaptureChanged(e);
            if (_wpfDragging)
            {
                _wpfDragging = false;
                _wpfDragCard = null;
                DisposeWpfDragGhost();
                _wpfPanel?.ResetStatus();
                Cursor.Current = Cursors.Default;
            }
        }

        private void DisposeWpfDragGhost()
        {
            if (_wpfDragGhost != null)
            {
                _wpfDragGhost.Close();
                _wpfDragGhost.Dispose();
                _wpfDragGhost = null;
            }
        }

        #endregion

        #region WinForms Fallback

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
                var thumb = ThumbnailService.LoadThumbnail(pptxPath, thumbPath);

                var card = new CardControl(
                    thumb, $"Header {i}", pptxPath,
                    setStatus: (text, color) => { _statusLabel.Text = text; _statusLabel.ForeColor = color; },
                    resetStatus: () => ResetStatus(),
                    getHostBounds: () => RectangleToScreen(ClientRectangle));

                card.Location = new Point(10, y);
                card.Width = _scrollPanel.ClientSize.Width - 20 - SystemInformation.VerticalScrollBarWidth;
                card.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
                _scrollPanel.Controls.Add(card);

                y += card.Height + 10;
                _assetCount++;
            }

            ResetStatus();
        }

        private void ResetStatus()
        {
            _statusLabel.Text = _assetCount > 0
                ? $"{_assetCount}개 에셋 \xb7 클릭 또는 드래그하여 삽입"
                : "Assets 폴더에 header_N.pptx 파일을 넣으세요";
            _statusLabel.ForeColor = TextDim;
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
    }
}

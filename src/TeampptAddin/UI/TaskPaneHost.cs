using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Newtonsoft.Json.Linq;
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

        // Remote asset support
        private RemoteAssetCache _remoteCache;
        private SupabaseClient _supaClient;

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
            var styles = StyleLoader.Load(assetsDir);
            IAiService ai;

            string supaUrl = null, supaAnon = null, gemini = null;
            try
            {
                var keysPath = Path.Combine(assetsDir, "api-keys.json");
                var keys = JObject.Parse(File.ReadAllText(keysPath));
                gemini = keys["gemini"]?.ToString();
                supaUrl = keys["supabaseUrl"]?.ToString();
                supaAnon = keys["supabaseAnonKey"]?.ToString();
            }
            catch (Exception ex)
            {
                Logger.Log($"[AI] api-keys 읽기 실패: {ex.Message}");
            }

            if (!string.IsNullOrEmpty(supaUrl) && !string.IsNullOrEmpty(supaAnon))
            {
                _remoteCache = new RemoteAssetCache(supaUrl, supaAnon);
                _supaClient = new SupabaseClient(supaUrl, supaAnon);
            }

            if (!string.IsNullOrEmpty(supaUrl) && !string.IsNullOrEmpty(supaAnon) && !string.IsNullOrEmpty(gemini))
            {
                ai = new VectorRecommendService(supaUrl, supaAnon, gemini);
                Logger.Log("[AI] VectorRecommendService (Supabase 벡터검색) 사용");
            }
            else
            {
                try { ai = GeminiAiService.FromAssetsDir(assetsDir); }
                catch { ai = new MockAiService(); }
                Logger.Log("[AI] Supabase 설정 없음 → 로컬 AI 사용");
            }

            _wpfPanel.InitAi(ai, styles, _remoteCache);

            if (_supaClient != null)
            {
                LoadRemoteAssetsAsync();
            }
            else
            {
                LoadLocalAssets(assetsDir);
            }

            var policy = new LocalFileAccessPolicy();
            if (policy.CanIngest)
            {
                _wpfPanel.ShowIngestButton();
                Logger.Log("[TaskPane] 관리자 모드 — 인제스트 버튼 표시");
            }
        }

        private void LoadLocalAssets(string assetsDir)
        {
            var assets = AssetLoader.Load(assetsDir);
            _wpfPanel.SetAssets(assets);

            foreach (var asset in assets)
            {
                var pptxPath = Path.Combine(assetsDir, asset.File);
                var thumbPath = Path.Combine(Globals.ThumbnailDir,
                    Path.GetFileNameWithoutExtension(asset.File) + ".png");
                var thumb = ThumbnailService.LoadThumbnail(pptxPath, thumbPath);
                _wpfPanel.AddAssetCard(
                    new AssetCard(thumb, asset.Name, pptxPath, asset.Category, asset.UseWhen),
                    asset);
            }
            _wpfPanel.ResetStatus();
        }

        private async void LoadRemoteAssetsAsync()
        {
            try
            {
                _wpfPanel.SetStatus("에셋 불러오는 중...", ThemeResources.TextSub.Color);
                var json = await _supaClient.GetAssetsAsync().ConfigureAwait(false);
                var rows = JArray.Parse(json);
                var assets = rows.OfType<JObject>()
                    .Select(SupabaseAssetMapper.Map)
                    .ToList();

                Logger.Log($"[TaskPane] Supabase에서 에셋 {assets.Count}개 로드");

                Invoke(new Action(() =>
                {
                    _wpfPanel.SetAssets(assets);

                    foreach (var asset in assets)
                    {
                        var card = new AssetCard(null, asset.Name, "",
                            asset.Category, asset.UseWhen);
                        _wpfPanel.AddAssetCard(card, asset);
                        LoadRemoteThumbAsync(card, asset);
                    }
                    _wpfPanel.ResetStatus();
                }));
            }
            catch (Exception ex)
            {
                Logger.Log($"[TaskPane] Supabase 에셋 로드 실패: {ex}");
                Invoke(new Action(() =>
                {
                    LoadLocalAssets(Globals.AssetsDir);
                }));
            }
        }

        private async void LoadRemoteThumbAsync(AssetCard card, HeaderAsset asset)
        {
            if (_remoteCache == null || asset.Extra == null) return;
            if (!asset.Extra.TryGetValue("remote_thumb", out var rt)) return;
            var remoteThumb = rt.ToString();
            if (string.IsNullOrEmpty(remoteThumb)) return;

            try
            {
                var localThumb = await _remoteCache.GetThumbAsync(remoteThumb).ConfigureAwait(false);
                Invoke(new Action(() =>
                {
                    var img = System.Drawing.Image.FromFile(localThumb);
                    card.SetThumbnail(img);
                }));
            }
            catch (Exception ex)
            {
                Logger.Log($"[TaskPane] 썸네일 로드 실패 ({asset.Name}): {ex.Message}");
            }
        }

        #endregion

        #region WPF Drag Handling

        private bool IsRemoteAsset(AssetCard card)
        {
            var asset = card.Tag as HeaderAsset;
            return asset?.Extra != null && asset.Extra.ContainsKey("remote_file")
                && string.IsNullOrEmpty(card.PptxPath);
        }

        private void OnWpfClickInsert(AssetCard card)
        {
            if (IsRemoteAsset(card))
            {
                var asset = (HeaderAsset)card.Tag;
                var remotePath = asset.Extra["remote_file"].ToString();
                InsertRemoteAssetCardAsync(card, remotePath, asset);
                return;
            }

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

        private async void InsertRemoteAssetCardAsync(AssetCard card, string remotePath, HeaderAsset asset)
        {
            try
            {
                _wpfPanel.SetStatus($"⬇ {card.Title} 다운로드 중...", ThemeResources.TextSub.Color);
                var localPath = await _remoteCache.GetPptxAsync(remotePath).ConfigureAwait(false);
                Invoke(new Action(() =>
                {
                    ShapeInserter.InsertToActiveSlide(localPath);
                    _wpfPanel.SetStyleAnchorByFile(remotePath);
                    _wpfPanel.SetStatus($"✓ {card.Title} 삽입 완료",
                        ThemeResources.StatusSuccess.Color);
                }));
            }
            catch (Exception ex)
            {
                Invoke(new Action(() =>
                    _wpfPanel.SetStatus($"삽입 실패: {ex.Message}",
                        ThemeResources.StatusError.Color)));
                Logger.Log($"[RemoteClickInsert] 실패: {ex}");
            }
        }

        private void OnWpfDragStart(AssetCard card)
        {
            if (IsRemoteAsset(card))
            {
                var asset = (HeaderAsset)card.Tag;
                var remotePath = asset.Extra["remote_file"].ToString();
                DragRemoteAssetCardAsync(card, remotePath, asset);
                return;
            }

            try
            {
                Logger.Log($"WPF BeginDrag: {card.Title}");

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

        private async void DragRemoteAssetCardAsync(AssetCard card, string remotePath, HeaderAsset asset)
        {
            try
            {
                _wpfPanel.SetStatus($"⬇ {card.Title} 다운로드 중...", ThemeResources.TextSub.Color);
                var localPath = await _remoteCache.GetPptxAsync(remotePath).ConfigureAwait(false);
                Invoke(new Action(() =>
                {
                    var thumbCachePath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "TeampptAddin", "cache", "thumb-shape",
                        Path.GetFileNameWithoutExtension(localPath) + ".png");
                    var thumbImg = ThumbnailService.LoadThumbnail(localPath, thumbCachePath);
                    var tempCard = new AssetCard(thumbImg, card.Title, localPath);
                    tempCard.Tag = asset;

                    _wpfDragGhost = new GhostWindow(thumbImg);
                    _wpfDragGhost.MoveTo(Cursor.Position);
                    _wpfDragGhost.Show();

                    _wpfDragCard = tempCard;
                    _wpfDragging = true;
                    Capture = true;
                    Cursor.Current = Cursors.Cross;

                    _wpfPanel.SetStatus($"{card.Title} → 슬라이드에 놓으세요",
                        ThemeResources.Accent.Color);
                }));
            }
            catch (Exception ex)
            {
                Invoke(new Action(() =>
                    _wpfPanel.SetStatus($"다운로드 실패: {ex.Message}",
                        ThemeResources.StatusError.Color)));
                Logger.Log($"[RemoteDrag] 다운로드 실패: {ex}");
            }
        }

        private void OnStyleApply(StylePalette palette, StyleFont font)
        {
            Logger.Log($"[Style] OnStyleApply fired: palette={palette?.Name ?? "NULL"}, font={font?.Name ?? "NULL"}");
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
                    var shapes = ShapeInserter.InsertToActiveSlide(card.PptxPath);

                    if (shapes != null)
                        CoordinateConverter.PositionShapesAtCursor(
                            shapes, screenPos, Globals.Application.ActiveWindow);

                    _wpfPanel.SetStyleAnchorByFile(card.PptxPath);
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

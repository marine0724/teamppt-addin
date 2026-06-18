using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace TeampptAddin
{
    /// <summary>
    /// PowerPoint Task Pane의 COM 호스팅 컨테이너.
    /// Connect.CTPFactoryAvailable에서 CreateCTP("TeampptAddin.TaskPaneHost")로 생성됨.
    ///
    /// COM 호스팅 요구사항:
    /// - [ComVisible], [Guid], [ProgId] 어트리뷰트 필수
    /// - IObjectSafety 구현 (ActiveX 보안 검증용)
    /// - 레지스트리에 Control 카테고리 {40FC6ED4-...} 수동 등록 필수
    ///
    /// 초기화 흐름:
    /// 1. 생성자: InitUI()로 WinForms 레이아웃 구성 (header, scrollPanel, statusLabel)
    /// 2. OnSizeChanged(Width > 0): LoadCards()로 에셋 카드 로드
    ///    → COM 초기화 직후에는 Size가 0x0이므로, Width > 0인 첫 SizeChanged에서만 실행
    ///    → WPF ElementHost도 이 시점에서 생성하면 COM 충돌 없음 (검증 완료)
    ///
    /// 썸네일 로딩 전략 (LoadThumbnail):
    /// 1순위: 캐시 파일 (pptx 수정일과 비교)
    /// 2순위: ThumbnailGenerator.Generate (COM Shape-only export)
    /// 3순위: pptx ZIP 내부의 docProps/thumbnail 이미지
    /// </summary>
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

        private Label _statusLabel;
        private Panel _scrollPanel;
        private bool _loaded;
        private int _assetCount;

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
                ? $"{_assetCount}개 에셋 · 클릭 또는 드래그하여 삽입"
                : "Assets 폴더에 header_N.pptx 파일을 넣으세요";
            _statusLabel.ForeColor = TextDim;
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
                ThumbnailGenerator.Generate(pptxPath, cachePath);
                if (File.Exists(cachePath))
                    return LoadImageNoLock(cachePath);
            }
            catch (Exception ex)
            {
                Logger.Log($"COM thumb fail [{Path.GetFileName(pptxPath)}]: {ex.Message}");
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
                Logger.Log($"ZIP thumb fail [{Path.GetFileName(pptxPath)}]: {ex.Message}");
            }

            return null;
        }

        private static Image LoadImageNoLock(string path)
        {
            return Image.FromStream(new MemoryStream(File.ReadAllBytes(path)));
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

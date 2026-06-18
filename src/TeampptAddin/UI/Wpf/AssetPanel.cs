using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace TeampptAddin
{
    internal class AssetPanel : UserControl
    {
        public event Action<AssetCard> CardClickInsert;
        public event Action<AssetCard> CardDragStart;

        private readonly StackPanel _cardStack;
        private readonly TextBlock _statusText;
        private int _assetCount;

        static readonly SolidColorBrush BgBrush = Freeze(new SolidColorBrush(Color.FromRgb(24, 24, 27)));
        static readonly SolidColorBrush HeaderBgBrush = Freeze(new SolidColorBrush(Color.FromRgb(30, 30, 34)));
        static readonly SolidColorBrush AccentBrush = Freeze(new SolidColorBrush(Color.FromRgb(99, 102, 241)));
        static readonly SolidColorBrush TextDimBrush = Freeze(new SolidColorBrush(Color.FromRgb(113, 113, 122)));
        static readonly SolidColorBrush BorderLineBrush = Freeze(new SolidColorBrush(Color.FromRgb(50, 50, 55)));

        public AssetPanel()
        {
            Background = BgBrush;
            FontFamily = new FontFamily("Segoe UI");

            var dock = new DockPanel { LastChildFill = true };

            // Header
            var header = new Border
            {
                Background = HeaderBgBrush,
                Height = 52,
                BorderBrush = BorderLineBrush,
                BorderThickness = new Thickness(0, 0, 0, 1)
            };
            var headerStack = new StackPanel { Margin = new Thickness(16, 8, 0, 0) };
            headerStack.Children.Add(new TextBlock
            {
                Text = "TEAMPPT",
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = AccentBrush
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = "헤더 에셋",
                FontSize = 9,
                Foreground = TextDimBrush,
                Margin = new Thickness(0, 2, 0, 0)
            });
            header.Child = headerStack;
            DockPanel.SetDock(header, Dock.Top);
            dock.Children.Add(header);

            // Status bar
            var statusBar = new Border
            {
                Background = HeaderBgBrush,
                Height = 28
            };
            _statusText = new TextBlock
            {
                Text = "로딩 중...",
                FontSize = 9,
                Foreground = TextDimBrush,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(14, 0, 0, 0)
            };
            statusBar.Child = _statusText;
            DockPanel.SetDock(statusBar, Dock.Bottom);
            dock.Children.Add(statusBar);

            // Scrollable card area
            _cardStack = new StackPanel { Margin = new Thickness(0, 5, 0, 5) };
            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Background = BgBrush,
                Content = _cardStack
            };
            dock.Children.Add(scrollViewer);

            Content = dock;
        }

        public void AddCard(AssetCard card)
        {
            card.ClickInsertRequested += c => CardClickInsert?.Invoke(c);
            card.DragStartRequested += c => CardDragStart?.Invoke(c);
            _cardStack.Children.Add(card);
            _assetCount++;
        }

        public void SetStatus(string text, Color color)
        {
            _statusText.Text = text;
            _statusText.Foreground = new SolidColorBrush(color);
        }

        public void ResetStatus()
        {
            _statusText.Text = _assetCount > 0
                ? $"{_assetCount}개 에셋 \xb7 클릭 또는 드래그하여 삽입"
                : "Assets 폴더에 header_N.pptx 파일을 넣으세요";
            _statusText.Foreground = TextDimBrush;
        }

        private static SolidColorBrush Freeze(SolidColorBrush brush)
        {
            brush.Freeze();
            return brush;
        }
    }
}

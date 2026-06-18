using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;

namespace TeampptAddin
{
    public class TaskPaneControl : System.Windows.Controls.UserControl
    {
        private readonly StackPanel _cardPanel;
        private readonly TextBlock _statusText;
        private readonly List<HeaderAsset> _assets = new List<HeaderAsset>();

        private static readonly Color BgColor = Color.FromRgb(24, 24, 27);
        private static readonly Color CardColor = Color.FromRgb(39, 39, 42);
        private static readonly Color CardHoverColor = Color.FromRgb(52, 52, 56);
        private static readonly Color AccentColor = Color.FromRgb(99, 102, 241);
        private static readonly Color TextColor = Color.FromRgb(228, 228, 231);
        private static readonly Color TextDimColor = Color.FromRgb(113, 113, 122);

        public TaskPaneControl()
        {
            Background = new SolidColorBrush(BgColor);

            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Header
            var header = CreateHeader();
            Grid.SetRow(header, 0);
            root.Children.Add(header);

            // Scrollable card list
            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(0)
            };

            _cardPanel = new StackPanel
            {
                Margin = new Thickness(12, 0, 12, 12)
            };
            scrollViewer.Content = _cardPanel;
            Grid.SetRow(scrollViewer, 1);
            root.Children.Add(scrollViewer);

            // Status bar
            _statusText = new TextBlock
            {
                Text = "로딩 중...",
                Foreground = new SolidColorBrush(TextDimColor),
                FontSize = 11,
                Margin = new Thickness(16, 8, 16, 8)
            };
            Grid.SetRow(_statusText, 2);
            root.Children.Add(_statusText);

            Content = root;

            Loaded += OnLoaded;
        }

        private Border CreateHeader()
        {
            var border = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(30, 30, 34)),
                Padding = new Thickness(16, 14, 16, 14),
                BorderBrush = new SolidColorBrush(Color.FromRgb(50, 50, 55)),
                BorderThickness = new Thickness(0, 0, 0, 1)
            };

            var stack = new StackPanel();

            var title = new TextBlock
            {
                Text = "TEAMPPT",
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(AccentColor)
            };
            stack.Children.Add(title);

            var subtitle = new TextBlock
            {
                Text = "헤더 에셋",
                FontSize = 12,
                Foreground = new SolidColorBrush(TextDimColor),
                Margin = new Thickness(0, 2, 0, 0)
            };
            stack.Children.Add(subtitle);

            border.Child = stack;
            return border;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            await LoadAssetsAsync();
        }

        private async System.Threading.Tasks.Task LoadAssetsAsync()
        {
            _statusText.Text = "에셋 로딩 중...";

            var assetsDir = Globals.AssetsDir;
            var thumbDir = Globals.ThumbnailDir;

            await System.Threading.Tasks.Task.Run(() =>
            {
                for (int i = 1; i <= 7; i++)
                {
                    var pptxPath = Path.Combine(assetsDir, $"header_{i}.pptx");
                    if (!File.Exists(pptxPath)) continue;

                    var thumbPath = Path.Combine(thumbDir, $"header_{i}.png");

                    _assets.Add(new HeaderAsset
                    {
                        Index = i,
                        Name = $"Header {i}",
                        PptxPath = pptxPath,
                        ThumbnailPath = thumbPath
                    });
                }
            });

            // Generate thumbnails (needs PowerPoint COM on UI thread)
            foreach (var asset in _assets)
            {
                try
                {
                    ShapeInserter.GenerateThumbnail(asset.PptxPath, asset.ThumbnailPath);
                }
                catch { /* thumbnail generation failed, will show placeholder */ }
            }

            // Build cards
            _cardPanel.Children.Clear();
            foreach (var asset in _assets)
            {
                _cardPanel.Children.Add(CreateCard(asset));
            }

            _statusText.Text = $"{_assets.Count}개 에셋 · 드래그하여 슬라이드에 삽입";
        }

        private Border CreateCard(HeaderAsset asset)
        {
            var card = new Border
            {
                CornerRadius = new CornerRadius(10),
                Background = new SolidColorBrush(CardColor),
                Margin = new Thickness(0, 0, 0, 10),
                Cursor = Cursors.Hand,
                ClipToBounds = true,
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    BlurRadius = 8,
                    ShadowDepth = 2,
                    Opacity = 0.3,
                    Direction = 270
                },
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new ScaleTransform(1, 1)
            };

            var stack = new StackPanel();

            // Thumbnail image
            var imageContainer = new Border
            {
                CornerRadius = new CornerRadius(10, 10, 0, 0),
                ClipToBounds = true,
                Background = new SolidColorBrush(Color.FromRgb(32, 32, 36))
            };

            var image = new Image
            {
                Height = 120,
                Stretch = Stretch.Uniform,
                Margin = new Thickness(8, 8, 8, 4)
            };

            // Load thumbnail
            if (File.Exists(asset.ThumbnailPath))
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(asset.ThumbnailPath);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                image.Source = bitmap;
                asset.Thumbnail = bitmap;
            }
            else
            {
                image.Source = CreatePlaceholder(asset.Index);
            }

            imageContainer.Child = image;
            stack.Children.Add(imageContainer);

            // Label area
            var labelArea = new Border
            {
                Padding = new Thickness(12, 8, 12, 10)
            };

            var labelStack = new StackPanel { Orientation = Orientation.Horizontal };

            var label = new TextBlock
            {
                Text = asset.Name,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(TextColor),
                VerticalAlignment = VerticalAlignment.Center
            };
            labelStack.Children.Add(label);

            var badge = new Border
            {
                CornerRadius = new CornerRadius(4),
                Background = new SolidColorBrush(Color.FromArgb(30, 99, 102, 241)),
                Padding = new Thickness(6, 2, 6, 2),
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            var badgeText = new TextBlock
            {
                Text = "DRAG",
                FontSize = 9,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(AccentColor)
            };
            badge.Child = badgeText;
            labelStack.Children.Add(badge);

            labelArea.Child = labelStack;
            stack.Children.Add(labelArea);

            card.Child = stack;
            card.Tag = asset;

            // Hover animation
            card.MouseEnter += (s, e) =>
            {
                var scaleAnim = new DoubleAnimation(1.03, TimeSpan.FromMilliseconds(150))
                { EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } };
                var transform = (ScaleTransform)card.RenderTransform;
                transform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
                transform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);

                card.Background = new SolidColorBrush(CardHoverColor);
                ((DropShadowEffect)card.Effect).BlurRadius = 16;
                ((DropShadowEffect)card.Effect).Opacity = 0.5;
            };

            card.MouseLeave += (s, e) =>
            {
                var scaleAnim = new DoubleAnimation(1.0, TimeSpan.FromMilliseconds(150))
                { EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } };
                var transform = (ScaleTransform)card.RenderTransform;
                transform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
                transform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);

                card.Background = new SolidColorBrush(CardColor);
                ((DropShadowEffect)card.Effect).BlurRadius = 8;
                ((DropShadowEffect)card.Effect).Opacity = 0.3;
            };

            // Drag-and-drop
            card.MouseLeftButtonDown += Card_MouseLeftButtonDown;

            // Double-click fallback: insert at slide center
            card.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ClickCount == 2)
                {
                    var a = (HeaderAsset)card.Tag;
                    ShapeInserter.InsertToActiveSlide(a.PptxPath);
                    _statusText.Text = $"✓ {a.Name} 삽입 완료";
                    e.Handled = true;
                }
            };

            return card;
        }

        private Point _dragStartPoint;

        private void Card_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount >= 2) return;
            _dragStartPoint = e.GetPosition(null);

            var card = (Border)sender;
            card.MouseMove += Card_MouseMove;
            card.MouseLeftButtonUp += (s2, e2) =>
            {
                card.MouseMove -= Card_MouseMove;
            };
        }

        private void Card_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed) return;

            var pos = e.GetPosition(null);
            var diff = pos - _dragStartPoint;

            if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            var card = (Border)sender;
            card.MouseMove -= Card_MouseMove;

            var asset = (HeaderAsset)card.Tag;
            _statusText.Text = $"드래그 중: {asset.Name}";

            try
            {
                // Copy shapes from source pptx to clipboard
                ShapeInserter.CopyShapesToClipboard(asset.PptxPath);

                // Get clipboard data (PowerPoint native shape format)
                var dataObject = System.Windows.Clipboard.GetDataObject();
                if (dataObject != null)
                {
                    // Start OLE drag-drop with clipboard data
                    var result = DragDrop.DoDragDrop(card, dataObject, DragDropEffects.Copy);

                    if (result == DragDropEffects.None)
                    {
                        // Drop was outside PowerPoint or cancelled
                        // Fallback: paste to active slide
                        try
                        {
                            var app = Globals.Application;
                            if (app?.ActiveWindow != null)
                            {
                                var slide = (Microsoft.Office.Interop.PowerPoint.Slide)app.ActiveWindow.View.Slide;
                                slide.Shapes.Paste();
                                _statusText.Text = $"✓ {asset.Name} 삽입 완료";
                            }
                        }
                        catch { _statusText.Text = "슬라이드를 선택한 후 다시 시도하세요"; }
                    }
                    else
                    {
                        _statusText.Text = $"✓ {asset.Name} 삽입 완료";
                    }
                }
            }
            catch (Exception ex)
            {
                _statusText.Text = $"삽입 실패: {ex.Message}";
            }
        }

        private BitmapSource CreatePlaceholder(int index)
        {
            int width = 480, height = 270;
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                dc.DrawRectangle(
                    new SolidColorBrush(Color.FromRgb(45, 45, 50)),
                    null, new Rect(0, 0, width, height));

                var text = new FormattedText(
                    $"Header {index}",
                    System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Segoe UI"),
                    24,
                    new SolidColorBrush(TextDimColor),
                    1.0);
                dc.DrawText(text,
                    new Point((width - text.Width) / 2, (height - text.Height) / 2));
            }

            var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            bitmap.Freeze();
            return bitmap;
        }
    }
}

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using DrawingImage = System.Drawing.Image;

namespace TeampptAddin
{
    internal class AssetCard : Border
    {
        public event Action<AssetCard> ClickInsertRequested;
        public event Action<AssetCard> DragStartRequested;

        public string PptxPath { get; }
        public string Title { get; }
        public DrawingImage DrawingThumbnail { get; }

        private bool _mousePressed;
        private Point _dragStart;

        static readonly SolidColorBrush CardBgBrush = Freeze(new SolidColorBrush(Color.FromRgb(39, 39, 42)));
        static readonly SolidColorBrush CardHoverBrush = Freeze(new SolidColorBrush(Color.FromRgb(52, 52, 56)));
        static readonly SolidColorBrush ThumbBgBrush = Freeze(new SolidColorBrush(Color.FromRgb(28, 28, 32)));
        static readonly SolidColorBrush AccentBrush = Freeze(new SolidColorBrush(Color.FromRgb(99, 102, 241)));
        static readonly SolidColorBrush TextMainBrush = Freeze(new SolidColorBrush(Color.FromRgb(228, 228, 231)));
        static readonly SolidColorBrush TextSubBrush = Freeze(new SolidColorBrush(Color.FromRgb(113, 113, 122)));
        static readonly SolidColorBrush BorderNormalBrush = Freeze(new SolidColorBrush(Color.FromRgb(50, 50, 55)));
        static readonly SolidColorBrush BadgeBgBrush = Freeze(new SolidColorBrush(Color.FromArgb(30, 99, 102, 241)));

        public AssetCard(DrawingImage thumb, string title, string pptxPath)
        {
            PptxPath = pptxPath;
            Title = title;
            DrawingThumbnail = thumb;

            CornerRadius = new CornerRadius(10);
            Background = CardBgBrush;
            BorderBrush = BorderNormalBrush;
            BorderThickness = new Thickness(1);
            Margin = new Thickness(10, 5, 10, 5);
            Cursor = Cursors.Hand;
            ClipToBounds = true;
            SnapsToDevicePixels = true;

            RenderTransform = new ScaleTransform(1, 1);
            RenderTransformOrigin = new Point(0.5, 0.5);

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(120) });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(36) });

            var thumbPanel = new Border { Background = ThumbBgBrush, ClipToBounds = true };
            if (thumb != null)
            {
                var bitmapSource = ConvertToBitmapSource(thumb);
                thumbPanel.Child = new Image
                {
                    Source = bitmapSource,
                    Stretch = Stretch.Uniform,
                    Margin = new Thickness(8),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
            }
            else
            {
                thumbPanel.Child = new TextBlock
                {
                    Text = title,
                    Foreground = TextSubBrush,
                    FontSize = 13,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
            }
            Grid.SetRow(thumbPanel, 0);
            grid.Children.Add(thumbPanel);

            var separator = new Border
            {
                Height = 1,
                Background = BorderNormalBrush,
                VerticalAlignment = VerticalAlignment.Bottom
            };
            Grid.SetRow(separator, 0);
            grid.Children.Add(separator);

            var labelGrid = new Grid();
            labelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            labelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var titleText = new TextBlock
            {
                Text = title,
                Foreground = TextMainBrush,
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                FontFamily = new FontFamily("Segoe UI"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(titleText, 0);
            labelGrid.Children.Add(titleText);

            var badge = new Border
            {
                Background = BadgeBgBrush,
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(6, 2, 6, 2),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0),
                Child = new TextBlock
                {
                    Text = "DRAG",
                    Foreground = AccentBrush,
                    FontSize = 8,
                    FontWeight = FontWeights.Bold,
                    FontFamily = new FontFamily("Segoe UI")
                }
            };
            Grid.SetColumn(badge, 1);
            labelGrid.Children.Add(badge);

            Grid.SetRow(labelGrid, 1);
            grid.Children.Add(labelGrid);

            Child = grid;

            MouseEnter += OnCardMouseEnter;
            MouseLeave += OnCardMouseLeave;
            MouseLeftButtonDown += OnCardMouseDown;
            MouseMove += OnCardMouseMove;
            MouseLeftButtonUp += OnCardMouseUp;
        }

        private void OnCardMouseEnter(object sender, MouseEventArgs e)
        {
            Background = CardHoverBrush;
            BorderBrush = AccentBrush;
            AnimateScale(1.02);
        }

        private void OnCardMouseLeave(object sender, MouseEventArgs e)
        {
            Background = CardBgBrush;
            BorderBrush = BorderNormalBrush;
            AnimateScale(1.0);
        }

        private void OnCardMouseDown(object sender, MouseButtonEventArgs e)
        {
            _mousePressed = true;
            _dragStart = e.GetPosition(this);
        }

        private void OnCardMouseMove(object sender, MouseEventArgs e)
        {
            if (!_mousePressed || e.LeftButton != MouseButtonState.Pressed) return;

            var pos = e.GetPosition(this);
            if (Math.Abs(pos.X - _dragStart.X) > SystemParameters.MinimumHorizontalDragDistance ||
                Math.Abs(pos.Y - _dragStart.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                _mousePressed = false;
                DragStartRequested?.Invoke(this);
            }
        }

        private void OnCardMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_mousePressed)
            {
                _mousePressed = false;
                ClickInsertRequested?.Invoke(this);
            }
        }

        private void AnimateScale(double target)
        {
            var scale = (ScaleTransform)RenderTransform;
            var anim = new DoubleAnimation(target, TimeSpan.FromMilliseconds(150))
            {
                EasingFunction = new QuadraticEase()
            };
            scale.BeginAnimation(ScaleTransform.ScaleXProperty, anim);
            scale.BeginAnimation(ScaleTransform.ScaleYProperty, anim);
        }

        private static BitmapSource ConvertToBitmapSource(DrawingImage img)
        {
            using (var ms = new MemoryStream())
            {
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.StreamSource = ms;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
        }

        private static SolidColorBrush Freeze(SolidColorBrush brush)
        {
            brush.Freeze();
            return brush;
        }
    }
}

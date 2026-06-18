using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
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

        private readonly string _category;
        private readonly string _useWhen;
        private readonly BitmapSource _bitmapThumb;

        private bool _mousePressed;
        private Point _dragStart;

        private static AssetCard _currentPopupCard;
        private static DispatcherTimer _closeTimer;

        private Popup _popup;
        private DispatcherTimer _popupTimer;

        static AssetCard()
        {
            _closeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _closeTimer.Tick += (s, e) =>
            {
                _closeTimer.Stop();
                if (_currentPopupCard != null)
                {
                    _currentPopupCard.ClosePopupInternal();
                    _currentPopupCard = null;
                }
            };
        }

        public AssetCard(DrawingImage thumb, string title, string pptxPath,
            string category = "", string useWhen = "")
        {
            PptxPath = pptxPath;
            Title = title;
            DrawingThumbnail = thumb;
            _category = category;
            _useWhen = useWhen;
            _bitmapThumb = thumb != null ? ConvertToBitmapSource(thumb) : null;

            CornerRadius = ThemeResources.RadiusCard;
            Background = ThemeResources.BgCard;
            BorderBrush = ThemeResources.BorderCard;
            BorderThickness = new Thickness(1);
            Margin = new Thickness(10, 5, 10, 5);
            Cursor = Cursors.Hand;
            ClipToBounds = true;
            SnapsToDevicePixels = true;

            ThemeResources.ApplyRoundedClip(this, 13);

            RenderTransform = new ScaleTransform(1, 1);
            RenderTransformOrigin = new Point(0.5, 0.5);

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(100) });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(36) });

            var thumbBorder = new Border { Background = ThemeResources.BgThumb, ClipToBounds = true };
            if (_bitmapThumb != null)
            {
                thumbBorder.Child = new Image
                {
                    Source = _bitmapThumb,
                    Stretch = Stretch.Uniform,
                    Margin = new Thickness(8),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
            }
            else
            {
                thumbBorder.Child = new TextBlock
                {
                    Text = title,
                    Foreground = ThemeResources.TextSub,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
            }
            Grid.SetRow(thumbBorder, 0);
            grid.Children.Add(thumbBorder);

            var sep = new Border
            {
                Height = 1,
                Background = ThemeResources.BorderBase,
                VerticalAlignment = VerticalAlignment.Bottom
            };
            Grid.SetRow(sep, 0);
            grid.Children.Add(sep);

            var labelGrid = new Grid();
            labelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            labelGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var titleText = new TextBlock
            {
                Text = title,
                Foreground = ThemeResources.TextMain,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(titleText, 0);
            labelGrid.Children.Add(titleText);

            var badge = new Border
            {
                Background = ThemeResources.BgBadge,
                CornerRadius = ThemeResources.RadiusBadge,
                Padding = new Thickness(6, 2, 6, 2),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0),
                Child = new TextBlock
                {
                    Text = "DRAG",
                    Foreground = ThemeResources.TextAccent,
                    FontSize = 8,
                    FontWeight = FontWeights.Bold,
                    FontFamily = ThemeResources.FontBase
                }
            };
            Grid.SetColumn(badge, 1);
            labelGrid.Children.Add(badge);

            Grid.SetRow(labelGrid, 1);
            grid.Children.Add(labelGrid);

            Child = grid;

            _popupTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _popupTimer.Tick += OnPopupTimerTick;

            MouseEnter += OnMouseEnter;
            MouseLeave += OnMouseLeave;
            MouseLeftButtonDown += OnMouseDown;
            MouseMove += OnMouseMove;
            MouseLeftButtonUp += OnMouseUp;
        }

        private void OnMouseEnter(object sender, MouseEventArgs e)
        {
            Background = ThemeResources.BgCardHover;
            BorderBrush = ThemeResources.BorderCardHover;
            AnimateScale(1.02);

            _closeTimer.Stop();

            if (_currentPopupCard != null && _currentPopupCard != this)
            {
                _currentPopupCard.ClosePopupInternal();
                _currentPopupCard = null;
            }

            if (_currentPopupCard == this) return;

            if (_closeTimer.Tag != null)
                ShowPopup();
            else
                _popupTimer.Start();
        }

        private void OnMouseLeave(object sender, MouseEventArgs e)
        {
            Background = ThemeResources.BgCard;
            BorderBrush = ThemeResources.BorderCard;
            AnimateScale(1.0);
            _popupTimer.Stop();

            if (_popup != null)
                _closeTimer.Start();
        }

        private void OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            _mousePressed = true;
            _dragStart = e.GetPosition(this);
            ClosePopup();
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
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

        private void OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_mousePressed) return;
            _mousePressed = false;
            ClickInsertRequested?.Invoke(this);
        }

        // ── Popup ────────────────────────────────────────────────────

        private void OnPopupTimerTick(object sender, EventArgs e)
        {
            _popupTimer.Stop();
            ShowPopup();
        }

        private void ShowPopup()
        {
            if (_popup != null) return;

            var content = BuildPopupContent();
            content.Opacity = 0;
            content.RenderTransform = new TranslateTransform(6, 0);

            _popup = new Popup
            {
                Child = content,
                PlacementTarget = this,
                Placement = PlacementMode.Left,
                AllowsTransparency = true,
                StaysOpen = true,
                IsHitTestVisible = false,
                IsOpen = true
            };

            _currentPopupCard = this;
            _closeTimer.Tag = this;

            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(150));
            var slideIn = new DoubleAnimation(6, 0, TimeSpan.FromMilliseconds(150));
            content.BeginAnimation(OpacityProperty, fadeIn);
            ((TranslateTransform)content.RenderTransform).BeginAnimation(TranslateTransform.XProperty, slideIn);
        }

        private void ClosePopupInternal()
        {
            _popupTimer.Stop();
            if (_popup != null)
            {
                _popup.IsOpen = false;
                _popup = null;
            }
        }

        private void ClosePopup()
        {
            ClosePopupInternal();
            _currentPopupCard = null;
            _closeTimer.Tag = null;
            _closeTimer.Stop();
        }

        private Border BuildPopupContent()
        {
            var outer = new Border
            {
                Width = 320,
                CornerRadius = new CornerRadius(16),
                ClipToBounds = true,
                Background = ThemeResources.BgBase,
                BorderBrush = ThemeResources.BorderCard,
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 8, 0),
                Effect = new DropShadowEffect
                {
                    Color = Color.FromRgb(0x19, 0x1F, 0x28),
                    BlurRadius = 24,
                    ShadowDepth = 4,
                    Opacity = 0.10,
                    Direction = 270
                }
            };

            var stack = new StackPanel();

            // ① Thumbnail
            var thumbArea = new Border
            {
                Height = 180,
                ClipToBounds = true,
                CornerRadius = new CornerRadius(16, 16, 0, 0),
                Background = ThemeResources.BgThumb
            };
            ThemeResources.ApplyRoundedClip(thumbArea, new CornerRadius(16, 16, 0, 0));
            if (_bitmapThumb != null)
            {
                thumbArea.Child = new Image
                {
                    Source = _bitmapThumb,
                    Stretch = Stretch.Uniform
                };
            }
            else
            {
                thumbArea.Child = new TextBlock
                {
                    Text = Title,
                    Foreground = ThemeResources.TextSub,
                    FontSize = 13,
                    FontWeight = FontWeights.SemiBold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
            }
            stack.Children.Add(thumbArea);

            // ② Separator
            stack.Children.Add(new Border { Height = 1, Background = ThemeResources.BorderBase });

            // ③ Meta (name + category badge)
            var metaGrid = new Grid();
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var nameText = new TextBlock
            {
                Text = Title,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain,
                FontFamily = ThemeResources.FontBase,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(nameText, 0);
            metaGrid.Children.Add(nameText);

            if (!string.IsNullOrEmpty(_category))
            {
                var catBadge = new Border
                {
                    Background = ThemeResources.BgBadge,
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(6, 2, 6, 2),
                    Margin = new Thickness(6, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = _category,
                        FontSize = 10,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = ThemeResources.TextAccent,
                        FontFamily = ThemeResources.FontBase
                    }
                };
                Grid.SetColumn(catBadge, 1);
                metaGrid.Children.Add(catBadge);
            }

            stack.Children.Add(new Border
            {
                Padding = new Thickness(12, 10, 12, 8),
                Child = metaGrid
            });

            // ④ UseWhen
            if (!string.IsNullOrEmpty(_useWhen))
            {
                stack.Children.Add(new Border
                {
                    Padding = new Thickness(12, 0, 12, 10),
                    Child = new TextBlock
                    {
                        Text = _useWhen,
                        FontSize = 11,
                        Foreground = ThemeResources.TextSub,
                        FontFamily = ThemeResources.FontBase,
                        TextWrapping = TextWrapping.Wrap,
                        LineHeight = 16
                    }
                });
            }

            // ⑤ Hint
            var hintPanel = new StackPanel { Orientation = Orientation.Horizontal };
            hintPanel.Children.Add(new TextBlock
            {
                Text = "클릭 삽입",
                FontSize = 10,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase
            });
            hintPanel.Children.Add(new TextBlock
            {
                Text = "  ·  ",
                FontSize = 10,
                Foreground = ThemeResources.TextDim,
                FontFamily = ThemeResources.FontBase
            });
            hintPanel.Children.Add(new TextBlock
            {
                Text = "드래그로 이동",
                FontSize = 10,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase
            });

            stack.Children.Add(new Border
            {
                Background = ThemeResources.BgSurface,
                CornerRadius = new CornerRadius(0, 0, 16, 16),
                Padding = new Thickness(12, 8, 12, 8),
                BorderBrush = ThemeResources.BorderBase,
                BorderThickness = new Thickness(0, 1, 0, 0),
                Child = hintPanel
            });

            outer.Child = stack;
            ThemeResources.ApplyRoundedClip(outer, 16);
            return outer;
        }

        // ── Animation ────────────────────────────────────────────────

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
    }
}

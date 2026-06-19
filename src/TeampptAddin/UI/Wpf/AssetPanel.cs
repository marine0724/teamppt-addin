using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using DrawingImage = System.Drawing.Image;

namespace TeampptAddin
{
    internal class AssetPanel : UserControl
    {
        public event Action<AssetCard> CardClickInsert;
        public event Action<AssetCard> CardDragStart;
        public event Action<StylePalette, StyleFont> StyleApplyRequested;

        private List<HeaderAsset> _allAssets = new List<HeaderAsset>();
        private List<AssetCard> _assetCards = new List<AssetCard>();
        private Dictionary<string, AssetCard> _cardByFile = new Dictionary<string, AssetCard>();
        private IAiService _aiService;
        private StyleConfig _styleConfig;

        private Border[] _tabBorders;
        private FrameworkElement[] _tabPanels;
        private int _activeTab;

        private StackPanel _chatStack;
        private ScrollViewer _chatScroll;
        private TextBox _inputBox;
        private bool _hasPlaceholder;
        private StackPanel _emptyState;

        private StackPanel _assetStack;
        private Border[] _categoryBtns;
        private readonly string[] _categories = { "전체", "헤더", "섹션", "레이아웃", "마무리" };
        private string _activeCategory = "전체";

        private StylePalette _selectedPalette;
        private StyleFont _selectedFont;
        private StackPanel _styleStack;
        private Border[] _paletteBtns;
        private Border[] _fontBtns;

        private TextBlock _statusText;

        public AssetPanel()
        {
            Background = ThemeResources.BgBase;
            FontFamily = ThemeResources.FontBase;

            var root = new DockPanel { LastChildFill = true };

            var header = BuildHeader();
            DockPanel.SetDock(header, Dock.Top);
            root.Children.Add(header);

            var statusBar = BuildStatusBar();
            DockPanel.SetDock(statusBar, Dock.Bottom);
            root.Children.Add(statusBar);

            var aiTab = BuildAiTab();
            var assetTab = BuildAssetTab();
            var styleTab = BuildStyleTab();
            _tabPanels = new FrameworkElement[] { aiTab, assetTab, styleTab };

            var content = new Grid();
            foreach (var p in _tabPanels) content.Children.Add(p);
            root.Children.Add(content);

            Content = root;
            SwitchTab(0);
        }

        // ══════════════════════════════════════════════════════════════
        //  HEADER + TAB BAR
        // ══════════════════════════════════════════════════════════════

        private StackPanel BuildHeader()
        {
            var header = new StackPanel { Background = ThemeResources.BgSurface };

            header.Children.Add(new TextBlock
            {
                Text = "TEAMPPT",
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeResources.Accent,
                FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(16, 12, 0, 8)
            });

            var tabGrid = new Grid { Height = 36 };
            for (int i = 0; i < 3; i++)
                tabGrid.ColumnDefinitions.Add(new ColumnDefinition());

            var sep = new Border
            {
                Height = 1,
                Background = ThemeResources.BorderBase,
                VerticalAlignment = VerticalAlignment.Bottom
            };
            Grid.SetColumnSpan(sep, 3);
            tabGrid.Children.Add(sep);

            string[] labels = { "AI", "에셋", "스타일" };
            _tabBorders = new Border[3];

            for (int i = 0; i < 3; i++)
            {
                var idx = i;

                var indicator = new Border
                {
                    Height = 2,
                    CornerRadius = new CornerRadius(1),
                    Background = Brushes.Transparent,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Margin = new Thickness(12, 0, 12, 0)
                };

                var lbl = new TextBlock
                {
                    Text = labels[i],
                    FontSize = 13,
                    FontFamily = ThemeResources.FontBase,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = ThemeResources.TextSub,
                    Margin = new Thickness(0, 0, 0, 4)
                };

                var inner = new Grid();
                inner.Children.Add(lbl);
                inner.Children.Add(indicator);

                var tab = new Border
                {
                    Child = inner,
                    Background = Brushes.Transparent,
                    Cursor = Cursors.Hand
                };
                tab.Tag = new object[] { lbl, indicator };
                tab.MouseLeftButtonUp += (s, e) => SwitchTab(idx);

                _tabBorders[i] = tab;
                Grid.SetColumn(tab, i);
                tabGrid.Children.Add(tab);
            }

            header.Children.Add(tabGrid);
            return header;
        }

        private void SwitchTab(int index)
        {
            _activeTab = index;
            for (int i = 0; i < _tabBorders.Length; i++)
            {
                var parts = (object[])_tabBorders[i].Tag;
                var lbl = (TextBlock)parts[0];
                var ind = (Border)parts[1];
                bool active = i == index;
                lbl.Foreground = active ? ThemeResources.TextAccent : ThemeResources.TextSub;
                lbl.FontWeight = active ? FontWeights.SemiBold : FontWeights.Normal;
                ind.Background = active ? ThemeResources.Accent : Brushes.Transparent;
            }
            for (int i = 0; i < _tabPanels.Length; i++)
                _tabPanels[i].Visibility = i == index ? Visibility.Visible : Visibility.Collapsed;
        }

        // ══════════════════════════════════════════════════════════════
        //  STATUS BAR
        // ══════════════════════════════════════════════════════════════

        private Border BuildStatusBar()
        {
            _statusText = new TextBlock
            {
                FontSize = 10,
                Foreground = ThemeResources.TextSub,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(14, 0, 0, 0),
                Text = "로딩 중..."
            };
            return new Border
            {
                Height = 24,
                Background = ThemeResources.BgSurface,
                BorderBrush = ThemeResources.BorderBase,
                BorderThickness = new Thickness(0, 1, 0, 0),
                Child = _statusText
            };
        }

        public void SetStatus(string text, Color color)
        {
            _statusText.Text = text;
            _statusText.Foreground = new SolidColorBrush(color);
        }

        public void ResetStatus()
        {
            int count = _assetCards.Count;
            _statusText.Text = count > 0
                ? $"{count}개 에셋 · 드래그 또는 클릭으로 삽입"
                : "Assets 폴더에 header_N.pptx 파일을 넣으세요";
            _statusText.Foreground = ThemeResources.TextSub;
        }

        // ══════════════════════════════════════════════════════════════
        //  AI TAB
        // ══════════════════════════════════════════════════════════════

        private Grid BuildAiTab()
        {
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            _chatStack = new StackPanel { Margin = new Thickness(0, 8, 0, 8) };
            _emptyState = BuildEmptyState();
            _chatStack.Children.Add(_emptyState);

            _chatScroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = _chatStack,
                Background = ThemeResources.BgBase
            };
            Grid.SetRow(_chatScroll, 0);
            grid.Children.Add(_chatScroll);

            var inputBar = BuildInputBar();
            Grid.SetRow(inputBar, 1);
            grid.Children.Add(inputBar);

            return grid;
        }

        private StackPanel BuildEmptyState()
        {
            var panel = new StackPanel
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(20, 48, 20, 20)
            };

            var iconScale = new ScaleTransform(1, 1);
            var iconBorder = new Border
            {
                Width = 52,
                Height = 52,
                CornerRadius = new CornerRadius(26),
                Background = ThemeResources.BgChip,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 12),
                RenderTransform = iconScale,
                RenderTransformOrigin = new Point(0.5, 0.5),
                Child = new TextBlock
                {
                    Text = "✦",
                    FontSize = 22,
                    Foreground = ThemeResources.Accent,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
            panel.Children.Add(iconBorder);

            try
            {
                var pulseAnim = new DoubleAnimation
                {
                    From = 1.0,
                    To = 1.06,
                    Duration = TimeSpan.FromSeconds(1.4),
                    AutoReverse = true,
                    RepeatBehavior = RepeatBehavior.Forever,
                    EasingFunction = new SineEase()
                };
                iconScale.BeginAnimation(ScaleTransform.ScaleXProperty, pulseAnim);
                iconScale.BeginAnimation(ScaleTransform.ScaleYProperty, pulseAnim);
            }
            catch { }

            panel.Children.Add(new TextBlock
            {
                Text = "슬라이드 디자인을\n도와드릴게요",
                FontSize = 15,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 6)
            });

            panel.Children.Add(new TextBlock
            {
                Text = "어떤 슬라이드가 필요한지 말씀해 주세요",
                FontSize = 12,
                Foreground = ThemeResources.TextSub,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 24)
            });

            var chips = new WrapPanel
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Orientation = Orientation.Horizontal
            };
            foreach (var ex in new[] { "회사 소개", "장점 3가지", "서비스 흐름", "팀 소개" })
                chips.Children.Add(BuildChip(ex));
            panel.Children.Add(chips);

            return panel;
        }

        private Border BuildChip(string text)
        {
            var lbl = new TextBlock
            {
                Text = text,
                FontSize = 12,
                Foreground = ThemeResources.TextMain,
                VerticalAlignment = VerticalAlignment.Center
            };

            var chip = new Border
            {
                Child = lbl,
                Background = ThemeResources.BgChip,
                BorderBrush = ThemeResources.BorderCard,
                BorderThickness = new Thickness(1),
                CornerRadius = ThemeResources.RadiusChip,
                Padding = new Thickness(13, 6, 13, 6),
                Margin = new Thickness(4),
                Cursor = Cursors.Hand
            };

            chip.MouseEnter += (s, e) =>
            {
                chip.Background = ThemeResources.BgCategoryActive;
                chip.BorderBrush = ThemeResources.AccentBorder;
                lbl.Foreground = ThemeResources.TextAccent;
            };
            chip.MouseLeave += (s, e) =>
            {
                chip.Background = ThemeResources.BgChip;
                chip.BorderBrush = ThemeResources.BorderCard;
                lbl.Foreground = ThemeResources.TextMain;
            };
            chip.MouseLeftButtonUp += async (s, e) => await SendAiMessage(text);

            return chip;
        }

        private Border BuildInputBar()
        {
            var bar = new Border
            {
                Background = ThemeResources.BgSurface,
                BorderBrush = ThemeResources.BorderBase,
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(12, 10, 12, 10)
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var inputWrap = new Border
            {
                Background = ThemeResources.BgInput,
                BorderBrush = ThemeResources.AccentBorder,
                BorderThickness = new Thickness(1),
                CornerRadius = ThemeResources.RadiusInput,
                Padding = new Thickness(12, 8, 12, 8)
            };

            _inputBox = new TextBox
            {
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                FontSize = 13,
                FontFamily = ThemeResources.FontBase,
                Foreground = ThemeResources.TextDim,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                MinHeight = 40,
                MaxHeight = 80,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "슬라이드에 뭘 넣고 싶어요?"
            };
            _hasPlaceholder = true;

            _inputBox.GotFocus += (s, e) =>
            {
                if (!_hasPlaceholder) return;
                _inputBox.Text = "";
                _inputBox.Foreground = ThemeResources.TextMain;
                _hasPlaceholder = false;
                inputWrap.BorderBrush = ThemeResources.Accent;
            };
            _inputBox.LostFocus += (s, e) =>
            {
                inputWrap.BorderBrush = ThemeResources.AccentBorder;
                if (!string.IsNullOrEmpty(_inputBox.Text)) return;
                _inputBox.Text = "슬라이드에 뭘 넣고 싶어요?";
                _inputBox.Foreground = ThemeResources.TextDim;
                _hasPlaceholder = true;
            };
            _inputBox.PreviewKeyDown += async (s, e) =>
            {
                if (e.Key == Key.Return && Keyboard.Modifiers != ModifierKeys.Shift)
                {
                    e.Handled = true;
                    await SendAiMessageFromInput();
                }
            };

            inputWrap.Child = _inputBox;
            Grid.SetColumn(inputWrap, 0);
            grid.Children.Add(inputWrap);

            var sendBtn = new Border
            {
                Background = ThemeResources.Accent,
                CornerRadius = new CornerRadius(9),
                Width = 32,
                Height = 32,
                Margin = new Thickness(8, 0, 0, 0),
                Cursor = Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Bottom,
                Child = new TextBlock
                {
                    Text = "↑",
                    FontSize = 16,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
            sendBtn.MouseLeftButtonUp += async (s, e) => await SendAiMessageFromInput();
            Grid.SetColumn(sendBtn, 1);
            grid.Children.Add(sendBtn);

            bar.Child = grid;
            return bar;
        }

        // ── AI messaging ─────────────────────────────────────────────

        private async Task SendAiMessageFromInput()
        {
            if (_hasPlaceholder || string.IsNullOrWhiteSpace(_inputBox.Text)) return;
            var intent = _inputBox.Text.Trim();
            _inputBox.Text = "";
            await SendAiMessage(intent);
        }

        private async Task SendAiMessage(string intent)
        {
            if (_emptyState != null && _emptyState.Visibility == Visibility.Visible)
            {
                try
                {
                    var fadeOut = new DoubleAnimation
                    {
                        From = 1,
                        To = 0,
                        Duration = TimeSpan.FromMilliseconds(150)
                    };
                    fadeOut.Completed += (s, e) => _emptyState.Visibility = Visibility.Collapsed;
                    _emptyState.BeginAnimation(OpacityProperty, fadeOut);
                }
                catch { _emptyState.Visibility = Visibility.Collapsed; }
            }

            AddUserBubble(intent);

            if (_aiService == null)
            {
                AddAiBubble("AI 서비스를 초기화 중입니다.");
                return;
            }

            var loading = AddAiLoadingBubble();

            try
            {
                var rec = await _aiService.RecommendAsync(
                    intent,
                    _allAssets,
                    _styleConfig?.Palettes ?? new List<StylePalette>(),
                    _styleConfig?.Fonts ?? new List<StyleFont>());

                _chatStack.Children.Remove(loading);
                ShowAiResponse(rec);
            }
            catch (Exception ex)
            {
                _chatStack.Children.Remove(loading);
                AddAiBubble($"오류가 발생했습니다: {ex.Message}");
            }

            _chatScroll.ScrollToBottom();
        }

        private bool _isFirstMessage = true;

        private void AddUserBubble(string text)
        {
            var bubble = new Border
            {
                Background = ThemeResources.BgUserBubble,
                CornerRadius = new CornerRadius(14, 4, 14, 14),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(40, 4, 12, 4),
                HorizontalAlignment = HorizontalAlignment.Right,
                MaxWidth = 220,
                Child = new TextBlock
                {
                    Text = text,
                    Foreground = Brushes.White,
                    FontSize = 12,
                    TextWrapping = TextWrapping.Wrap,
                    FontFamily = ThemeResources.FontBase
                }
            };

            if (_isFirstMessage)
            {
                _isFirstMessage = false;
                bubble.Opacity = 0;
                _chatStack.Children.Add(bubble);
                try
                {
                    var fadeIn = new DoubleAnimation
                    {
                        From = 0,
                        To = 1,
                        Duration = TimeSpan.FromMilliseconds(200),
                        BeginTime = TimeSpan.FromMilliseconds(50)
                    };
                    bubble.BeginAnimation(OpacityProperty, fadeIn);
                }
                catch { bubble.Opacity = 1; }
            }
            else
            {
                _chatStack.Children.Add(bubble);
            }
        }

        private FrameworkElement AddAiLoadingBubble()
        {
            var wrapper = new StackPanel { Margin = new Thickness(12, 4, 40, 4) };

            wrapper.Children.Add(new TextBlock
            {
                Text = "TEAMPPT AI",
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.Accent,
                Margin = new Thickness(4, 0, 0, 3)
            });

            var dotsPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center
            };

            for (int i = 0; i < 3; i++)
            {
                var idx = i;
                var dot = new Ellipse
                {
                    Width = 6,
                    Height = 6,
                    Fill = ThemeResources.BgThumb,
                    Margin = new Thickness(i == 0 ? 0 : 4, 0, 0, 0),
                    RenderTransform = new TranslateTransform(0, 0)
                };

                dot.Loaded += (s, e) =>
                {
                    try
                    {
                        var bounce = new DoubleAnimation
                        {
                            From = 0,
                            To = -5,
                            Duration = TimeSpan.FromSeconds(0.6),
                            AutoReverse = true,
                            RepeatBehavior = RepeatBehavior.Forever,
                            EasingFunction = new SineEase(),
                            BeginTime = TimeSpan.FromMilliseconds(idx * 200)
                        };
                        ((TranslateTransform)((Ellipse)s).RenderTransform).BeginAnimation(TranslateTransform.YProperty, bounce);
                    }
                    catch { }
                };

                dotsPanel.Children.Add(dot);
            }

            wrapper.Children.Add(new Border
            {
                Background = ThemeResources.BgAiResponse,
                CornerRadius = new CornerRadius(4, 13, 13, 13),
                Padding = new Thickness(14, 12, 14, 12),
                HorizontalAlignment = HorizontalAlignment.Left,
                Child = dotsPanel
            });

            _chatStack.Children.Add(wrapper);
            return wrapper;
        }

        private void AddAiBubble(string text)
        {
            var wrapper = new StackPanel
            {
                Margin = new Thickness(12, 4, 40, 4),
                Opacity = 0
            };

            wrapper.Children.Add(new TextBlock
            {
                Text = "TEAMPPT AI",
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.Accent,
                Margin = new Thickness(4, 0, 0, 3)
            });

            var tb = new TextBlock
            {
                Text = "",
                Foreground = ThemeResources.TextMain,
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
                FontFamily = ThemeResources.FontBase
            };

            wrapper.Children.Add(new Border
            {
                Background = ThemeResources.BgAiResponse,
                CornerRadius = new CornerRadius(4, 13, 13, 13),
                Padding = new Thickness(12, 8, 12, 8),
                HorizontalAlignment = HorizontalAlignment.Left,
                Child = tb
            });

            _chatStack.Children.Add(wrapper);
            wrapper.Opacity = 1;

            try
            {
                var rng = new Random();
                int charIdx = 0;
                tb.Text = "|";

                var typeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(30) };
                typeTimer.Tick += (s, e) =>
                {
                    try
                    {
                        typeTimer.Interval = TimeSpan.FromMilliseconds(30 + rng.Next(16));
                        if (charIdx < text.Length)
                        {
                            tb.Text = text.Substring(0, charIdx + 1) + "|";
                            charIdx++;
                            _chatScroll.ScrollToBottom();
                        }
                        else
                        {
                            typeTimer.Stop();
                            int blinks = 0;
                            var cursorTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                            cursorTimer.Tick += (s2, e2) =>
                            {
                                blinks++;
                                tb.Text = blinks % 2 == 0 ? text + "|" : text;
                                if (blinks >= 2)
                                {
                                    cursorTimer.Stop();
                                    tb.Text = text;
                                }
                            };
                            cursorTimer.Start();
                        }
                    }
                    catch { typeTimer.Stop(); tb.Text = text; }
                };
                typeTimer.Start();
            }
            catch { tb.Text = text; }
        }

        private void ShowAiResponse(AiRecommendation rec)
        {
            AddAiBubble(rec.Message ?? "추천 에셋을 확인해보세요.");

            if (rec.Assets == null || rec.Assets.Count == 0) return;

            _chatStack.Children.Add(new TextBlock
            {
                Text = "추천 에셋",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextSub,
                Margin = new Thickness(12, 8, 0, 2)
            });

            for (int i = 0; i < rec.Assets.Count; i++)
            {
                var card = BuildAiAssetCard(rec.Assets[i]);
                card.Opacity = 0;
                card.RenderTransform = new TranslateTransform(0, 12);
                _chatStack.Children.Add(card);

                try
                {
                    var fadeIn = new DoubleAnimation
                    {
                        From = 0,
                        To = 1,
                        Duration = TimeSpan.FromMilliseconds(300),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut },
                        BeginTime = TimeSpan.FromMilliseconds(i * 80)
                    };

                    var slideUp = new DoubleAnimation
                    {
                        From = 12,
                        To = 0,
                        Duration = TimeSpan.FromMilliseconds(300),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut },
                        BeginTime = TimeSpan.FromMilliseconds(i * 80)
                    };

                    var sb = new Storyboard();
                    Storyboard.SetTarget(fadeIn, card);
                    Storyboard.SetTargetProperty(fadeIn, new PropertyPath(OpacityProperty));
                    Storyboard.SetTarget(slideUp, card);
                    Storyboard.SetTargetProperty(slideUp,
                        new PropertyPath("(UIElement.RenderTransform).(TranslateTransform.Y)"));
                    sb.Children.Add(fadeIn);
                    sb.Children.Add(slideUp);
                    sb.Begin();
                }
                catch { card.Opacity = 1; card.RenderTransform = null; }
            }
        }

        private Border BuildAiAssetCard(AssetSuggestion suggestion)
        {
            _cardByFile.TryGetValue(suggestion.Asset?.File ?? "", out var realCard);

            var thumbBorder = new Border
            {
                Width = 66,
                Height = 48,
                Background = ThemeResources.BgThumb,
                CornerRadius = new CornerRadius(8),
                ClipToBounds = true,
                Margin = new Thickness(0, 0, 10, 0)
            };
            ThemeResources.ApplyRoundedClip(thumbBorder, 8);
            if (realCard?.DrawingThumbnail != null)
            {
                thumbBorder.Child = new Image
                {
                    Source = ConvertToBitmapSource(realCard.DrawingThumbnail),
                    Stretch = Stretch.Uniform,
                    Margin = new Thickness(4)
                };
            }

            var info = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            info.Children.Add(new TextBlock
            {
                Text = suggestion.Asset?.Name ?? "에셋",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain,
                FontFamily = ThemeResources.FontBase,
                TextTrimming = TextTrimming.CharacterEllipsis
            });
            info.Children.Add(new TextBlock
            {
                Text = suggestion.Reason ?? "",
                FontSize = 9.5,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(0, 2, 0, 0)
            });

            var badge = new Border
            {
                Background = ThemeResources.BgBadge,
                CornerRadius = ThemeResources.RadiusBadge,
                Padding = new Thickness(6, 2, 6, 2),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0, 0, 0),
                Visibility = realCard != null ? Visibility.Visible : Visibility.Collapsed,
                Child = new TextBlock
                {
                    Text = "DRAG",
                    Foreground = ThemeResources.TextAccent,
                    FontSize = 8,
                    FontWeight = FontWeights.Bold,
                    FontFamily = ThemeResources.FontBase
                }
            };

            var row = new Grid();
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(66) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(thumbBorder, 0);
            Grid.SetColumn(info, 1);
            Grid.SetColumn(badge, 2);
            row.Children.Add(thumbBorder);
            row.Children.Add(info);
            row.Children.Add(badge);

            var card = new Border
            {
                Child = row,
                Background = ThemeResources.BgCard,
                BorderBrush = ThemeResources.BorderCard,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(11),
                ClipToBounds = true,
                Padding = new Thickness(8),
                Margin = new Thickness(12, 3, 12, 3),
                Cursor = realCard != null ? Cursors.Hand : Cursors.Arrow
            };
            ThemeResources.ApplyRoundedClip(card, 11);

            if (realCard != null)
            {
                bool mouseDown = false;
                Point dragOrigin = default;

                card.MouseEnter += (s, e) =>
                {
                    card.Background = ThemeResources.BgCardHover;
                    card.BorderBrush = ThemeResources.BorderCardHover;
                };
                card.MouseLeave += (s, e) =>
                {
                    card.Background = ThemeResources.BgCard;
                    card.BorderBrush = ThemeResources.BorderCard;
                };
                card.MouseLeftButtonDown += (s, e) =>
                {
                    mouseDown = true;
                    dragOrigin = e.GetPosition(card);
                };
                card.MouseMove += (s, e) =>
                {
                    if (!mouseDown || e.LeftButton != MouseButtonState.Pressed) return;
                    var pos = e.GetPosition(card);
                    if (Math.Abs(pos.X - dragOrigin.X) > SystemParameters.MinimumHorizontalDragDistance ||
                        Math.Abs(pos.Y - dragOrigin.Y) > SystemParameters.MinimumVerticalDragDistance)
                    {
                        mouseDown = false;
                        CardDragStart?.Invoke(realCard);
                    }
                };
                card.MouseLeftButtonUp += (s, e) =>
                {
                    if (!mouseDown) return;
                    mouseDown = false;
                    CardClickInsert?.Invoke(realCard);
                };
            }

            return card;
        }

        // ══════════════════════════════════════════════════════════════
        //  ASSET TAB
        // ══════════════════════════════════════════════════════════════

        private DockPanel BuildAssetTab()
        {
            var dock = new DockPanel { LastChildFill = true };

            var filterBar = BuildCategoryBar();
            DockPanel.SetDock(filterBar, Dock.Top);
            dock.Children.Add(filterBar);

            _assetStack = new StackPanel { Margin = new Thickness(0, 4, 0, 8) };
            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = _assetStack,
                Background = ThemeResources.BgBase
            };
            dock.Children.Add(scroll);

            return dock;
        }

        private Border BuildCategoryBar()
        {
            var chips = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(8, 6, 8, 6)
            };
            _categoryBtns = new Border[_categories.Length];

            for (int i = 0; i < _categories.Length; i++)
            {
                var cat = _categories[i];
                var lbl = new TextBlock
                {
                    Text = cat,
                    FontSize = 11,
                    FontFamily = ThemeResources.FontBase,
                    VerticalAlignment = VerticalAlignment.Center
                };
                var chip = new Border
                {
                    Child = lbl,
                    CornerRadius = ThemeResources.RadiusChip,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(10, 4, 10, 4),
                    Margin = new Thickness(2, 0, 2, 0),
                    Cursor = Cursors.Hand
                };
                chip.Tag = lbl;
                chip.MouseLeftButtonUp += (s, e) => FilterCategory(cat);
                _categoryBtns[i] = chip;
                chips.Children.Add(chip);
            }

            var scroll = new ScrollViewer
            {
                Content = chips,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
            };

            ApplyCategoryStyles();

            return new Border
            {
                Child = scroll,
                Background = ThemeResources.BgSurface,
                BorderBrush = ThemeResources.BorderBase,
                BorderThickness = new Thickness(0, 0, 0, 1)
            };
        }

        private void FilterCategory(string category)
        {
            _activeCategory = category;
            ApplyCategoryStyles();

            foreach (var card in _assetStack.Children.OfType<AssetCard>())
            {
                var asset = card.Tag as HeaderAsset;
                card.Visibility = (category == "전체" || asset?.Category == category)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }
        }

        private void ApplyCategoryStyles()
        {
            for (int i = 0; i < _categories.Length; i++)
            {
                bool active = _categories[i] == _activeCategory;
                var chip = _categoryBtns[i];
                var lbl = (TextBlock)chip.Tag;
                chip.Background = active ? ThemeResources.BgCategoryActive : Brushes.Transparent;
                chip.BorderBrush = active ? ThemeResources.AccentBorder : Brushes.Transparent;
                lbl.Foreground = active ? ThemeResources.TextAccent : ThemeResources.TextSub;
                lbl.FontWeight = active ? FontWeights.SemiBold : FontWeights.Normal;
            }
        }

        // ══════════════════════════════════════════════════════════════
        //  STYLE TAB
        // ══════════════════════════════════════════════════════════════

        private FrameworkElement BuildStyleTab()
        {
            _styleStack = new StackPanel { Margin = new Thickness(0, 4, 0, 12) };
            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = _styleStack,
                Background = ThemeResources.BgBase
            };

            return new Border { Background = ThemeResources.BgBase, Child = scroll };
        }

        private void PopulateStylePanel()
        {
            if (_styleConfig == null || _styleStack == null) return;
            _styleStack.Children.Clear();

            var palettes = _styleConfig.Palettes ?? new List<StylePalette>();
            var fonts    = _styleConfig.Fonts    ?? new List<StyleFont>();

            _selectedPalette = palettes.Count > 0 ? palettes[0] : null;
            _selectedFont    = fonts.Count    > 0 ? fonts[0]    : null;

            _styleStack.Children.Add(BuildSectionLabel("컬러 팔레트"));

            var paletteWrap = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(12, 0, 12, 0)
            };
            _paletteBtns = new Border[palettes.Count];
            for (int i = 0; i < palettes.Count; i++)
            {
                var idx = i;
                var swatch = BuildPaletteSwatch(palettes[i]);
                swatch.MouseLeftButtonUp += (s, e) =>
                {
                    _selectedPalette = palettes[idx];
                    RefreshPaletteSelection(idx);
                    StyleApplyRequested?.Invoke(_selectedPalette, _selectedFont);
                };
                _paletteBtns[i] = swatch;
                paletteWrap.Children.Add(swatch);
            }
            _styleStack.Children.Add(paletteWrap);
            if (palettes.Count > 0) RefreshPaletteSelection(0);

            _styleStack.Children.Add(BuildSectionLabel("폰트"));

            var fontPanel = new StackPanel { Margin = new Thickness(12, 0, 12, 0) };
            _fontBtns = new Border[fonts.Count];
            for (int i = 0; i < fonts.Count; i++)
            {
                var idx = i;
                var row = BuildFontRow(fonts[i]);
                row.MouseLeftButtonUp += (s, e) =>
                {
                    _selectedFont = fonts[idx];
                    RefreshFontSelection(idx);
                    StyleApplyRequested?.Invoke(_selectedPalette, _selectedFont);
                };
                _fontBtns[i] = row;
                fontPanel.Children.Add(row);
            }
            _styleStack.Children.Add(fontPanel);
            if (fonts.Count > 0) RefreshFontSelection(0);
        }

        private static TextBlock BuildSectionLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(14, 16, 0, 8)
            };
        }

        private static Border BuildPaletteSwatch(StylePalette p)
        {
            var colorGrid = new UniformGrid { Rows = 2, Columns = 2, Height = 56 };
            var hexColors = new[] { p.Colors?.Main, p.Colors?.Sub1, p.Colors?.Sub2, p.Colors?.Text };
            foreach (var hex in hexColors)
            {
                colorGrid.Children.Add(new System.Windows.Shapes.Rectangle
                {
                    Fill = BrushFromHex(hex ?? "#CCCCCC")
                });
            }

            var colorStrip = new Border
            {
                ClipToBounds = true,
                Child = colorGrid
            };

            var nameText = new TextBlock
            {
                Text = p.Name,
                FontSize = 9,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase,
                HorizontalAlignment = HorizontalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Padding = new Thickness(4, 3, 4, 5)
            };

            var content = new StackPanel();
            content.Children.Add(colorStrip);
            content.Children.Add(nameText);

            var check = new Border
            {
                Background = ThemeResources.Accent,
                CornerRadius = new CornerRadius(99),
                Width = 16, Height = 16,
                Visibility = Visibility.Collapsed,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 4, 4, 0),
                Child = new TextBlock
                {
                    Text = "✓",
                    FontSize = 9,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };

            var overlay = new Grid();
            overlay.Children.Add(content);
            overlay.Children.Add(check);

            var card = new Border
            {
                Width = 80,
                CornerRadius = new CornerRadius(12),
                BorderBrush = ThemeResources.BorderCard,
                BorderThickness = new Thickness(1.5),
                Margin = new Thickness(4),
                Cursor = Cursors.Hand,
                ClipToBounds = true,
                Child = overlay
            };
            card.Tag = check;
            ThemeResources.ApplyRoundedClip(card, 12);

            card.MouseEnter += (s, e) =>
            {
                if (card.BorderBrush != ThemeResources.Accent)
                    card.BorderBrush = ThemeResources.BorderCardHover;
            };
            card.MouseLeave += (s, e) =>
            {
                if (card.BorderBrush != ThemeResources.Accent)
                    card.BorderBrush = ThemeResources.BorderCard;
            };

            return card;
        }

        private static Border BuildFontRow(StyleFont f)
        {
            var nameText = new TextBlock
            {
                Text = f.Name,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center
            };

            var sampleText = new TextBlock
            {
                Text = "가나다라",
                FontSize = 12,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 0, 0)
            };

            var leftPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center
            };
            leftPanel.Children.Add(nameText);
            leftPanel.Children.Add(sampleText);

            var check = new Border
            {
                Background = ThemeResources.Accent,
                CornerRadius = new CornerRadius(99),
                Width = 18, Height = 18,
                Visibility = Visibility.Collapsed,
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = "✓",
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(leftPanel, 0);
            Grid.SetColumn(check, 1);
            grid.Children.Add(leftPanel);
            grid.Children.Add(check);

            var row = new Border
            {
                CornerRadius = new CornerRadius(12),
                ClipToBounds = true,
                BorderBrush = ThemeResources.BorderCard,
                BorderThickness = new Thickness(1.5),
                Padding = new Thickness(14, 11, 14, 11),
                Margin = new Thickness(0, 0, 0, 5),
                Cursor = Cursors.Hand,
                Background = Brushes.Transparent,
                Child = grid
            };
            row.Tag = check;
            ThemeResources.ApplyRoundedClip(row, 12);

            row.MouseEnter += (s, e) =>
            {
                if (row.Background != ThemeResources.BgCategoryActive)
                    row.Background = ThemeResources.BgChip;
            };
            row.MouseLeave += (s, e) =>
            {
                if (row.Background != ThemeResources.BgCategoryActive)
                    row.Background = Brushes.Transparent;
            };

            return row;
        }

        private void RefreshPaletteSelection(int selectedIdx)
        {
            for (int i = 0; i < _paletteBtns.Length; i++)
            {
                var btn   = _paletteBtns[i];
                var check = (Border)btn.Tag;
                bool active = i == selectedIdx;
                btn.BorderBrush = active ? ThemeResources.Accent : ThemeResources.BorderCard;
                if (check != null)
                    check.Visibility = active ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void RefreshFontSelection(int selectedIdx)
        {
            for (int i = 0; i < _fontBtns.Length; i++)
            {
                var row   = _fontBtns[i];
                var check = (Border)row.Tag;
                bool active = i == selectedIdx;
                row.Background  = active ? ThemeResources.BgCategoryActive : Brushes.Transparent;
                row.BorderBrush = active ? ThemeResources.Accent : ThemeResources.BorderCard;
                if (check != null)
                    check.Visibility = active ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private static SolidColorBrush BrushFromHex(string hex)
        {
            hex = hex.TrimStart('#');
            byte r = Convert.ToByte(hex.Substring(0, 2), 16);
            byte g = Convert.ToByte(hex.Substring(2, 2), 16);
            byte b = Convert.ToByte(hex.Substring(4, 2), 16);
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        // ══════════════════════════════════════════════════════════════
        //  PUBLIC API (TaskPaneHost 호출용)
        // ══════════════════════════════════════════════════════════════

        public void SetAssets(List<HeaderAsset> assets)
        {
            _allAssets = assets ?? new List<HeaderAsset>();
        }

        public void InitAi(IAiService aiService, StyleConfig styles)
        {
            _aiService = aiService;
            _styleConfig = styles;
            PopulateStylePanel();
        }

        public void AddAssetCard(AssetCard card, HeaderAsset asset)
        {
            card.Tag = asset;
            card.ClickInsertRequested += c => CardClickInsert?.Invoke(c);
            card.DragStartRequested += c => CardDragStart?.Invoke(c);

            _cardByFile[asset.File] = card;
            _assetCards.Add(card);
            _assetStack?.Children.Add(card);
        }

        // ══════════════════════════════════════════════════════════════
        //  UTILS
        // ══════════════════════════════════════════════════════════════

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

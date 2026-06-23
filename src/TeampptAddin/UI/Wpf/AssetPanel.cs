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
        private RemoteAssetCache _remoteCache;
        private HeaderAsset _anchorAsset;

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
        private Border _ingestButton;

        private Border _ingestProgressTrack;
        private Border _ingestProgressFill;
        private TextBlock _ingestProgressText;
        private Grid _ingestCurrentContainer;
        private TextBlock _ingestStageText;
        private StackPanel _ingestCompletedStack;
        private Border _completedCollapsedBox;
        private TextBlock _completedCountText;
        private TextBlock _completedChevron;
        private Image _completedLatestThumb;
        private TextBlock _completedLatestName;
        private bool _completedExpanded;
        private int _completedCount;
        private DispatcherTimer _ingestDotsTimer;
        private string _ingestStageBase;
        private Border _ingestScanLine;
        private TextBlock _ingestNameText;
        private int _ingestLastIndex;
        private string _retryOutputDir;
        private string _retryBundleName;
        private int _retryStartFrom;

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

            var titleRow = new Grid();
            titleRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            titleRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            titleRow.Children.Add(new TextBlock
            {
                Text = "TEAMPPT",
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeResources.Accent,
                FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(16, 12, 0, 8)
            });

            _ingestButton = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(234, 88, 12)),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 4, 10, 4),
                Margin = new Thickness(0, 10, 12, 6),
                Cursor = Cursors.Hand,
                Visibility = Visibility.Collapsed,
                Child = new TextBlock
                {
                    Text = "인제스트",
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brushes.White,
                    FontFamily = ThemeResources.FontBase
                }
            };
            _ingestButton.MouseLeftButtonUp += async (s, e) => await RunIngestAsync();
            Grid.SetColumn(_ingestButton, 1);
            titleRow.Children.Add(_ingestButton);
            header.Children.Add(titleRow);

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

                RemoveLoadingBubble(loading);
                ShowAiResponse(rec);
            }
            catch (Exception ex)
            {
                RemoveLoadingBubble(loading);
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

            var dotsTb = new TextBlock
            {
                Text = "·",
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase
            };

            var dotTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            int dotCount = 1;
            dotTimer.Tick += (s, e) =>
            {
                dotCount = dotCount % 3 + 1;
                dotsTb.Text = new string('·', dotCount);
            };
            dotTimer.Start();

            var border = new Border
            {
                Background = ThemeResources.BgAiResponse,
                CornerRadius = new CornerRadius(4, 13, 13, 13),
                Padding = new Thickness(14, 10, 14, 10),
                HorizontalAlignment = HorizontalAlignment.Left,
                MinWidth = 50,
                Child = dotsTb,
                Tag = dotTimer
            };

            wrapper.Children.Add(border);

            _chatStack.Children.Add(wrapper);
            return wrapper;
        }

        private void RemoveLoadingBubble(FrameworkElement loading)
        {
            if (loading is StackPanel sp)
            {
                foreach (var child in sp.Children)
                {
                    if (child is Border b && b.Tag is DispatcherTimer timer)
                        timer.Stop();
                }
            }
            _chatStack.Children.Remove(loading);
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

            int charIdx = 0;
            var typeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(12) };
            typeTimer.Tick += (s, e) =>
            {
                int charsPerTick = Math.Min(3, text.Length - charIdx);
                charIdx += charsPerTick;
                tb.Text = text.Substring(0, charIdx);
                _chatScroll.ScrollToBottom();
                if (charIdx >= text.Length)
                    typeTimer.Stop();
            };
            typeTimer.Start();
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
            var isRemote = realCard == null && suggestion.Asset?.Extra != null
                && suggestion.Asset.Extra.ContainsKey("remote_thumb");

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
            else if (isRemote && _remoteCache != null)
            {
                var remotePptx = suggestion.Asset.Extra["remote_file"].ToString();
                LoadRemoteThumbAsync(thumbBorder, remotePptx);
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
                Visibility = (realCard != null || isRemote) ? Visibility.Visible : Visibility.Collapsed,
                Child = new TextBlock
                {
                    Text = isRemote ? "INSERT" : "DRAG",
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
                Cursor = (realCard != null || isRemote) ? Cursors.Hand : Cursors.Arrow
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
            else if (isRemote && _remoteCache != null)
            {
                var remoteFile = suggestion.Asset.Extra["remote_file"].ToString();
                var assetName = suggestion.Asset?.Name ?? "에셋";
                var category = suggestion.Asset?.Category ?? "";
                var useWhen = suggestion.Reason ?? suggestion.Asset?.UseWhen ?? "";

                bool mouseDown = false;
                Point dragOrigin = default;
                Popup remotePopup = null;
                var popupTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
                var closeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };

                popupTimer.Tick += (s2, e2) =>
                {
                    popupTimer.Stop();
                    if (remotePopup != null) return;
                    var popupContent = BuildRemotePopupContent(assetName, category, useWhen, thumbBorder);
                    popupContent.Opacity = 0;
                    popupContent.RenderTransform = new TranslateTransform(6, 0);
                    remotePopup = new Popup
                    {
                        Child = popupContent,
                        PlacementTarget = card,
                        Placement = PlacementMode.Left,
                        AllowsTransparency = true,
                        StaysOpen = true,
                        IsHitTestVisible = false,
                        IsOpen = true
                    };
                    var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(150));
                    var slideIn = new DoubleAnimation(6, 0, TimeSpan.FromMilliseconds(150));
                    popupContent.BeginAnimation(OpacityProperty, fadeIn);
                    ((TranslateTransform)popupContent.RenderTransform).BeginAnimation(TranslateTransform.XProperty, slideIn);
                };

                closeTimer.Tick += (s2, e2) =>
                {
                    closeTimer.Stop();
                    if (remotePopup != null) { remotePopup.IsOpen = false; remotePopup = null; }
                };

                card.MouseEnter += (s, e) =>
                {
                    card.Background = ThemeResources.BgCardHover;
                    card.BorderBrush = ThemeResources.BorderCardHover;
                    closeTimer.Stop();
                    if (remotePopup == null) popupTimer.Start();
                };
                card.MouseLeave += (s, e) =>
                {
                    card.Background = ThemeResources.BgCard;
                    card.BorderBrush = ThemeResources.BorderCard;
                    popupTimer.Stop();
                    if (remotePopup != null) closeTimer.Start();
                };
                card.MouseLeftButtonDown += (s, e) =>
                {
                    mouseDown = true;
                    dragOrigin = e.GetPosition(card);
                    popupTimer.Stop();
                    if (remotePopup != null) { remotePopup.IsOpen = false; remotePopup = null; }
                };
                var capturedAsset = suggestion.Asset;
                card.MouseMove += (s, e) =>
                {
                    if (!mouseDown || e.LeftButton != MouseButtonState.Pressed) return;
                    var pos = e.GetPosition(card);
                    if (Math.Abs(pos.X - dragOrigin.X) > SystemParameters.MinimumHorizontalDragDistance ||
                        Math.Abs(pos.Y - dragOrigin.Y) > SystemParameters.MinimumVerticalDragDistance)
                    {
                        mouseDown = false;
                        DragRemoteAssetAsync(remoteFile, assetName, capturedAsset);
                    }
                };
                card.MouseLeftButtonUp += (s, e) =>
                {
                    if (!mouseDown) return;
                    mouseDown = false;
                    InsertRemoteAssetAsync(remoteFile, assetName, capturedAsset);
                };
            }

            return card;
        }

        private async void LoadRemoteThumbAsync(Border thumbBorder, string remotePptxPath)
        {
            try
            {
                var localPptx = await _remoteCache.GetPptxAsync(remotePptxPath).ConfigureAwait(false);
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        var thumbCachePath = System.IO.Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "TeampptAddin", "cache", "thumb-shape",
                            System.IO.Path.GetFileNameWithoutExtension(localPptx) + ".png");
                        var thumbImg = ThumbnailService.LoadThumbnail(localPptx, thumbCachePath);
                        if (thumbImg != null)
                        {
                            var bmpSrc = ConvertToBitmapSource(thumbImg);
                            thumbBorder.Child = new Image
                            {
                                Source = bmpSrc,
                                Stretch = Stretch.Uniform,
                                Margin = new Thickness(4)
                            };
                        }
                    }
                    catch (Exception ex) { Logger.Log($"[RemoteThumb] Shape-only 생성 실패: {ex.Message}"); }
                });
            }
            catch (Exception ex) { Logger.Log($"[RemoteThumb] pptx 다운로드 실패: {ex.Message}"); }
        }

        private Border BuildRemotePopupContent(string name, string category, string useWhen, Border thumbSource)
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
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    Color = Color.FromRgb(0x19, 0x1F, 0x28),
                    BlurRadius = 24, ShadowDepth = 4, Opacity = 0.10, Direction = 270
                }
            };

            var stack = new StackPanel();

            var thumbArea = new Border
            {
                Height = 180, ClipToBounds = true,
                CornerRadius = new CornerRadius(16, 16, 0, 0),
                Background = ThemeResources.BgThumb
            };
            ThemeResources.ApplyRoundedClip(thumbArea, new CornerRadius(16, 16, 0, 0));
            var srcImage = thumbSource.Child as Image;
            if (srcImage?.Source != null)
            {
                thumbArea.Child = new Image { Source = srcImage.Source, Stretch = Stretch.Uniform };
            }
            stack.Children.Add(thumbArea);

            stack.Children.Add(new Border { Height = 1, Background = ThemeResources.BorderBase });

            var metaGrid = new Grid();
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            var nameText = new TextBlock
            {
                Text = name, FontSize = 13, FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain, FontFamily = ThemeResources.FontBase,
                TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(nameText, 0);
            metaGrid.Children.Add(nameText);
            if (!string.IsNullOrEmpty(category))
            {
                var catBadge = new Border
                {
                    Background = ThemeResources.BgBadge, CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(6, 2, 6, 2), Margin = new Thickness(6, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = category, FontSize = 10, FontWeight = FontWeights.SemiBold,
                        Foreground = ThemeResources.TextAccent, FontFamily = ThemeResources.FontBase
                    }
                };
                Grid.SetColumn(catBadge, 1);
                metaGrid.Children.Add(catBadge);
            }
            stack.Children.Add(new Border { Padding = new Thickness(12, 10, 12, 8), Child = metaGrid });

            if (!string.IsNullOrEmpty(useWhen))
            {
                stack.Children.Add(new Border
                {
                    Padding = new Thickness(12, 0, 12, 10),
                    Child = new TextBlock
                    {
                        Text = useWhen, FontSize = 11, Foreground = ThemeResources.TextSub,
                        FontFamily = ThemeResources.FontBase, TextWrapping = TextWrapping.Wrap, LineHeight = 16
                    }
                });
            }

            var hintPanel = new StackPanel { Orientation = Orientation.Horizontal };
            hintPanel.Children.Add(new TextBlock { Text = "클릭 삽입", FontSize = 10, Foreground = ThemeResources.TextSub, FontFamily = ThemeResources.FontBase });
            hintPanel.Children.Add(new TextBlock { Text = "  ·  ", FontSize = 10, Foreground = ThemeResources.TextDim, FontFamily = ThemeResources.FontBase });
            hintPanel.Children.Add(new TextBlock { Text = "드래그로 이동", FontSize = 10, Foreground = ThemeResources.TextSub, FontFamily = ThemeResources.FontBase });
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

        private void EnsureAssetRegistered(HeaderAsset asset)
        {
            if (asset == null) return;
            if (_allAssets == null) _allAssets = new List<HeaderAsset>();
            var key = System.IO.Path.GetFileName(asset.File ?? "");
            var remoteKey = asset.Extra != null && asset.Extra.TryGetValue("remote_file", out var rf)
                ? System.IO.Path.GetFileName(rf.ToString()) : null;
            bool exists = _allAssets.Any(a =>
                (!string.IsNullOrEmpty(key) && string.Equals(System.IO.Path.GetFileName(a.File ?? ""), key, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrEmpty(remoteKey) && a.Extra != null && a.Extra.TryGetValue("remote_file", out var arf)
                    && string.Equals(System.IO.Path.GetFileName(arf.ToString()), remoteKey, StringComparison.OrdinalIgnoreCase)));
            if (!exists)
            {
                _allAssets.Add(asset);
                Logger.Log($"[Style] EnsureAssetRegistered added '{asset.Name}', allAssets={_allAssets.Count}");
            }
        }

        private async void DragRemoteAssetAsync(string remotePptx, string assetName, HeaderAsset sourceAsset = null)
        {
            try
            {
                SetStatus($"⬇ {assetName} 다운로드 중...", ThemeResources.TextSub.Color);
                var localPptx = await _remoteCache.GetPptxAsync(remotePptx).ConfigureAwait(false);
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        EnsureAssetRegistered(sourceAsset);
                        var thumbCachePath = System.IO.Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "TeampptAddin", "cache", "thumb-shape",
                            System.IO.Path.GetFileNameWithoutExtension(localPptx) + ".png");
                        var thumbImg = ThumbnailService.LoadThumbnail(localPptx, thumbCachePath);
                        var tempCard = new AssetCard(thumbImg, assetName, localPptx);
                        CardDragStart?.Invoke(tempCard);
                    }
                    catch (Exception ex)
                    {
                        SetStatus($"드래그 실패: {ex.Message}", ThemeResources.StatusError.Color);
                        Logger.Log($"[RemoteDrag] 실패: {ex}");
                    }
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                    SetStatus($"다운로드 실패: {ex.Message}", ThemeResources.StatusError.Color));
                Logger.Log($"[RemoteDrag] 다운로드 실패: {ex}");
            }
        }

        private async void InsertRemoteAssetAsync(string remotePath, string assetName, HeaderAsset sourceAsset = null)
        {
            try
            {
                SetStatus($"⬇ {assetName} 다운로드 중...", ThemeResources.TextSub.Color);
                var localPath = await _remoteCache.GetPptxAsync(remotePath).ConfigureAwait(false);
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        EnsureAssetRegistered(sourceAsset);
                        ShapeInserter.InsertToActiveSlide(localPath);
                        SetStyleAnchorByFile(remotePath);
                        SetStatus($"✓ {assetName} 삽입 완료", ThemeResources.StatusSuccess.Color);
                    }
                    catch (Exception ex)
                    {
                        SetStatus($"삽입 실패: {ex.Message}", ThemeResources.StatusError.Color);
                        Logger.Log($"[RemoteInsert] 삽입 실패: {ex}");
                    }
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                    SetStatus($"다운로드 실패: {ex.Message}", ThemeResources.StatusError.Color));
                Logger.Log($"[RemoteInsert] 다운로드 실패: {ex}");
            }
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
            if (_styleStack == null) return;
            var effective = BuildEffectiveStyleConfig();
            _styleStack.Children.Clear();

            var palettes = effective.Palettes ?? new List<StylePalette>();
            var fonts    = effective.Fonts    ?? new List<StyleFont>();

            if (palettes.Count == 0 && fonts.Count == 0) return;

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
                    Logger.Log($"[Style] Palette clicked: idx={idx}, palette={palettes[idx]?.Name ?? "NULL"}, fontNull={_selectedFont == null}");
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
                    Logger.Log($"[Style] Font row clicked: idx={idx}, font={fonts[idx]?.Name ?? "NULL"}, paletteNull={_selectedPalette == null}");
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

        public void SetStyleAnchorByFile(string pathOrFile)
        {
            if (string.IsNullOrEmpty(pathOrFile)) return;
            var key = System.IO.Path.GetFileName(pathOrFile);
            Logger.Log($"[Style] SetStyleAnchorByFile key={key}, allAssets={_allAssets?.Count ?? -1}");
            _anchorAsset = (_allAssets ?? new List<HeaderAsset>())
                .FirstOrDefault(a =>
                    string.Equals(System.IO.Path.GetFileName(a.File ?? ""), key,
                        StringComparison.OrdinalIgnoreCase)
                    || (a.Extra != null && a.Extra.TryGetValue("remote_file", out var rf)
                        && string.Equals(System.IO.Path.GetFileName(rf.ToString()), key,
                            StringComparison.OrdinalIgnoreCase)));
            Logger.Log($"[Style] anchor={_anchorAsset?.Name ?? "NULL"}, colors={_anchorAsset?.Colors?.Count ?? -1}, styleStack={(_styleStack != null ? "OK" : "NULL")}");
            PopulateStylePanel();
        }

        private StyleConfig BuildEffectiveStyleConfig()
        {
            var palettes = new List<StylePalette>();
            var fonts = new List<StyleFont>();

            if (_anchorAsset != null)
            {
                var np = PaletteRoleMapper.Map(_anchorAsset.Colors);
                palettes.AddRange(PaletteGenerator.Generate(np));

                if (_anchorAsset.Fonts != null)
                {
                    foreach (var f in _anchorAsset.Fonts)
                    {
                        if (f == null || string.IsNullOrWhiteSpace(f.Family)) continue;
                        if (fonts.Any(x => string.Equals(x.Name, f.Family,
                                StringComparison.OrdinalIgnoreCase))) continue;
                        fonts.Add(new StyleFont { Name = f.Family, Fallback = f.Fallback,
                            Mood = new List<string>(), UseWhen = "에셋 폰트" });
                    }
                }
            }

            if (_styleConfig?.Palettes != null) palettes.AddRange(_styleConfig.Palettes);
            if (_styleConfig?.Fonts != null)
            {
                foreach (var f in _styleConfig.Fonts)
                {
                    if (f == null || string.IsNullOrWhiteSpace(f.Name)) continue;
                    if (fonts.Any(x => string.Equals(x.Name, f.Name,
                            StringComparison.OrdinalIgnoreCase))) continue;
                    fonts.Add(f);
                }
            }

            return new StyleConfig { Palettes = palettes, Fonts = fonts };
        }

        public void InitAi(IAiService aiService, StyleConfig styles, RemoteAssetCache remoteCache = null)
        {
            _aiService = aiService;
            _styleConfig = styles;
            _remoteCache = remoteCache;
            PopulateStylePanel();
        }

        public void ShowIngestButton()
        {
            if (_ingestButton != null)
                _ingestButton.Visibility = Visibility.Visible;
        }

        private static readonly SolidColorBrush IngestAccent = Freeze(new SolidColorBrush(Color.FromRgb(249, 115, 22)));
        private static readonly SolidColorBrush IngestAccentDim = Freeze(new SolidColorBrush(Color.FromArgb(30, 249, 115, 22)));
        private static readonly SolidColorBrush IngestTrack = Freeze(new SolidColorBrush(Color.FromRgb(63, 63, 70)));
        private static SolidColorBrush Freeze(SolidColorBrush b) { b.Freeze(); return b; }

        private async Task RunIngestAsync()
        {
            var dlg = new System.Windows.Forms.OpenFileDialog
            {
                Title = "인제스트할 에셋 번들 선택",
                Filter = "PowerPoint 파일 (*.pptx)|*.pptx",
                Multiselect = false
            };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            var bundlePath = dlg.FileName;
            var bundleName = System.IO.Path.GetFileName(bundlePath);

            SwitchTab(0);
            if (_emptyState != null) _emptyState.Visibility = Visibility.Collapsed;

            _ingestButton.IsEnabled = false;
            ((TextBlock)_ingestButton.Child).Text = "실행 중...";
            _ingestLastIndex = 0;

            BuildIngestPanel(bundleName);

            try
            {
                var outputDir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TeampptAddin", "ingest-output");

                if (Directory.Exists(outputDir))
                {
                    foreach (var f in Directory.GetFiles(outputDir))
                        System.IO.File.Delete(f);
                }

                SetIngestStage("슬라이드 분해 중");
                var items = IngestRunner.Run(bundlePath, outputDir);
                _ingestProgressText.Text = $"0/{items.Count} · 분해 완료";

                _retryOutputDir = outputDir;
                _retryBundleName = bundleName;
                _retryStartFrom = 0;

                var count = await RunUploadAsync(outputDir, bundleName, 0);

                StopIngestAnimations();
                ShowIngestComplete(count, items.Count);
                SetStatus($"인제스트 완료: {count}개", ThemeResources.StatusSuccess.Color);
                Logger.Log($"[IngestButton] 완료: split={items.Count}, uploaded={count}");
            }
            catch (Exception ex)
            {
                StopIngestAnimations();
                _retryStartFrom = Math.Max(0, _ingestLastIndex - 1);
                ShowIngestError(ex.Message);
                SetStatus("인제스트 실패", ThemeResources.StatusError.Color);
                Logger.Log($"[IngestButton] 실패: {ex}");
            }
            finally
            {
                _ingestButton.IsEnabled = true;
                ((TextBlock)_ingestButton.Child).Text = "인제스트";
            }
        }

        private async Task<int> RunUploadAsync(string outputDir, string bundleName, int startFrom)
        {
            var uploader = new AssetIngestUploader();
            var dispatcher = Dispatcher;
            return await uploader.UploadDirectoryAsync(outputDir, bundleName, p =>
            {
                dispatcher.Invoke(() => HandleIngestProgress(p));
            }, startFrom);
        }

        private async Task ResumeIngestAsync()
        {
            _ingestButton.IsEnabled = false;
            ((TextBlock)_ingestButton.Child).Text = "실행 중...";

            _ingestCurrentContainer.Children.Clear();
            SetIngestStage("재시도 중");
            _ingestProgressText.Foreground = ThemeResources.TextSub;

            try
            {
                var count = await RunUploadAsync(_retryOutputDir, _retryBundleName, _retryStartFrom);

                StopIngestAnimations();
                var totalFiles = Directory.GetFiles(_retryOutputDir, "*.png").Length;
                ShowIngestComplete(count + _retryStartFrom, totalFiles);
                SetStatus($"인제스트 완료: {count + _retryStartFrom}개", ThemeResources.StatusSuccess.Color);
            }
            catch (Exception ex)
            {
                StopIngestAnimations();
                _retryStartFrom = Math.Max(0, _ingestLastIndex - 1);
                ShowIngestError(ex.Message);
                SetStatus("인제스트 실패", ThemeResources.StatusError.Color);
                Logger.Log($"[IngestRetry] 실패: {ex}");
            }
            finally
            {
                _ingestButton.IsEnabled = true;
                ((TextBlock)_ingestButton.Child).Text = "인제스트";
            }
        }

        // ── Ingest Panel ─────────────────────────────────────────────

        private void BuildIngestPanel(string bundleName)
        {
            var panel = new StackPanel { Margin = new Thickness(8, 8, 8, 4) };

            var card = new Border
            {
                Background = ThemeResources.BgAiResponse,
                CornerRadius = new CornerRadius(16),
                Padding = new Thickness(16, 14, 16, 14)
            };

            var inner = new StackPanel();

            inner.Children.Add(new TextBlock
            {
                Text = "INGEST",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = IngestAccent,
                FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(0, 0, 0, 2)
            });

            inner.Children.Add(new TextBlock
            {
                Text = bundleName,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain,
                FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(0, 0, 0, 10)
            });

            var progressGrid = new Grid { Height = 4, Margin = new Thickness(0, 0, 0, 6) };
            _ingestProgressTrack = new Border
            {
                CornerRadius = new CornerRadius(2),
                Background = IngestTrack
            };
            _ingestProgressFill = new Border
            {
                CornerRadius = new CornerRadius(2),
                Background = IngestAccent,
                HorizontalAlignment = HorizontalAlignment.Left,
                Width = 0
            };
            progressGrid.Children.Add(_ingestProgressTrack);
            progressGrid.Children.Add(_ingestProgressFill);
            inner.Children.Add(progressGrid);

            _ingestProgressText = new TextBlock
            {
                Text = "준비 중...",
                FontSize = 10,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(0, 0, 0, 12)
            };
            inner.Children.Add(_ingestProgressText);

            _ingestCurrentContainer = new Grid { MinHeight = 10 };
            inner.Children.Add(_ingestCurrentContainer);

            card.Child = inner;
            panel.Children.Add(card);

            _chatStack.Children.Add(panel);

            _completedCount = 0;
            _completedExpanded = false;
            _ingestCompletedStack = new StackPanel { Visibility = Visibility.Collapsed, Margin = new Thickness(0, 8, 0, 0) };

            var latestThumbBorder = new Border
            {
                Width = 36,
                Height = 26,
                CornerRadius = new CornerRadius(6),
                ClipToBounds = true,
                Background = ThemeResources.BgChip,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            };
            _completedLatestThumb = new Image { Stretch = Stretch.UniformToFill };
            latestThumbBorder.Child = _completedLatestThumb;
            ThemeResources.ApplyRoundedClip(latestThumbBorder, 6);

            _completedLatestName = new TextBlock
            {
                Text = "",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            _completedCountText = new TextBlock
            {
                Text = "",
                FontSize = 10,
                Foreground = ThemeResources.TextDim,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0, 0, 0)
            };

            _completedChevron = new TextBlock
            {
                Text = "▾",
                FontSize = 12,
                Foreground = ThemeResources.TextDim,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center
            };

            var previewRow = new Grid();
            previewRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            previewRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            previewRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            previewRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(latestThumbBorder, 0);
            Grid.SetColumn(_completedLatestName, 1);
            Grid.SetColumn(_completedCountText, 2);
            Grid.SetColumn(_completedChevron, 3);
            previewRow.Children.Add(latestThumbBorder);
            previewRow.Children.Add(_completedLatestName);
            previewRow.Children.Add(_completedCountText);
            previewRow.Children.Add(_completedChevron);

            _completedCollapsedBox = new Border
            {
                Background = ThemeResources.BgAiResponse,
                CornerRadius = new CornerRadius(14),
                Padding = new Thickness(12, 10, 14, 10),
                Margin = new Thickness(8, 6, 8, 4),
                Cursor = Cursors.Hand,
                Visibility = Visibility.Collapsed
            };

            var collapsedContent = new StackPanel();
            collapsedContent.Children.Add(previewRow);
            collapsedContent.Children.Add(_ingestCompletedStack);
            _completedCollapsedBox.Child = collapsedContent;

            _completedCollapsedBox.MouseLeftButtonUp += (s, e) =>
            {
                _completedExpanded = !_completedExpanded;
                _completedChevron.Text = _completedExpanded ? "▴" : "▾";
                _ingestCompletedStack.Visibility = _completedExpanded ? Visibility.Visible : Visibility.Collapsed;
            };

            _chatStack.Children.Add(_completedCollapsedBox);
            _chatScroll.ScrollToBottom();
        }

        private void HandleIngestProgress(IngestProgress p)
        {
            if (p.Stage == IngestStage.Understanding && p.Index != _ingestLastIndex)
            {
                MoveCurrentToCompleted();
                ShowCurrentCard(p);
                _ingestLastIndex = p.Index;
            }

            UpdateProgressBar(p.Index, p.Total, p.Stage);

            switch (p.Stage)
            {
                case IngestStage.Understanding:
                    SetIngestStage("이해 중");
                    break;
                case IngestStage.Embedding:
                    if (!string.IsNullOrEmpty(p.Name) && _ingestNameText != null)
                        _ingestNameText.Text = p.Name;
                    SetIngestStage("임베딩 중");
                    break;
                case IngestStage.Uploading:
                    SetIngestStage("업로드 중");
                    break;
                case IngestStage.AssetDone:
                    StopIngestAnimations();
                    ShowAssetDone(p.Kind);
                    break;
            }

            _chatScroll.ScrollToBottom();
        }

        private BitmapImage _currentThumbBmp;

        private void ShowCurrentCard(IngestProgress p)
        {
            _ingestCurrentContainer.Children.Clear();
            _currentThumbBmp = null;

            var cardStack = new StackPanel();

            var thumbClip = new Border
            {
                CornerRadius = new CornerRadius(10),
                ClipToBounds = true,
                Background = ThemeResources.BgChip,
                Margin = new Thickness(0, 0, 0, 10)
            };

            var thumbGrid = new Grid();
            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new Uri(p.PngPath);
                bmp.DecodePixelWidth = 400;
                bmp.EndInit();
                bmp.Freeze();
                _currentThumbBmp = bmp;
                thumbGrid.Children.Add(new Image
                {
                    Source = bmp,
                    Stretch = Stretch.Uniform
                });
            }
            catch { }

            _ingestScanLine = new Border
            {
                Height = 32,
                VerticalAlignment = VerticalAlignment.Top,
                IsHitTestVisible = false,
                Background = new LinearGradientBrush(
                    Color.FromArgb(0, 249, 115, 22),
                    Color.FromArgb(60, 249, 115, 22),
                    new Point(0, 0), new Point(0, 1))
            };
            thumbGrid.Children.Add(_ingestScanLine);

            try
            {
                var scanTransform = new TranslateTransform(0, -32);
                _ingestScanLine.RenderTransform = scanTransform;
                var scanAnim = new DoubleAnimation
                {
                    From = -32,
                    To = 160,
                    Duration = TimeSpan.FromSeconds(1.4),
                    RepeatBehavior = RepeatBehavior.Forever,
                    EasingFunction = new SineEase()
                };
                scanTransform.BeginAnimation(TranslateTransform.YProperty, scanAnim);
            }
            catch { }

            thumbClip.Child = thumbGrid;
            ThemeResources.ApplyRoundedClip(thumbClip, 10);
            cardStack.Children.Add(thumbClip);

            _ingestNameText = new TextBlock
            {
                Text = p.AssetId.Replace("_", " "),
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.TextMain,
                FontFamily = ThemeResources.FontBase,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(2, 0, 0, 2)
            };
            cardStack.Children.Add(_ingestNameText);

            _ingestStageText = new TextBlock
            {
                FontSize = 11,
                Foreground = IngestAccent,
                FontFamily = ThemeResources.FontBase,
                Margin = new Thickness(2, 0, 0, 0)
            };
            cardStack.Children.Add(_ingestStageText);

            var wrapper = new Border
            {
                Child = cardStack,
                Background = IngestAccentDim,
                CornerRadius = new CornerRadius(14),
                Padding = new Thickness(12, 12, 12, 12),
                Margin = new Thickness(0, 0, 0, 8),
                Opacity = 0,
                RenderTransform = new TranslateTransform(0, 10)
            };

            _ingestCurrentContainer.Children.Add(wrapper);

            try
            {
                var fadeIn = new DoubleAnimation { To = 1, Duration = TimeSpan.FromMilliseconds(300), EasingFunction = new QuadraticEase() };
                var slideUp = new DoubleAnimation { To = 0, Duration = TimeSpan.FromMilliseconds(300), EasingFunction = new QuadraticEase() };
                wrapper.BeginAnimation(OpacityProperty, fadeIn);
                ((TranslateTransform)wrapper.RenderTransform).BeginAnimation(TranslateTransform.YProperty, slideUp);
            }
            catch { wrapper.Opacity = 1; }
        }

        private void SetIngestStage(string stage)
        {
            _ingestStageBase = stage;
            if (_ingestStageText != null)
                _ingestStageText.Text = stage + "···";

            if (_ingestDotsTimer != null) _ingestDotsTimer.Stop();
            int dotCount = 1;
            _ingestDotsTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(350) };
            _ingestDotsTimer.Tick += (s, e) =>
            {
                dotCount = dotCount % 3 + 1;
                if (_ingestStageText != null)
                    _ingestStageText.Text = _ingestStageBase + new string('·', dotCount);
            };
            _ingestDotsTimer.Start();
        }

        private void ShowAssetDone(string kind)
        {
            if (_ingestStageText == null) return;
            _ingestStageText.Text = $"완료 · {kind}";
            _ingestStageText.Foreground = ThemeResources.StatusSuccess;
        }

        private void MoveCurrentToCompleted()
        {
            if (_ingestCurrentContainer == null || _ingestCurrentContainer.Children.Count == 0) return;

            var current = _ingestCurrentContainer.Children[0] as Border;
            if (current == null) return;

            var cardStack = current.Child as StackPanel;
            if (cardStack == null) return;

            string name = "";
            string kind = "";
            BitmapImage thumbBmp = _currentThumbBmp;
            foreach (var child in cardStack.Children)
            {
                var tb = child as TextBlock;
                if (tb == null) continue;
                if (tb.FontWeight == FontWeights.SemiBold) name = tb.Text;
                else if (tb.Foreground == IngestAccent || tb.Foreground == ThemeResources.StatusSuccess) kind = tb.Text;
            }

            _ingestCurrentContainer.Children.Clear();
            _currentThumbBmp = null;

            var row = new Grid { Margin = new Thickness(0, 0, 0, 4) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var miniThumb = new Border
            {
                Width = 34,
                Height = 24,
                CornerRadius = new CornerRadius(5),
                ClipToBounds = true,
                Background = ThemeResources.BgChip,
                VerticalAlignment = VerticalAlignment.Center
            };
            if (thumbBmp != null)
            {
                miniThumb.Child = new Image
                {
                    Source = thumbBmp,
                    Stretch = Stretch.UniformToFill
                };
            }
            ThemeResources.ApplyRoundedClip(miniThumb, 5);
            Grid.SetColumn(miniThumb, 0);
            row.Children.Add(miniThumb);

            var textStack = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };

            textStack.Children.Add(new TextBlock
            {
                Text = "✓",
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeResources.StatusSuccess,
                Margin = new Thickness(0, 0, 5, 0),
                VerticalAlignment = VerticalAlignment.Center
            });

            textStack.Children.Add(new TextBlock
            {
                Text = name,
                FontSize = 11,
                Foreground = ThemeResources.TextSub,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            });

            if (!string.IsNullOrEmpty(kind))
            {
                textStack.Children.Add(new TextBlock
                {
                    Text = $" · {kind.Replace("완료 · ", "")}",
                    FontSize = 10,
                    Foreground = ThemeResources.TextDim,
                    FontFamily = ThemeResources.FontBase,
                    VerticalAlignment = VerticalAlignment.Center
                });
            }

            Grid.SetColumn(textStack, 1);
            row.Children.Add(textStack);

            row.Opacity = 0;
            _ingestCompletedStack.Children.Insert(0, row);

            try
            {
                var fadeIn = new DoubleAnimation { To = 1, Duration = TimeSpan.FromMilliseconds(200) };
                row.BeginAnimation(OpacityProperty, fadeIn);
            }
            catch { row.Opacity = 1; }

            _completedCount++;
            _completedCountText.Text = $"· {_completedCount}개";
            _completedLatestName.Text = "✓ " + name;
            if (thumbBmp != null)
                _completedLatestThumb.Source = thumbBmp;
            if (_completedCollapsedBox.Visibility == Visibility.Collapsed)
                _completedCollapsedBox.Visibility = Visibility.Visible;
        }

        private void UpdateProgressBar(int index, int total, IngestStage stage)
        {
            var trackWidth = _ingestProgressTrack.ActualWidth;
            if (trackWidth <= 0) trackWidth = 200;

            double stageOffset = 0;
            switch (stage)
            {
                case IngestStage.Understanding: stageOffset = 0; break;
                case IngestStage.Embedding: stageOffset = 0.33; break;
                case IngestStage.Uploading: stageOffset = 0.66; break;
                case IngestStage.AssetDone: stageOffset = 1.0; break;
            }

            var pct = ((index - 1) + stageOffset) / total;
            var targetWidth = pct * trackWidth;

            try
            {
                var anim = new DoubleAnimation
                {
                    To = targetWidth,
                    Duration = TimeSpan.FromMilliseconds(300),
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                };
                _ingestProgressFill.BeginAnimation(WidthProperty, anim);
            }
            catch { _ingestProgressFill.Width = targetWidth; }

            var pctInt = (int)(pct * 100);
            _ingestProgressText.Text = $"{index}/{total} · {pctInt}%";
        }

        private void ShowIngestComplete(int uploaded, int total)
        {
            MoveCurrentToCompleted();

            _ingestProgressFill.BeginAnimation(WidthProperty, null);
            _ingestProgressFill.Width = _ingestProgressTrack.ActualWidth > 0 ? _ingestProgressTrack.ActualWidth : 200;
            _ingestProgressText.Text = $"{uploaded}/{total} · 완료";
            _ingestProgressText.Foreground = ThemeResources.StatusSuccess;

            _ingestCurrentContainer.Children.Clear();
            var done = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(20, 16, 185, 129)),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(14, 10, 14, 10),
                Margin = new Thickness(0, 4, 0, 0)
            };
            var doneRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
            doneRow.Children.Add(new TextBlock
            {
                Text = "✓",
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeResources.StatusSuccess,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            });
            doneRow.Children.Add(new TextBlock
            {
                Text = $"{uploaded}개 에셋 Supabase 업로드 완료",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeResources.StatusSuccess,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center
            });
            done.Child = doneRow;

            done.Opacity = 0;
            done.RenderTransform = new ScaleTransform(0.9, 0.9, 0.5, 0.5);
            done.RenderTransformOrigin = new Point(0.5, 0.5);
            _ingestCurrentContainer.Children.Add(done);

            try
            {
                var fadeIn = new DoubleAnimation { To = 1, Duration = TimeSpan.FromMilliseconds(400), EasingFunction = new QuadraticEase() };
                var scaleX = new DoubleAnimation { To = 1, Duration = TimeSpan.FromMilliseconds(400), EasingFunction = new ElasticEase { Oscillations = 1, Springiness = 8 } };
                var scaleY = new DoubleAnimation { To = 1, Duration = TimeSpan.FromMilliseconds(400), EasingFunction = new ElasticEase { Oscillations = 1, Springiness = 8 } };
                done.BeginAnimation(OpacityProperty, fadeIn);
                ((ScaleTransform)done.RenderTransform).BeginAnimation(ScaleTransform.ScaleXProperty, scaleX);
                ((ScaleTransform)done.RenderTransform).BeginAnimation(ScaleTransform.ScaleYProperty, scaleY);
            }
            catch { done.Opacity = 1; }
        }

        private void ShowIngestError(string message)
        {
            MoveCurrentToCompleted();
            _ingestProgressText.Text = $"오류 · {_retryStartFrom}개 완료";
            _ingestProgressText.Foreground = ThemeResources.StatusError;

            _ingestCurrentContainer.Children.Clear();

            var errPanel = new StackPanel();
            var errCard = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(20, 248, 113, 113)),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(14, 10, 14, 10)
            };

            bool isServerBusy = message.Contains("503") || message.Contains("429")
                || message.Contains("high demand") || message.Contains("UNAVAILABLE");

            var errStack = new StackPanel();
            if (isServerBusy)
            {
                errStack.Children.Add(new TextBlock
                {
                    Text = "서버가 바빠요",
                    FontSize = 13,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = ThemeResources.StatusError,
                    FontFamily = ThemeResources.FontBase,
                    Margin = new Thickness(0, 0, 0, 4)
                });
            }
            errStack.Children.Add(new TextBlock
            {
                Text = isServerBusy ? "일시적 과부하입니다. 잠시 후 다시 시도해주세요." : message,
                FontSize = 11,
                Foreground = ThemeResources.StatusError,
                FontFamily = ThemeResources.FontBase,
                TextWrapping = TextWrapping.Wrap
            });
            errCard.Child = errStack;
            errPanel.Children.Add(errCard);

            var retryBtn = new Border
            {
                Background = IngestAccent,
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(16, 9, 16, 9),
                Margin = new Thickness(0, 10, 0, 0),
                Cursor = Cursors.Hand,
                HorizontalAlignment = HorizontalAlignment.Left
            };

            var retryRow = new StackPanel { Orientation = Orientation.Horizontal };
            retryRow.Children.Add(new TextBlock
            {
                Text = "↻",
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            });
            retryRow.Children.Add(new TextBlock
            {
                Text = $"다시 시도 ({_retryStartFrom}번부터)",
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                FontFamily = ThemeResources.FontBase,
                VerticalAlignment = VerticalAlignment.Center
            });
            retryBtn.Child = retryRow;

            retryBtn.MouseEnter += (s, e) => retryBtn.Opacity = 0.85;
            retryBtn.MouseLeave += (s, e) => retryBtn.Opacity = 1;
            retryBtn.MouseLeftButtonUp += async (s, e) => await ResumeIngestAsync();

            errPanel.Children.Add(retryBtn);
            _ingestCurrentContainer.Children.Add(errPanel);
        }

        private void StopIngestAnimations()
        {
            if (_ingestDotsTimer != null) { _ingestDotsTimer.Stop(); _ingestDotsTimer = null; }
            try { _ingestScanLine?.RenderTransform?.BeginAnimation(TranslateTransform.YProperty, null); } catch { }
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

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using WinForms = System.Windows.Forms;

namespace WindowBackRecorder
{
    public sealed class AppEntry : Application
    {
        [STAThread]
        public static int Main()
        {
            var app = new AppEntry();
            app.ShutdownMode = ShutdownMode.OnMainWindowClose;
            app.Run(new MainWindow());
            return 0;
        }
    }

    public sealed class MainWindow : Window
    {
        private readonly string appDir;
        private readonly string supportDir;
        private readonly string settingsPath;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly DispatcherTimer processTimer = new DispatcherTimer();

        private ListView windowList;
        private TextBlock audioStatusText;
        private TextBox saveFolderBox;
        private TextBox logBox;
        private TextBlock statusText;
        private TextBlock targetText;
        private TextBlock engineText;
        private TextBlock outputText;
        private Button startButton;
        private Button stopButton;
        private ToggleButton listenToggle;
        private ToggleButton windowVisibilityToggle;
        private ToggleButton silentPlaybackToggle;
        private Slider fpsSlider;
        private Slider audioSyncSlider;
        private TextBlock audioSyncValueText;

        private Process ffmpegProcess;
        private Process audioProcess;
        private RecordingState activeRecording;
        private WindowInfo selectedWindow;
        private IntRect? savedBounds;
        private bool gfxCaptureAvailable;
        private bool isStopping;
        private bool isPausing;
        private bool isPaused;
        private bool pendingResumeAfterPause;
        private bool targetWindowHidden;
        private bool silentPlaybackApplied;
        private bool processAudioSupportChecked;
        private bool processAudioSupported;
        private readonly List<AudioSessionSnapshot> audioSessionSnapshots = new List<AudioSessionSnapshot>();
        private int busyTick;
        private int logLineCount;
        private DateTime lastFfmpegProgressLogUtc = DateTime.MinValue;
        private DateTime lastWindowScanUtc = DateTime.MinValue;
        private DateTime lastSilentPlaybackAttemptUtc = DateTime.MinValue;
        private string lastWindowSnapshot = "";

        private const string SupportFolderName = "프로그램 구성 파일";
        private const string NoAudioLabel = "소리 없이 녹화";
        private const string DeveloperLabel = "developed by yeohj0710";
        private const string DefaultRecordingsFolderName = "녹화 완료된 동영상";
        private const string OldDefaultRecordingsFolderName = "녹화 완료 영상";
        private const string TempFolderName = "recording-temp";
        private const double ReducedPlaybackVolume = 0.04;
        private const double ReducedPlaybackGain = 25.0;
        private const double AudioBoostFadeSeconds = 0.45;
        private const double DefaultAudioSyncDelaySeconds = 0.0;
        private const double MinAudioSyncDelaySeconds = -1.0;
        private const double MaxAudioSyncDelaySeconds = 1.0;
        private const double StartWarmupTrimSeconds = 3.0;
        private const int AudioReadyWaitMilliseconds = 2500;
        private const int VideoReadyWaitMilliseconds = 3500;

        public MainWindow()
        {
            appDir = AppDomain.CurrentDomain.BaseDirectory;
            string packagedSupportDir = Path.Combine(appDir, SupportFolderName);
            supportDir = Directory.Exists(packagedSupportDir) ? packagedSupportDir : appDir;
            settingsPath = Path.Combine(supportDir, "settings.json");

            Title = "백그라운드 영상 녹화 프로그램";
            string iconPath = Path.Combine(supportDir, "app.ico");
            if (File.Exists(iconPath))
            {
                Icon = BitmapFrame.Create(new Uri(iconPath, UriKind.Absolute));
            }
            var workArea = WinForms.Screen.PrimaryScreen.WorkingArea;
            Width = Math.Min(1160, Math.Max(1040, workArea.Width - 120));
            Height = Math.Min(820, Math.Max(720, workArea.Height - 80));
            MinWidth = Math.Min(1040, workArea.Width - 40);
            MinHeight = Math.Min(720, workArea.Height - 40);
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = workArea.Left + 40;
            Top = workArea.Top + 40;
            Background = Brush("#edf1f6");
            FontFamily = new FontFamily("Segoe UI");
            Foreground = Brush("#111827");

            BuildUi();
            LoadSettings();
            RefreshCapabilities();
            RefreshWindows();
            RefreshAudioSources();

            processTimer.Interval = TimeSpan.FromMilliseconds(800);
            processTimer.Tick += OnProcessTimer;
            processTimer.Start();

            Closing += OnClosing;
            AppDomain.CurrentDomain.ProcessExit += delegate { RestorePlaybackVolume(); };
        }

        private void BuildUi()
        {
            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            var header = new Border
            {
                Background = Brush("#f8fafc"),
                BorderBrush = Brush("#d9e2ec"),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(32, 22, 32, 22)
            };
            Grid.SetRow(header, 0);
            root.Children.Add(header);

            var headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            header.Child = headerGrid;

            var titleStack = new StackPanel { Orientation = Orientation.Vertical };
            headerGrid.Children.Add(titleStack);

            var titleRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center
            };
            titleStack.Children.Add(titleRow);

            var title = new TextBlock
            {
                Text = "백그라운드 영상 녹화 프로그램",
                FontSize = 30,
                FontWeight = FontWeights.Bold,
                Foreground = Brush("#111827"),
                VerticalAlignment = VerticalAlignment.Center
            };
            titleRow.Children.Add(title);
            titleRow.Children.Add(CreateDeveloperBadge());

            statusText = new TextBlock
            {
                Text = "준비됨",
                Margin = new Thickness(0, 8, 0, 0),
                FontSize = 14,
                Foreground = Brush("#475569")
            };
            titleStack.Children.Add(statusText);

            var headerButtons = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(headerButtons, 1);
            headerGrid.Children.Add(headerButtons);

            headerButtons.Children.Add(CreateButton("사용법", OpenUserGuide, "#eef4ff", "#dbeafe"));
            headerButtons.Children.Add(Spacer(8, 1));
            headerButtons.Children.Add(CreateButton("소리 확인", OpenSoundMixer, "#eef4ff", "#dbeafe"));
            headerButtons.Children.Add(Spacer(8, 1));
            headerButtons.Children.Add(CreateButton("새로고침", delegate { RefreshWindows(); RefreshAudioSources(); }, "#eef4ff", "#dbeafe"));

            var main = new Grid { Margin = new Thickness(24) };
            main.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            main.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(440) });
            Grid.SetRow(main, 1);
            root.Children.Add(main);

            var leftPanel = CreatePanel();
            Grid.SetColumn(leftPanel, 0);
            main.Children.Add(leftPanel);

            var leftDock = new DockPanel();
            leftPanel.Child = leftDock;

            var listHeader = new Grid { Margin = new Thickness(0, 0, 0, 12) };
            listHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            listHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            listHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            DockPanel.SetDock(listHeader, Dock.Top);
            leftDock.Children.Add(listHeader);

            listHeader.Children.Add(new TextBlock
            {
                Text = "녹화할 창",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brush("#111827"),
                VerticalAlignment = VerticalAlignment.Center
            });

            targetText = new TextBlock
            {
                Text = "창을 선택해주세요",
                FontSize = 12,
                Foreground = Brush("#64748b"),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(targetText, 1);
            listHeader.Children.Add(targetText);

            var listRefreshButton = CreateSmallButton("목록 새로고침", delegate { RefreshWindows(); });
            listRefreshButton.Height = 30;
            listRefreshButton.Margin = new Thickness(10, 0, 0, 0);
            Grid.SetColumn(listRefreshButton, 2);
            listHeader.Children.Add(listRefreshButton);

            windowList = new ListView
            {
                Background = Brush("#ffffff"),
                Foreground = Brush("#111827"),
                BorderBrush = Brush("#d9e2ec"),
                BorderThickness = new Thickness(1),
                SelectionMode = SelectionMode.Single
            };
            windowList.SelectionChanged += OnWindowSelected;
            windowList.MouseDoubleClick += delegate { BringTargetFront(); };
            StyleListView(windowList);
            leftDock.Children.Add(windowList);

            var gridView = new GridView();
            gridView.Columns.Add(CreateColumn("창 제목", "Title", 300));
            gridView.Columns.Add(CreateColumn("프로그램", "ProcessName", 100));
            gridView.Columns.Add(CreateColumn("크기", "SizeText", 82));
            windowList.View = gridView;

            var rightPanel = CreatePanel();
            rightPanel.Margin = new Thickness(12, 0, 0, 0);
            Grid.SetColumn(rightPanel, 1);
            main.Children.Add(rightPanel);

            var controlScroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Focusable = false
            };
            TryApplyDarkScrollBar(controlScroll);
            rightPanel.Child = controlScroll;

            var controls = new StackPanel();
            controlScroll.Content = controls;

            controls.Children.Add(SectionLabel("녹화 설정"));
            controls.Children.Add(FormLabel("저장 폴더"));

            var folderRow = new Grid { Margin = new Thickness(0, 4, 0, 14) };
            folderRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            folderRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(54) });
            folderRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(54) });
            controls.Children.Add(folderRow);

            saveFolderBox = CreateTextBox();
            Grid.SetColumn(saveFolderBox, 0);
            folderRow.Children.Add(saveFolderBox);

            var browse = CreateSmallButton("변경", BrowseFolder);
            browse.Margin = new Thickness(6, 0, 0, 0);
            Grid.SetColumn(browse, 1);
            folderRow.Children.Add(browse);

            var openFolder = CreateSmallButton("열기", OpenSaveFolder);
            openFolder.Margin = new Thickness(6, 0, 0, 0);
            Grid.SetColumn(openFolder, 2);
            folderRow.Children.Add(openFolder);

            var buttonGrid = new Grid { Margin = new Thickness(0, 4, 0, 18) };
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            controls.Children.Add(buttonGrid);

            startButton = CreateButton("녹화 시작", ToggleRecordingAction, "#1e6bff", "#3d82ff");
            startButton.Height = 46;
            Grid.SetColumn(startButton, 0);
            buttonGrid.Children.Add(startButton);

            stopButton = CreateButton("녹화 종료", StopRecordingFromUi, "#be123c", "#9f1239");
            stopButton.Height = 46;
            stopButton.IsEnabled = false;
            Grid.SetColumn(stopButton, 2);
            buttonGrid.Children.Add(stopButton);

            controls.Children.Add(SectionLabel("녹화창 숨기기"));
            var minimizeWarning = MetaText("창을 직접 최소화하면 녹화가 안 됩니다. 아래 버튼을 사용하세요.");
            minimizeWarning.Foreground = Brush("#b45309");
            minimizeWarning.TextWrapping = TextWrapping.Wrap;
            minimizeWarning.Margin = new Thickness(0, 0, 0, 8);
            controls.Children.Add(minimizeWarning);

            windowVisibilityToggle = CreateToggle("녹화창 숨기기 (다른 작업 가능)", false);
            windowVisibilityToggle.Height = 38;
            windowVisibilityToggle.Checked += delegate { HideTargetWindowForRecording(); };
            windowVisibilityToggle.Unchecked += delegate { RestoreTargetWindowFromHidden(); };
            controls.Children.Add(windowVisibilityToggle);

            controls.Children.Add(SectionLabel("소리"));
            audioStatusText = MetaText("선택한 앱 소리만 자동 녹음");
            audioStatusText.Margin = new Thickness(0, 4, 0, 6);
            audioStatusText.Foreground = Brush("#1d4ed8");
            controls.Children.Add(audioStatusText);

            var audioWarningText = MetaText("소리를 완전히 끄면 녹음도 안 될 수 있어요. 대신 아래 옵션으로 대상 앱 소리만 줄입니다.");
            audioWarningText.Margin = new Thickness(0, 0, 0, 14);
            audioWarningText.Foreground = Brush("#b45309");
            audioWarningText.TextWrapping = TextWrapping.Wrap;
            controls.Children.Add(audioWarningText);

            silentPlaybackToggle = CreateSwitchToggle("소리 줄이고 녹화하기", false);
            silentPlaybackToggle.Checked += delegate
            {
                ChangeReducedAudioMode(true);
            };
            silentPlaybackToggle.Unchecked += delegate
            {
                ChangeReducedAudioMode(false);
            };
            controls.Children.Add(silentPlaybackToggle);

            controls.Children.Add(SectionLabel("소리 싱크"));
            var syncHelpText = MetaText("소리가 영상보다 빠르면 +로, 소리가 늦으면 -로 조정하세요.");
            syncHelpText.TextWrapping = TextWrapping.Wrap;
            syncHelpText.Margin = new Thickness(0, 0, 0, 8);
            controls.Children.Add(syncHelpText);

            var audioSyncRow = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            audioSyncRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            audioSyncRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
            controls.Children.Add(audioSyncRow);

            audioSyncSlider = new Slider
            {
                Minimum = MinAudioSyncDelaySeconds * 1000.0,
                Maximum = MaxAudioSyncDelaySeconds * 1000.0,
                Value = DefaultAudioSyncDelaySeconds * 1000.0,
                TickFrequency = 50,
                SmallChange = 50,
                LargeChange = 100,
                IsSnapToTickEnabled = true
            };
            audioSyncSlider.ValueChanged += delegate
            {
                UpdateAudioSyncValueText();
                SaveCurrentSettings();
            };
            Grid.SetColumn(audioSyncSlider, 0);
            audioSyncRow.Children.Add(audioSyncSlider);

            audioSyncValueText = new TextBlock
            {
                Foreground = Brush("#334155"),
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Right
            };
            Grid.SetColumn(audioSyncValueText, 1);
            audioSyncRow.Children.Add(audioSyncValueText);
            UpdateAudioSyncValueText();

            controls.Children.Add(SectionLabel("화면 부드러움"));
            var fpsRow = new Grid { Margin = new Thickness(0, 4, 0, 16) };
            fpsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            fpsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(52) });
            controls.Children.Add(fpsRow);

            fpsSlider = new Slider { Minimum = 5, Maximum = 60, Value = 60, TickFrequency = 5, IsSnapToTickEnabled = true };
            Grid.SetColumn(fpsSlider, 0);
            fpsRow.Children.Add(fpsSlider);
            var fpsValue = new TextBlock { Foreground = Brush("#64748b"), VerticalAlignment = VerticalAlignment.Center, TextAlignment = TextAlignment.Right };
            fpsValue.SetBinding(TextBlock.TextProperty, new Binding("Value") { Source = fpsSlider, StringFormat = "{0:0} fps" });
            Grid.SetColumn(fpsValue, 1);
            fpsRow.Children.Add(fpsValue);

            listenToggle = CreateToggle("내 스피커로 듣기", false);
            listenToggle.Visibility = Visibility.Collapsed;

            controls.Children.Add(SectionLabel("상태"));
            engineText = MetaText("화면 캡처: 확인 중");
            outputText = MetaText("저장 파일: -");
            controls.Children.Add(engineText);
            controls.Children.Add(outputText);

            var footer = new Border
            {
                Background = Brush("#edf1f6"),
                BorderBrush = Brush("#d9e2ec"),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(20, 12, 20, 12)
            };
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            logBox = new TextBox
            {
                Height = 96,
                Background = Brush("#101827"),
                Foreground = Brush("#e2e8f0"),
                BorderThickness = new Thickness(0),
                IsReadOnly = true,
                TextWrapping = TextWrapping.NoWrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 12
            };
            footer.Child = logBox;
        }

        private Border CreatePanel()
        {
            return new Border
            {
                Background = Brush("#ffffff"),
                BorderBrush = Brush("#d9e2ec"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(18)
            };
        }

        private Border CreateDeveloperBadge()
        {
            return new Border
            {
                Margin = new Thickness(14, 4, 0, 0),
                Padding = new Thickness(10, 4, 10, 5),
                Background = Brush("#eaf2ff"),
                BorderThickness = new Thickness(0),
                CornerRadius = new CornerRadius(6),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = DeveloperLabel,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brush("#2563eb"),
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
        }

        private TextBlock SectionLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = Brush("#111827"),
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 18, 0, 12)
            };
        }

        private TextBlock FormLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = Brush("#334155"),
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 2)
            };
        }

        private TextBlock MetaText(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = Brush("#64748b"),
                FontSize = 13,
                Margin = new Thickness(0, 5, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
        }

        private Style CreateComboItemStyle()
        {
            var style = new Style(typeof(ComboBoxItem));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#17212b")));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#f8fafc")));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 6, 8, 6)));
            return style;
        }

        private TextBox CreateTextBox()
        {
            return new TextBox
            {
                Height = 34,
                Background = Brush("#ffffff"),
                Foreground = Brush("#111827"),
                BorderBrush = Brush("#cbd5e1"),
                Padding = new Thickness(10, 6, 10, 6)
            };
        }

        private void TryApplyDarkScrollBar(ScrollViewer scrollViewer)
        {
            try
            {
                scrollViewer.Resources[typeof(ScrollBar)] = XamlReader.Parse(@"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type ScrollBar}"">
  <Setter Property=""Width"" Value=""8""/>
  <Setter Property=""MinWidth"" Value=""8""/>
  <Setter Property=""Background"" Value=""#edf1f6""/>
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type ScrollBar}"">
        <Border Background=""#edf1f6"">
          <Track x:Name=""PART_Track"" IsDirectionReversed=""True"">
            <Track.DecreaseRepeatButton>
              <RepeatButton Command=""{x:Static ScrollBar.PageUpCommand}"" Opacity=""0"" Focusable=""False""/>
            </Track.DecreaseRepeatButton>
            <Track.Thumb>
              <Thumb MinHeight=""36"">
                <Thumb.Template>
                  <ControlTemplate TargetType=""{x:Type Thumb}"">
                    <Border Width=""6"" Margin=""1"" CornerRadius=""3"" Background=""#94a3b8""/>
                  </ControlTemplate>
                </Thumb.Template>
              </Thumb>
            </Track.Thumb>
            <Track.IncreaseRepeatButton>
              <RepeatButton Command=""{x:Static ScrollBar.PageDownCommand}"" Opacity=""0"" Focusable=""False""/>
            </Track.IncreaseRepeatButton>
          </Track>
        </Border>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>") as Style;
            }
            catch
            {
                // The app still works with the default scrollbar if this style fails.
            }
        }

        private ToggleButton CreateToggle(string text, bool isChecked)
        {
            var toggle = new ToggleButton
            {
                Content = text,
                IsChecked = isChecked,
                Margin = new Thickness(0, 0, 0, 8),
                Height = 34,
                Style = CreateToggleStyle(),
                Cursor = System.Windows.Input.Cursors.Hand,
                Focusable = false
            };
            return toggle;
        }

        private ToggleButton CreateSwitchToggle(string text, bool isChecked)
        {
            var toggle = new ToggleButton
            {
                Content = text,
                IsChecked = false,
                Height = 34,
                Margin = new Thickness(0, 0, 0, 8),
                Style = CreateSwitchStyle(),
                Cursor = System.Windows.Input.Cursors.Hand,
                Focusable = false
            };
            if (isChecked)
            {
                toggle.Loaded += delegate { toggle.IsChecked = true; };
            }
            return toggle;
        }

        private Button CreateSmallButton(string text, Action action)
        {
            var button = CreateButton(text, action, "#eef4ff", "#dbeafe");
            button.Height = 34;
            button.Margin = new Thickness(0, 0, 8, 8);
            return button;
        }

        private Button CreateButton(string text, Action action, string background, string hover)
        {
            var button = new Button
            {
                Content = text,
                Padding = new Thickness(10, 7, 10, 7),
                Margin = new Thickness(0),
                BorderThickness = new Thickness(1),
                Style = CreateButtonStyle(background, hover),
                Cursor = System.Windows.Input.Cursors.Hand,
                Focusable = false
            };
            button.Click += delegate { action(); };
            return button;
        }

        private Style CreateButtonStyle(string background, string hover)
        {
            bool lightButton = IsLightColor(background);
            Brush normalText = lightButton ? Brush("#1d4ed8") : Brush("#f5f7fa");
            Brush hoverText = lightButton ? Brush("#1e3a8a") : Brush("#ffffff");

            var style = new Style(typeof(Button));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brush(background)));
            style.Setters.Add(new Setter(Control.ForegroundProperty, normalText));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, Brush(hover)));
            style.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(FrameworkElement.FocusVisualStyleProperty, null));
            style.Setters.Add(new Setter(Control.TemplateProperty, CreateButtonTemplate()));

            var over = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            over.Setters.Add(new Setter(Control.BackgroundProperty, Brush(hover)));
            over.Setters.Add(new Setter(Control.ForegroundProperty, hoverText));
            style.Triggers.Add(over);

            var pressed = new Trigger { Property = ButtonBase.IsPressedProperty, Value = true };
            pressed.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#0f2438")));
            pressed.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#ffffff")));
            style.Triggers.Add(pressed);

            var disabled = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabled.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#18222c")));
            disabled.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#9fb1c4")));
            disabled.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2b3a49")));
            disabled.Setters.Add(new Setter(UIElement.OpacityProperty, 1.0));
            style.Triggers.Add(disabled);

            return style;
        }

        private bool IsLightColor(string hex)
        {
            try
            {
                Color color = (Color)ColorConverter.ConvertFromString(hex);
                double brightness = (color.R * 299 + color.G * 587 + color.B * 114) / 1000.0;
                return brightness > 185;
            }
            catch
            {
                return false;
            }
        }

        private Style CreateToggleStyle()
        {
            var style = new Style(typeof(ToggleButton));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#f8fafc")));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#111827")));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#cbd5e1")));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(FrameworkElement.FocusVisualStyleProperty, null));
            style.Setters.Add(new Setter(Control.TemplateProperty, CreateButtonTemplate()));

            var checkedTrigger = new Trigger { Property = ToggleButton.IsCheckedProperty, Value = true };
            checkedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#eaf2ff")));
            checkedTrigger.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#1d4ed8")));
            checkedTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2563eb")));
            style.Triggers.Add(checkedTrigger);

            var over = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            over.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#eef4ff")));
            over.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#1d4ed8")));
            style.Triggers.Add(over);

            var disabled = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabled.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#18222c")));
            disabled.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#9fb1c4")));
            disabled.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2b3a49")));
            disabled.Setters.Add(new Setter(UIElement.OpacityProperty, 1.0));
            style.Triggers.Add(disabled);

            return style;
        }

        private Style CreateSwitchStyle()
        {
            var style = new Style(typeof(ToggleButton));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#f8fafc")));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#111827")));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#cbd5e1")));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(FrameworkElement.FocusVisualStyleProperty, null));
            style.Setters.Add(new Setter(Control.TemplateProperty, CreateSwitchTemplate()));

            var checkedTrigger = new Trigger { Property = ToggleButton.IsCheckedProperty, Value = true };
            checkedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#eaf2ff")));
            checkedTrigger.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#1d4ed8")));
            checkedTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2563eb")));
            style.Triggers.Add(checkedTrigger);

            var over = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            over.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2563eb")));
            over.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#1d4ed8")));
            style.Triggers.Add(over);

            var disabled = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabled.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#18222c")));
            disabled.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#9fb1c4")));
            disabled.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2b3a49")));
            disabled.Setters.Add(new Setter(UIElement.OpacityProperty, 1.0));
            style.Triggers.Add(disabled);

            return style;
        }

        private ControlTemplate CreateButtonTemplate()
        {
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));
            border.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            border.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(8));
            border.SetValue(Border.SnapsToDevicePixelsProperty, true);

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            content.SetValue(ContentPresenter.MarginProperty, new TemplateBindingExtension(Control.PaddingProperty));
            content.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);
            border.AppendChild(content);

            return new ControlTemplate(typeof(ButtonBase)) { VisualTree = border };
        }

        private ControlTemplate CreateSwitchTemplate()
        {
            var outer = new FrameworkElementFactory(typeof(Border));
            outer.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));
            outer.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            outer.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));
            outer.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            outer.SetValue(Border.SnapsToDevicePixelsProperty, true);

            var dock = new FrameworkElementFactory(typeof(DockPanel));
            dock.SetValue(FrameworkElement.MarginProperty, new Thickness(10, 0, 10, 0));
            outer.AppendChild(dock);

            var track = new FrameworkElementFactory(typeof(Border));
            track.Name = "SwitchTrack";
            track.SetValue(DockPanel.DockProperty, Dock.Right);
            track.SetValue(FrameworkElement.WidthProperty, 42.0);
            track.SetValue(FrameworkElement.HeightProperty, 20.0);
            track.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Right);
            track.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            track.SetValue(FrameworkElement.MarginProperty, new Thickness(10, 0, 0, 0));
            track.SetValue(Border.CornerRadiusProperty, new CornerRadius(10));
            track.SetValue(Border.BackgroundProperty, Brush("#cbd5e1"));
            track.SetValue(Border.BorderBrushProperty, Brush("#cbd5e1"));
            track.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            dock.AppendChild(track);

            var thumb = new FrameworkElementFactory(typeof(System.Windows.Shapes.Ellipse));
            thumb.Name = "SwitchThumb";
            thumb.SetValue(FrameworkElement.WidthProperty, 14.0);
            thumb.SetValue(FrameworkElement.HeightProperty, 14.0);
            thumb.SetValue(FrameworkElement.MarginProperty, new Thickness(3));
            thumb.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Left);
            thumb.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            thumb.SetValue(UIElement.RenderTransformProperty, new TranslateTransform(0, 0));
            thumb.SetValue(System.Windows.Shapes.Shape.FillProperty, Brush("#ffffff"));
            track.AppendChild(thumb);

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            content.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);
            dock.AppendChild(content);

            var template = new ControlTemplate(typeof(ToggleButton)) { VisualTree = outer };

            var checkedTrigger = new Trigger { Property = ToggleButton.IsCheckedProperty, Value = true };
            checkedTrigger.Setters.Add(new Setter(Border.BackgroundProperty, Brush("#35d0c2"), "SwitchTrack"));
            checkedTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, Brush("#35d0c2"), "SwitchTrack"));
            checkedTrigger.Setters.Add(new Setter(System.Windows.Shapes.Ellipse.FillProperty, Brush("#ffffff"), "SwitchThumb"));
            checkedTrigger.EnterActions.Add(new BeginStoryboard { Storyboard = CreateSwitchThumbStoryboard(22) });
            checkedTrigger.ExitActions.Add(new BeginStoryboard { Storyboard = CreateSwitchThumbStoryboard(0) });
            template.Triggers.Add(checkedTrigger);

            var disabledTrigger = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabledTrigger.Setters.Add(new Setter(Border.BackgroundProperty, Brush("#2a3440"), "SwitchTrack"));
            disabledTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, Brush("#344252"), "SwitchTrack"));
            disabledTrigger.Setters.Add(new Setter(System.Windows.Shapes.Ellipse.FillProperty, Brush("#7f91a5"), "SwitchThumb"));
            template.Triggers.Add(disabledTrigger);

            return template;
        }

        private Storyboard CreateSwitchThumbStoryboard(double x)
        {
            var animation = new DoubleAnimation
            {
                To = x,
                Duration = new Duration(TimeSpan.FromMilliseconds(170)),
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };
            Storyboard.SetTargetName(animation, "SwitchThumb");
            Storyboard.SetTargetProperty(animation, new PropertyPath("(UIElement.RenderTransform).(TranslateTransform.X)"));

            var storyboard = new Storyboard();
            storyboard.Children.Add(animation);
            return storyboard;
        }

        private FrameworkElement Spacer(double width, double height)
        {
            return new Border { Width = width, Height = height };
        }

        private GridViewColumn CreateColumn(string header, string path, double width)
        {
            var textBlock = new FrameworkElementFactory(typeof(TextBlock));
            textBlock.SetBinding(TextBlock.TextProperty, new Binding(path));
            textBlock.SetValue(TextBlock.ForegroundProperty, Brush("#111827"));
            textBlock.SetValue(TextBlock.PaddingProperty, new Thickness(8, 0, 8, 0));
            textBlock.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);
            textBlock.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);

            return new GridViewColumn
            {
                Header = header,
                Width = width,
                CellTemplate = new DataTemplate { VisualTree = textBlock }
            };
        }

        private void StyleListView(ListView list)
        {
            var itemStyle = new Style(typeof(ListViewItem));
            itemStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(0, 7, 0, 7)));
            itemStyle.Setters.Add(new Setter(Control.MarginProperty, new Thickness(0, 0, 0, 1)));
            itemStyle.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#ffffff")));
            itemStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#111827")));
            itemStyle.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#ffffff")));
            itemStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            itemStyle.Setters.Add(new Setter(FrameworkElement.FocusVisualStyleProperty, null));

            var selectedTrigger = new Trigger { Property = ListViewItem.IsSelectedProperty, Value = true };
            selectedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#dbeafe")));
            selectedTrigger.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#111827")));
            selectedTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2563eb")));
            itemStyle.Triggers.Add(selectedTrigger);

            var mouseOverTrigger = new Trigger { Property = ListViewItem.IsMouseOverProperty, Value = true };
            mouseOverTrigger.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#f1f5f9")));
            mouseOverTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#bfdbfe")));
            itemStyle.Triggers.Add(mouseOverTrigger);

            list.ItemContainerStyle = itemStyle;
        }

        private void RefreshCapabilities()
        {
            gfxCaptureAvailable = TestFfmpegFilter("gfxcapture");
            engineText.Text = gfxCaptureAvailable ? "화면 캡처: gfxcapture" : "화면 캡처: 예비 방식";
        }

        private void RefreshWindows()
        {
            RefreshWindows(false);
        }

        private void RefreshWindows(bool quiet)
        {
            var selectedHandle = selectedWindow == null ? IntPtr.Zero : selectedWindow.Handle;
            var windows = NativeWindows.GetOpenWindows();
            string snapshot = BuildWindowSnapshot(windows);
            if (quiet && string.Equals(snapshot, lastWindowSnapshot, StringComparison.Ordinal))
            {
                return;
            }
            lastWindowSnapshot = snapshot;

            windowList.ItemsSource = windows;
            selectedWindow = null;
            targetText.Text = "창을 선택해주세요";

            foreach (WindowInfo info in windows)
            {
                if (info.Handle == selectedHandle)
                {
                    windowList.SelectedItem = info;
                    break;
                }
            }

            if (!quiet)
            {
                AppendLog("창 목록 새로고침: " + windows.Count.ToString(CultureInfo.InvariantCulture) + "개");
            }
        }

        private string BuildWindowSnapshot(List<WindowInfo> windows)
        {
            var builder = new StringBuilder();
            foreach (WindowInfo info in windows)
            {
                builder.Append(info.Handle.ToInt64().ToString(CultureInfo.InvariantCulture));
                builder.Append('|');
                builder.Append(info.Title);
                builder.Append('|');
                builder.Append(info.ProcessName);
                builder.Append('|');
                builder.Append(info.ProcessId.ToString(CultureInfo.InvariantCulture));
                builder.Append(';');
            }
            return builder.ToString();
        }

        private void RefreshAudioSources()
        {
            AudioCaptureMode mode = ResolveAudioCaptureMode();
            bool available = mode != AudioCaptureMode.None;
            if (audioStatusText != null)
            {
                audioStatusText.Text = mode == AudioCaptureMode.Process
                    ? "선택한 앱 소리만 자동 녹음"
                    : mode == AudioCaptureMode.SystemLoopback
                        ? "기본 출력 소리 자동 녹음"
                        : "소리 녹음 준비 안 됨";
                audioStatusText.Foreground = available ? Brush("#1d4ed8") : Brush("#be123c");
            }

            AppendLog(mode == AudioCaptureMode.Process
                ? "소리 녹음: 선택한 앱 우선 사용"
                : mode == AudioCaptureMode.SystemLoopback
                    ? "소리 녹음: 기본 출력 소리 사용"
                    : "소리 녹음 준비 안 됨");
        }

        private void OnWindowSelected(object sender, SelectionChangedEventArgs e)
        {
            selectedWindow = windowList.SelectedItem as WindowInfo;
            if (selectedWindow == null)
            {
                targetText.Text = "창을 선택해주세요";
                return;
            }
            targetText.Text = selectedWindow.ProcessName + "  " + selectedWindow.SizeText;
        }

        private void ToggleRecordingAction()
        {
            if (isPausing)
            {
                if (isPaused && activeRecording != null)
                {
                    pendingResumeAfterPause = true;
                    startButton.Content = "다시 시작 준비 중";
                    SetStatus("일시정지 정리 후 다시 시작할게요");
                }
                return;
            }

            if (activeRecording == null)
            {
                StartRecording();
                return;
            }

            if (isPaused)
            {
                ResumeRecording();
            }
            else
            {
                PauseRecording();
            }
        }

        private void StartRecording()
        {
            if (isStopping)
            {
                SetStatus("이전 녹화를 정리하는 중이에요");
                return;
            }

            if (selectedWindow == null)
            {
                SetStatus("먼저 녹화할 창을 선택해주세요");
                return;
            }

            if (!NativeMethods.IsWindow(selectedWindow.Handle))
            {
                SetStatus("선택한 창을 찾을 수 없어요");
                RefreshWindows();
                return;
            }

            string saveDir = GetSaveFolder();
            if (saveDir == null) return;

            if (!Directory.Exists(saveDir)) Directory.CreateDirectory(saveDir);
            SaveSettings(saveDir);

            AudioCaptureMode audioMode = ResolveAudioCaptureMode();
            if (audioMode == AudioCaptureMode.None)
            {
                AppendLog("소리 녹음 준비 안 됨. 화면만 녹화합니다.");
            }
            else if (audioMode == AudioCaptureMode.SystemLoopback)
            {
                AppendLog("이 PC에서는 앱별 소리 캡처가 지원되지 않아 기본 출력 소리로 녹음합니다.");
            }

            string baseName = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss", CultureInfo.InvariantCulture);
            string finalPath = UniquePath(saveDir, baseName, ".mp4");

            savedBounds = selectedWindow.Bounds;

            try
            {
                var recording = new RecordingState
                {
                    FinalPath = finalPath,
                    TempDirectory = PrepareRecordingTempDirectory(),
                    Target = selectedWindow,
                    Fps = (int)fpsSlider.Value,
                    DrawCursor = true,
                    AudioMode = audioMode,
                    BoostAudio = false,
                    AudioGain = 1.0,
                    CurrentAudioGain = 1.0,
                    AudioSyncDelaySeconds = GetAudioSyncDelaySeconds()
                };

                activeRecording = recording;
                StartRecordingSegment(recording);
                ApplySilentPlaybackIfRecording();

                isPaused = false;
                pendingResumeAfterPause = false;
                startButton.Content = "일시정지";
                startButton.IsEnabled = true;
                stopButton.IsEnabled = true;
                saveFolderBox.IsEnabled = false;
                fpsSlider.IsEnabled = false;
                audioSyncSlider.IsEnabled = false;
                windowList.IsEnabled = false;
                outputText.Text = "저장 파일: " + finalPath;
                SetStatus(recording.HasLoopbackAudio ? "녹화 중" : "화면 녹화 중");
                AppendLog("녹화 시작: " + finalPath);
            }
            catch (Exception ex)
            {
                AppendLog("녹화를 시작하지 못했어요: " + ex.Message);
                StopProcessesOnly();
                activeRecording = null;
                fpsSlider.IsEnabled = true;
                audioSyncSlider.IsEnabled = true;
                windowList.IsEnabled = true;
                SetStatus("시작 실패");
            }
        }

        private void PauseRecording()
        {
            if (activeRecording == null || isStopping || isPausing) return;

            var recording = activeRecording;
            bool keepReducedMode = silentPlaybackToggle != null && silentPlaybackToggle.IsChecked == true;

            isPausing = true;
            isPaused = true;
            pendingResumeAfterPause = false;
            startButton.Content = "다시 시작";
            startButton.IsEnabled = true;
            stopButton.IsEnabled = true;
            SetStatus("일시정지 처리 중...");
            AppendLog("녹화를 일시정지하는 중...");

            ThreadPool.QueueUserWorkItem(delegate
            {
                Exception stopError = null;
                try
                {
                    StopCurrentSegment(recording, 1800, 1200, keepReducedMode);
                }
                catch (Exception ex)
                {
                    stopError = ex;
                }

                Dispatcher.BeginInvoke(new Action(delegate
                {
                    isPausing = false;
                    if (stopError != null)
                    {
                        AppendLog("일시정지 처리 중 오류가 났어요: " + stopError.Message);
                    }

                    if (activeRecording != recording || isStopping)
                    {
                        pendingResumeAfterPause = false;
                        return;
                    }

                    if (pendingResumeAfterPause)
                    {
                        pendingResumeAfterPause = false;
                        ResumeRecording();
                        return;
                    }

                    isPaused = true;
                    startButton.Content = "다시 시작";
                    startButton.IsEnabled = true;
                    stopButton.IsEnabled = true;
                    SetStatus("일시정지");
                    AppendLog("녹화를 일시정지했어요");
                }));
            });
        }

        private void ResumeRecording()
        {
            if (activeRecording == null || isStopping || isPausing) return;

            try
            {
                StartRecordingSegment(activeRecording);
                lastSilentPlaybackAttemptUtc = DateTime.MinValue;
                ApplySilentPlaybackIfRecording();
                isPaused = false;
                pendingResumeAfterPause = false;
                startButton.Content = "일시정지";
                startButton.IsEnabled = true;
                stopButton.IsEnabled = true;
                SetStatus("녹화 중");
                AppendLog("녹화를 다시 시작했어요");
            }
            catch (Exception ex)
            {
                AppendLog("녹화를 다시 시작하지 못했어요: " + ex.Message);
                SetStatus("다시 시작 실패");
            }
        }

        private void StartRecordingSegment(RecordingState recording)
        {
            recording.SegmentIndex++;
            string stem = Path.GetFileNameWithoutExtension(recording.FinalPath);
            string segmentBase = Path.Combine(recording.TempDirectory, stem + ".part" + recording.SegmentIndex.ToString("000", CultureInfo.InvariantCulture));
            string videoPath = segmentBase + ".video.mkv";
            string audioPath = recording.AudioMode != AudioCaptureMode.None ? segmentBase + ".audio.wav" : null;
            bool isFirstSegment = recording.Segments.Count == 0;

            bool audioStarted = false;
            DateTime audioStartedUtc = DateTime.MinValue;
            if (recording.AudioMode != AudioCaptureMode.None)
            {
                try
                {
                    if (recording.AudioMode == AudioCaptureMode.Process)
                    {
                        audioProcess = StartTargetAudio(audioPath, recording.Target.ProcessId, out audioStartedUtc);
                    }
                    else
                    {
                        audioProcess = StartLoopbackAudio(audioPath, null, false, out audioStartedUtc);
                    }
                    audioStarted = true;
                    recording.HasLoopbackAudio = true;
                }
                catch (Exception audioEx)
                {
                    audioStartedUtc = DateTime.MinValue;
                    AppendLog("선택한 앱 소리 녹음 실패: " + audioEx.Message);
                    if (recording.AudioMode == AudioCaptureMode.Process && IsSystemAudioCaptureAvailable())
                    {
                        try
                        {
                            recording.AudioMode = AudioCaptureMode.SystemLoopback;
                            audioProcess = StartLoopbackAudio(audioPath, null, false, out audioStartedUtc);
                            audioStarted = true;
                            recording.HasLoopbackAudio = true;
                            AppendLog("기본 출력 소리 녹음으로 대신 저장합니다.");
                        }
                        catch (Exception fallbackEx)
                        {
                            audioStartedUtc = DateTime.MinValue;
                            audioProcess = null;
                            AppendLog("기본 출력 소리 녹음도 실패: " + fallbackEx.Message);
                            AppendLog("이 구간은 소리 없이 화면만 녹화합니다.");
                        }
                    }
                    else
                    {
                        audioProcess = null;
                        AppendLog("이 구간은 소리 없이 화면만 녹화합니다.");
                    }
                }
            }

            DateTime videoStartedUtc;
            ffmpegProcess = StartFfmpegCapture(recording.Target, videoPath, recording.Fps, recording.DrawCursor, false, out videoStartedUtc);

            var segment = new RecordingSegment
            {
                VideoPath = videoPath,
                AudioPath = audioStarted ? audioPath : null,
                HasLoopbackAudio = audioStarted,
                BoostAudio = recording.CurrentAudioGain > 1.01,
                AudioGain = recording.CurrentAudioGain,
                AudioSyncDelaySeconds = recording.AudioSyncDelaySeconds,
                Fps = recording.Fps,
                TrimStartSeconds = isFirstSegment ? StartWarmupTrimSeconds : 0,
                VideoStartedUtc = videoStartedUtc,
                AudioStartedUtc = audioStarted ? audioStartedUtc : DateTime.MinValue
            };

            recording.CurrentSegment = segment;
            recording.SegmentStartedUtc = videoStartedUtc;
            recording.Segments.Add(segment);

            if (audioStarted && audioStartedUtc != DateTime.MinValue)
            {
                double syncMs = (videoStartedUtc - audioStartedUtc).TotalMilliseconds;
                if (Math.Abs(syncMs) >= 20)
                {
                    AppendLog("싱크 보정: 소리 시작 차이 " + syncMs.ToString("0", CultureInfo.InvariantCulture) + "ms");
                }
                double manualSyncMs = segment.AudioSyncDelaySeconds * 1000.0;
                AppendLog("싱크 보정: 저장 시 소리 위치 " + manualSyncMs.ToString("+0;-0;0", CultureInfo.InvariantCulture) + "ms 적용");
            }
        }

        private void StopCurrentSegment()
        {
            StopCurrentSegment(activeRecording, 3500, 2000, false);
        }

        private void StopCurrentSegment(RecordingState recording, int videoWaitMs, int audioWaitMs, bool keepReducedMode)
        {
            EndAudioBoostRange(recording, keepReducedMode);
            var videoProcess = ffmpegProcess;
            var soundProcess = audioProcess;
            ffmpegProcess = null;
            audioProcess = null;
            StopProcessNicely(videoProcess, videoWaitMs);
            StopProcessNicely(soundProcess, audioWaitMs);
            if (recording != null) recording.CurrentSegment = null;
        }

        private Process StartFfmpegCapture(WindowInfo target, string outputPath, int fps, bool drawCursor, bool finalOutput, out DateTime captureStartedUtc)
        {
            captureStartedUtc = DateTime.MinValue;
            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-y");
            args.Add("-stats_period");
            args.Add("0.10");

            if (gfxCaptureAvailable)
            {
                string source = string.Format(
                    CultureInfo.InvariantCulture,
                    "gfxcapture=hwnd={0}:capture_cursor={1}:display_border=false:max_framerate={2}:width=-2:height=-2",
                    target.Handle.ToInt64(),
                    drawCursor ? "true" : "false",
                    fps);
                args.Add("-f");
                args.Add("lavfi");
                args.Add("-i");
                args.Add(source);
                args.Add("-map");
                args.Add("0:v:0");
                args.Add("-an");
                args.Add("-vf");
                args.Add("hwdownload,format=bgra,pad=ceil(iw/2)*2:ceil(ih/2)*2,format=yuv420p");
            }
            else
            {
                args.Add("-thread_queue_size");
                args.Add("512");
                args.Add("-f");
                args.Add("gdigrab");
                args.Add("-draw_mouse");
                args.Add(drawCursor ? "1" : "0");
                args.Add("-framerate");
                args.Add(fps.ToString(CultureInfo.InvariantCulture));
                args.Add("-i");
                args.Add("title=" + target.Title);
                args.Add("-map");
                args.Add("0:v:0");
                args.Add("-an");
                args.Add("-vf");
                args.Add("pad=ceil(iw/2)*2:ceil(ih/2)*2,format=yuv420p");
            }

            args.Add("-c:v");
            args.Add("libx264");
            args.Add("-preset");
            args.Add("veryfast");
            args.Add("-crf");
            args.Add("23");
            args.Add(outputPath);

            DateTime processStartedUtc = DateTime.UtcNow;
            var ready = new ManualResetEventSlim(false);
            DateTime readyUtc = DateTime.MinValue;
            var process = StartProcess(GetFfmpegPath(), args, "[화면] ", delegate(string line)
            {
                DateTime estimatedStart = EstimateFfmpegCaptureStartUtc(line, fps);
                if (estimatedStart != DateTime.MinValue && readyUtc == DateTime.MinValue)
                {
                    readyUtc = estimatedStart;
                    ready.Set();
                }
            });

            if (!WaitForVideoReady(process, ready, VideoReadyWaitMilliseconds))
            {
                if (process.HasExited)
                {
                    throw new InvalidOperationException("화면 캡처가 바로 종료됐어요. 창이 열려 있는지 확인해주세요.");
                }
                readyUtc = processStartedUtc;
                AppendLog("화면 첫 프레임 신호를 받지 못해 실행 시각으로 싱크를 맞춥니다.");
            }

            captureStartedUtc = readyUtc;
            return process;
        }

        private bool WaitForVideoReady(Process process, ManualResetEventSlim ready, int waitMs)
        {
            int waited = 0;
            while (waited < waitMs)
            {
                if (ready.Wait(100)) return true;
                if (process.HasExited) return false;
                waited += 100;
            }
            return ready.IsSet;
        }

        private DateTime EstimateFfmpegCaptureStartUtc(string line, int fps)
        {
            if (string.IsNullOrEmpty(line) || line.IndexOf("frame=", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return DateTime.MinValue;
            }

            int frameCount = ParseFfmpegProgressFrameCount(line);
            double seconds;
            if (TryParseFfmpegProgressSeconds(line, out seconds))
            {
                if (seconds <= 0 && frameCount <= 0) return DateTime.MinValue;
                return DateTime.UtcNow.AddSeconds(-seconds);
            }

            if (frameCount <= 0) return DateTime.MinValue;
            double frameSeconds = Math.Max(0, (frameCount - 1) / (double)Math.Max(1, fps));
            return DateTime.UtcNow.AddSeconds(-frameSeconds);
        }

        private int ParseFfmpegProgressFrameCount(string line)
        {
            Match match = Regex.Match(line, @"frame=\s*(\d+)", RegexOptions.IgnoreCase);
            if (!match.Success) return -1;
            int frameCount;
            if (!int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out frameCount)) return -1;
            return frameCount;
        }

        private bool TryParseFfmpegProgressSeconds(string line, out double seconds)
        {
            seconds = 0;
            Match match = Regex.Match(line, @"time=(\d+):(\d+):(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);
            if (!match.Success) return false;

            int hours;
            int minutes;
            double secs;
            if (!int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out hours)) return false;
            if (!int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out minutes)) return false;
            if (!double.TryParse(match.Groups[3].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out secs)) return false;

            seconds = hours * 3600.0 + minutes * 60.0 + secs;
            return seconds >= 0;
        }

        private Process StartLoopbackAudio(string audioPath, string sourceName, bool monitorOn, out DateTime captureStartedUtc)
        {
            captureStartedUtc = DateTime.MinValue;
            string fileName;
            var args = new List<string>();
            string helperPath = GetAudioHelperPath();
            if (!string.IsNullOrEmpty(helperPath))
            {
                fileName = helperPath;
            }
            else
            {
                string script = Path.Combine(supportDir, "loopback_audio_recorder.py");
                if (!File.Exists(script)) script = Path.Combine(appDir, "loopback_audio_recorder.py");
                fileName = "python";
                args.Add(script);
            }
            args.Add(audioPath);
            if (!string.IsNullOrEmpty(sourceName))
            {
                args.Add("--source");
                args.Add(sourceName);
            }
            if (monitorOn) args.Add("--monitor-on");

            var ready = new ManualResetEventSlim(false);
            DateTime readyUtc = DateTime.MinValue;
            var process = StartProcess(fileName, args, "[소리] ", delegate(string line)
            {
                if (line.IndexOf("loopback audio capture started", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    readyUtc = DateTime.UtcNow;
                    ready.Set();
                }
            });

            if (!WaitForAudioReady(process, ready, AudioReadyWaitMilliseconds))
            {
                if (process.HasExited)
                {
                    throw new InvalidOperationException("소리 녹음이 바로 종료됐어요. Windows 기본 출력 장치와 녹화할 앱의 음소거 상태를 확인해주세요.");
                }
                readyUtc = DateTime.UtcNow;
                AppendLog("소리 녹음 시작 신호를 받지 못해 현재 시각으로 싱크를 맞춥니다.");
            }
            captureStartedUtc = readyUtc;
            return process;
        }

        private Process StartTargetAudio(string audioPath, int processId, out DateTime captureStartedUtc)
        {
            captureStartedUtc = DateTime.MinValue;
            string fileName;
            var args = new List<string>();
            string helperPath = GetProcessAudioHelperPath();
            if (!string.IsNullOrEmpty(helperPath))
            {
                fileName = helperPath;
            }
            else
            {
                string script = Path.Combine(supportDir, "process_audio_recorder.py");
                if (!File.Exists(script)) script = Path.Combine(appDir, "process_audio_recorder.py");
                if (!File.Exists(script))
                {
                    throw new InvalidOperationException("선택한 앱 소리 녹음 파일을 찾지 못했어요.");
                }
                fileName = "python";
                args.Add(script);
            }

            args.Add(audioPath);
            args.Add("--pid");
            args.Add(processId.ToString(CultureInfo.InvariantCulture));

            var ready = new ManualResetEventSlim(false);
            DateTime readyUtc = DateTime.MinValue;
            var process = StartProcess(fileName, args, "[소리] ", delegate(string line)
            {
                if (line.IndexOf("process audio capture started", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    readyUtc = DateTime.UtcNow;
                    ready.Set();
                }
            });

            if (!WaitForAudioReady(process, ready, AudioReadyWaitMilliseconds))
            {
                if (process.HasExited)
                {
                    throw new InvalidOperationException("선택한 앱 소리 녹음이 바로 종료됐어요. 앱에서 실제로 소리가 나오는지 확인해주세요.");
                }
                readyUtc = DateTime.UtcNow;
                AppendLog("앱 소리 녹음 시작 신호를 받지 못해 현재 시각으로 싱크를 맞춥니다.");
            }
            captureStartedUtc = readyUtc;
            return process;
        }

        private bool WaitForAudioReady(Process process, ManualResetEventSlim ready, int waitMs)
        {
            int waited = 0;
            while (waited < waitMs)
            {
                if (ready.Wait(100)) return true;
                if (process.HasExited) return false;
                waited += 100;
            }
            return ready.IsSet;
        }

        private Process StartProcess(string fileName, List<string> args, string logPrefix)
        {
            return StartProcess(fileName, args, logPrefix, null);
        }

        private Process StartProcess(string fileName, List<string> args, string logPrefix, Action<string> onOutputLine)
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = JoinArgs(args),
                WorkingDirectory = supportDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            psi.StandardOutputEncoding = Encoding.UTF8;
            psi.StandardErrorEncoding = Encoding.UTF8;
            ApplyPythonUtf8Environment(psi);

            var process = new Process();
            process.StartInfo = psi;
            process.EnableRaisingEvents = true;
            process.OutputDataReceived += delegate(object sender, DataReceivedEventArgs e)
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    if (onOutputLine != null) onOutputLine(e.Data);
                    QueueProcessLog(logPrefix, e.Data);
                }
            };
            process.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs e)
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    if (onOutputLine != null) onOutputLine(e.Data);
                    QueueProcessLog(logPrefix, e.Data);
                }
            };

            if (!process.Start()) throw new InvalidOperationException(fileName + " 실행에 실패했어요");
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            return process;
        }

        private void StopRecordingFromUi()
        {
            StopRecording();
        }

        private void StopRecording()
        {
            if (isStopping) return;
            var recording = activeRecording;

            isStopping = true;
            activeRecording = null;

            startButton.IsEnabled = false;
            stopButton.IsEnabled = false;
            stopButton.Content = "처리 중";
            isPausing = false;
            isPaused = false;
            pendingResumeAfterPause = false;
            startButton.Content = "녹화 시작";
            saveFolderBox.IsEnabled = true;
            if (targetWindowHidden) RestoreTargetWindowFromHidden();
            SetStatus("녹화 종료 처리 중...");
            AppendLog("녹화 종료 처리 중...");

            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    StopCurrentSegment(recording, 3500, 2000, false);
                    RestorePlaybackVolume();

                    if (recording != null)
                    {
                        FinalizeRecording(recording);
                    }
                }
                catch (Exception ex)
                {
                    QueueUiLog("녹화 종료 처리 중 오류가 났어요: " + ex.Message);
                }
                finally
                {
                    try
                    {
                        Dispatcher.BeginInvoke(new Action(delegate
                        {
                            isStopping = false;
                            startButton.IsEnabled = true;
                            stopButton.IsEnabled = false;
                            stopButton.Content = "녹화 종료";
                            startButton.Content = "녹화 시작";
                            saveFolderBox.IsEnabled = true;
                            fpsSlider.IsEnabled = true;
                            audioSyncSlider.IsEnabled = true;
                            silentPlaybackToggle.IsEnabled = true;
                            windowList.IsEnabled = true;
                            SetStatus("준비됨");
                        }));
                    }
                    catch { }
                }
            });
        }

        private void StopProcessesOnly()
        {
            StopCurrentSegment(activeRecording, 3500, 2000, false);
            RestorePlaybackVolume();
        }

        private void StopProcessNicely(Process process)
        {
            StopProcessNicely(process, 3500);
        }

        private void StopProcessNicely(Process process, int waitMs)
        {
            if (process == null) return;
            try
            {
                if (!process.HasExited)
                {
                    try { process.StandardInput.WriteLine("q"); } catch { }
                    if (!process.WaitForExit(waitMs))
                    {
                        process.Kill();
                        process.WaitForExit(2000);
                    }
                }
            }
            catch
            {
                try { if (!process.HasExited) process.Kill(); } catch { }
            }
        }

        private void FinalizeRecording(RecordingState recording)
        {
            try
            {
                if (recording.Segments.Count == 0)
                {
                    QueueUiLog("저장할 녹화 구간이 없어요.");
                    return;
                }

                bool allSegmentsHaveAudio = recording.HasLoopbackAudio;
                foreach (RecordingSegment segment in recording.Segments)
                {
                    allSegmentsHaveAudio = allSegmentsHaveAudio && segment.HasLoopbackAudio && File.Exists(segment.AudioPath);
                }

                if (allSegmentsHaveAudio)
                {
                    FinalizeAudioVideoRecording(recording);
                }
                else
                {
                    if (recording.HasLoopbackAudio)
                    {
                        QueueUiLog("일부 구간의 소리 녹음이 없어 화면만 저장합니다.");
                    }
                    FinalizeVideoOnlyRecording(recording);
                }
            }
            catch (Exception ex)
            {
                QueueUiLog("파일을 저장하지 못했어요: " + ex.Message);
            }
        }

        private void FinalizeAudioVideoRecording(RecordingState recording)
        {
            QueueUiLog("영상과 소리를 합치는 중...");
            var muxedSegments = new List<string>();
            string err;

            foreach (RecordingSegment segment in recording.Segments)
            {
                string muxedPath = Path.ChangeExtension(segment.VideoPath, ".mux.mp4");
                int code = RunMux(segment, muxedPath, true, out err);
                if (code != 0 && segment.AudioBoostRanges.Count > 0)
                {
                    QueueUiLog("소리 보정 처리에 실패해서 기본 방식으로 다시 저장합니다.");
                    code = RunMux(segment, muxedPath, false, out err);
                }
                if (code != 0)
                {
                    QueueUiLog("파일을 합치지 못했어요: " + err);
                    return;
                }
                muxedSegments.Add(muxedPath);
            }

            int concatCode = muxedSegments.Count == 1
                ? RunRemux(muxedSegments[0], recording.FinalPath, out err)
                : RunConcatAudioVideo(muxedSegments, recording.FinalPath, out err);

            if (concatCode == 0)
            {
                CleanupSegments(recording, muxedSegments);
                MarkRecordingSaved(recording.FinalPath);
            }
            else
            {
                QueueUiLog("파일을 합치지 못했어요: " + err);
            }
        }

        private int RunMux(RecordingSegment segment, string outputPath, bool applyAudioFilters, out string err)
        {
            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-y");
            double audioTrimSeconds = GetAudioTrimSeconds(segment);
            double audioDelaySeconds = GetAudioDelaySeconds(segment);
            args.Add("-i");
            args.Add(segment.VideoPath);
            args.Add("-i");
            args.Add(segment.AudioPath);
            args.Add("-filter_complex");
            args.Add(BuildMuxFilterComplex(segment, applyAudioFilters, audioTrimSeconds, audioDelaySeconds));
            args.Add("-map");
            args.Add("[v]");
            args.Add("-map");
            args.Add("[a]");
            args.Add("-c:v");
            args.Add("libx264");
            args.Add("-preset");
            args.Add("veryfast");
            args.Add("-crf");
            args.Add("23");
            args.Add("-pix_fmt");
            args.Add("yuv420p");
            args.Add("-c:a");
            args.Add("aac");
            args.Add("-b:a");
            args.Add("160k");
            args.Add("-shortest");
            args.Add("-movflags");
            args.Add("+faststart");
            args.Add(outputPath);
            return RunFfmpeg(args, out err);
        }

        private string BuildMuxFilterComplex(RecordingSegment segment, bool applyAudioFilters, double audioTrimSeconds, double audioDelaySeconds)
        {
            var videoFilters = new List<string>();
            if (segment.TrimStartSeconds > 0)
            {
                videoFilters.Add("trim=start=" + FormatSeconds(segment.TrimStartSeconds));
            }
            videoFilters.Add("setpts=PTS-STARTPTS");
            int outputFps = GetSegmentOutputFps(segment);
            videoFilters.Add("fps=fps=" + outputFps.ToString(CultureInfo.InvariantCulture));
            videoFilters.Add("setpts=N/(" + outputFps.ToString(CultureInfo.InvariantCulture) + "*TB)");

            var audioFilters = new List<string>();
            if (audioTrimSeconds > 0)
            {
                audioFilters.Add("atrim=start=" + FormatSeconds(audioTrimSeconds));
            }
            audioFilters.Add("asetpts=PTS-STARTPTS");
            audioFilters.AddRange(BuildAudioVolumeFilters(segment, applyAudioFilters, audioTrimSeconds));
            if (audioFilters.Count > 1 && audioFilters[audioFilters.Count - 1] != "asetpts=PTS-STARTPTS")
            {
                audioFilters.Add("alimiter=limit=0.98");
            }
            if (audioDelaySeconds > 0.001)
            {
                int delayMs = Math.Max(1, (int)Math.Round(audioDelaySeconds * 1000.0));
                audioFilters.Add("adelay=" + delayMs.ToString(CultureInfo.InvariantCulture) + ":all=1");
            }
            audioFilters.Add("aresample=async=1:first_pts=0");
            audioFilters.Add("apad");

            return "[0:v]" + string.Join(",", videoFilters.ToArray()) + "[v];[1:a]" + string.Join(",", audioFilters.ToArray()) + "[a]";
        }

        private int GetSegmentOutputFps(RecordingSegment segment)
        {
            if (segment == null || segment.Fps <= 0) return 60;
            return Math.Max(5, Math.Min(60, segment.Fps));
        }

        private string BuildAudioFilter(RecordingSegment segment, bool applyAudioFilters, double audioDelaySeconds)
        {
            var filters = BuildAudioVolumeFilters(segment, applyAudioFilters, segment.TrimStartSeconds);
            if (audioDelaySeconds > 0.001)
            {
                int delayMs = Math.Max(1, (int)Math.Round(audioDelaySeconds * 1000.0));
                filters.Insert(0, "adelay=" + delayMs.ToString(CultureInfo.InvariantCulture) + ":all=1");
            }
            if (filters.Count > 0)
            {
                filters.Add("alimiter=limit=0.98");
            }
            filters.Add("apad");
            return string.Join(",", filters.ToArray());
        }

        private double GetAudioTrimSeconds(RecordingSegment segment)
        {
            double raw = GetAdjustedAudioSeekSeconds(segment);
            return raw > 0 ? raw : 0;
        }

        private double GetAudioDelaySeconds(RecordingSegment segment)
        {
            double raw = GetAdjustedAudioSeekSeconds(segment);
            return raw < 0 ? -raw : 0;
        }

        private double GetAdjustedAudioSeekSeconds(RecordingSegment segment)
        {
            if (segment == null) return 0;
            return GetRawAudioSeekSeconds(segment) - segment.AudioSyncDelaySeconds;
        }

        private double GetRawAudioSeekSeconds(RecordingSegment segment)
        {
            if (segment == null || segment.AudioStartedUtc == DateTime.MinValue || segment.VideoStartedUtc == DateTime.MinValue)
            {
                return segment == null ? 0 : segment.TrimStartSeconds;
            }

            double audioStartedBeforeVideo = (segment.VideoStartedUtc - segment.AudioStartedUtc).TotalSeconds;
            return segment.TrimStartSeconds + audioStartedBeforeVideo;
        }

        private List<string> BuildAudioVolumeFilters(RecordingSegment segment, bool applyAudioFilters, double offsetSeconds)
        {
            var filters = new List<string>();
            if (!applyAudioFilters || segment.AudioBoostRanges.Count == 0) return filters;

            foreach (AudioBoostRange range in segment.AudioBoostRanges)
            {
                double start = Math.Max(0, range.StartSeconds - offsetSeconds);
                double end = Math.Max(0, range.EndSeconds - offsetSeconds);
                if (end <= start + 0.05) continue;

                double fade = Math.Min(AudioBoostFadeSeconds, Math.Max(0.05, (end - start) / 2.0));
                double fadeOutStart = Math.Max(start, end - fade);
                string gain = ReducedPlaybackGain.ToString("0.###", CultureInfo.InvariantCulture);
                string startText = FormatSeconds(start);
                string endText = FormatSeconds(end);
                string fadeText = FormatSeconds(fade);
                string fadeOutStartText = FormatSeconds(fadeOutStart);

                string expression =
                    "if(lt(t\\," + startText + ")\\,1\\," +
                    "if(lt(t\\," + (start + fade).ToString("0.###", CultureInfo.InvariantCulture) + ")\\,1+(" + gain + "-1)*(t-" + startText + ")/" + fadeText + "\\," +
                    "if(lt(t\\," + fadeOutStartText + ")\\," + gain + "\\," +
                    "if(lt(t\\," + endText + ")\\,1+(" + gain + "-1)*(" + endText + "-t)/" + fadeText + "\\,1))))";

                filters.Add("volume='" + expression + "':eval=frame");
            }

            return filters;
        }

        private string BuildAudioVideoTimelineFilter(RecordingSegment segment, bool applyAudioFilters)
        {
            string keepExpression = BuildKeepExpression(segment);
            string videoFilter = "[0:v]select='" + keepExpression + "',setpts=N/FRAME_RATE/TB[v]";

            var audioFilters = BuildAudioVolumeFilters(segment, applyAudioFilters, 0);
            if (audioFilters.Count > 0) audioFilters.Add("alimiter=limit=0.98");
            audioFilters.Add("aselect='" + keepExpression + "'");
            audioFilters.Add("asetpts=N/SR/TB");
            string audioFilter = "[1:a]" + string.Join(",", audioFilters.ToArray()) + "[a]";

            return videoFilter + ";" + audioFilter;
        }

        private string BuildKeepExpression(RecordingSegment segment)
        {
            var terms = new List<string>();
            if (segment.TrimStartSeconds > 0)
            {
                terms.Add("gte(t\\," + FormatSeconds(segment.TrimStartSeconds) + ")");
            }

            foreach (PauseRange range in segment.PauseRanges)
            {
                if (range.EndSeconds <= range.StartSeconds + 0.05) continue;
                terms.Add("not(between(t\\," + FormatSeconds(range.StartSeconds) + "\\," + FormatSeconds(range.EndSeconds) + "))");
            }

            return terms.Count == 0 ? "1" : string.Join("*", terms.ToArray());
        }

        private void FinalizeVideoOnlyRecording(RecordingState recording)
        {
            QueueUiLog("영상 파일을 저장하는 중...");
            string err;
            var videoSegments = new List<string>();
            var extraPaths = new List<string>();
            foreach (RecordingSegment segment in recording.Segments)
            {
                if (segment.PauseRanges.Count > 0)
                {
                    string filteredPath = Path.ChangeExtension(segment.VideoPath, ".filtered.mp4");
                    int filterCode = RunFilterVideoTimeline(segment.VideoPath, filteredPath, segment, out err);
                    if (filterCode != 0)
                    {
                        QueueUiLog("일시정지 구간 정리에 실패해서 기본 영상으로 저장합니다.");
                        videoSegments.Add(segment.VideoPath);
                    }
                    else
                    {
                        videoSegments.Add(filteredPath);
                        extraPaths.Add(filteredPath);
                    }
                }
                else if (segment.TrimStartSeconds > 0)
                {
                    string trimmedPath = Path.ChangeExtension(segment.VideoPath, ".trim.mp4");
                    int trimCode = RunTrimVideo(segment.VideoPath, trimmedPath, segment.TrimStartSeconds, out err);
                    if (trimCode != 0)
                    {
                        QueueUiLog("시작 화면 보정에 실패해서 기본 영상으로 저장합니다.");
                        videoSegments.Add(segment.VideoPath);
                    }
                    else
                    {
                        videoSegments.Add(trimmedPath);
                        extraPaths.Add(trimmedPath);
                    }
                }
                else
                {
                    videoSegments.Add(segment.VideoPath);
                }
            }

            int code = videoSegments.Count == 1
                ? RunRemux(videoSegments[0], recording.FinalPath, out err)
                : RunConcat(videoSegments, recording.FinalPath, false, out err);

            if (code == 0)
            {
                CleanupSegments(recording, extraPaths);
                MarkRecordingSaved(recording.FinalPath);
            }
            else
            {
                QueueUiLog("영상 파일을 저장하지 못했어요: " + err);
            }
        }

        private int RunFilterVideoTimeline(string inputPath, string outputPath, RecordingSegment segment, out string err)
        {
            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-y");
            args.Add("-i");
            args.Add(inputPath);
            args.Add("-vf");
            args.Add("select='" + BuildKeepExpression(segment) + "',setpts=N/FRAME_RATE/TB");
            args.Add("-map");
            args.Add("0:v:0");
            args.Add("-c:v");
            args.Add("libx264");
            args.Add("-preset");
            args.Add("veryfast");
            args.Add("-crf");
            args.Add("23");
            args.Add("-pix_fmt");
            args.Add("yuv420p");
            args.Add("-an");
            args.Add("-movflags");
            args.Add("+faststart");
            args.Add(outputPath);
            return RunFfmpeg(args, out err);
        }

        private int RunTrimVideo(string inputPath, string outputPath, double trimStartSeconds, out string err)
        {
            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-y");
            args.Add("-ss");
            args.Add(FormatSeconds(trimStartSeconds));
            args.Add("-i");
            args.Add(inputPath);
            args.Add("-map");
            args.Add("0:v:0");
            args.Add("-c:v");
            args.Add("libx264");
            args.Add("-preset");
            args.Add("veryfast");
            args.Add("-crf");
            args.Add("23");
            args.Add("-pix_fmt");
            args.Add("yuv420p");
            args.Add("-an");
            args.Add("-movflags");
            args.Add("+faststart");
            args.Add(outputPath);
            return RunFfmpeg(args, out err);
        }

        private int RunRemux(string inputPath, string outputPath, out string err)
        {
            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-y");
            args.Add("-i");
            args.Add(inputPath);
            args.Add("-c");
            args.Add("copy");
            args.Add("-movflags");
            args.Add("+faststart");
            args.Add(outputPath);
            return RunFfmpeg(args, out err);
        }

        private string FormatSeconds(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private int RunConcat(List<string> inputPaths, string outputPath, bool hasAudio, out string err)
        {
            string listPath = Path.Combine(Path.GetDirectoryName(outputPath), Path.GetFileNameWithoutExtension(outputPath) + ".concat.txt");
            File.WriteAllText(listPath, BuildConcatList(inputPaths), new UTF8Encoding(false));
            try
            {
                var args = new List<string>();
                args.Add("-hide_banner");
                args.Add("-y");
                args.Add("-fflags");
                args.Add("+genpts");
                args.Add("-f");
                args.Add("concat");
                args.Add("-safe");
                args.Add("0");
                args.Add("-i");
                args.Add(listPath);
                args.Add("-c:v");
                args.Add("libx264");
                args.Add("-preset");
                args.Add("veryfast");
                args.Add("-crf");
                args.Add("23");
                args.Add("-pix_fmt");
                args.Add("yuv420p");
                if (hasAudio)
                {
                    args.Add("-c:a");
                    args.Add("aac");
                    args.Add("-b:a");
                    args.Add("160k");
                }
                else
                {
                    args.Add("-an");
                }
                args.Add("-movflags");
                args.Add("+faststart");
                args.Add(outputPath);
                return RunFfmpeg(args, out err);
            }
            finally
            {
                TryDelete(listPath);
            }
        }

        private int RunConcatAudioVideo(List<string> inputPaths, string outputPath, out string err)
        {
            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-y");
            foreach (string path in inputPaths)
            {
                args.Add("-i");
                args.Add(path);
            }

            var filters = new List<string>();
            var concatInputs = new StringBuilder();
            for (int i = 0; i < inputPaths.Count; i++)
            {
                filters.Add("[" + i.ToString(CultureInfo.InvariantCulture) + ":v]setpts=PTS-STARTPTS[v" + i.ToString(CultureInfo.InvariantCulture) + "]");
                filters.Add("[" + i.ToString(CultureInfo.InvariantCulture) + ":a]aresample=async=1:first_pts=0,asetpts=PTS-STARTPTS[a" + i.ToString(CultureInfo.InvariantCulture) + "]");
                concatInputs.Append("[v");
                concatInputs.Append(i.ToString(CultureInfo.InvariantCulture));
                concatInputs.Append("][a");
                concatInputs.Append(i.ToString(CultureInfo.InvariantCulture));
                concatInputs.Append("]");
            }
            concatInputs.Append("concat=n=");
            concatInputs.Append(inputPaths.Count.ToString(CultureInfo.InvariantCulture));
            concatInputs.Append(":v=1:a=1[v][a]");
            filters.Add(concatInputs.ToString());

            args.Add("-filter_complex");
            args.Add(string.Join(";", filters.ToArray()));
            args.Add("-map");
            args.Add("[v]");
            args.Add("-map");
            args.Add("[a]");
            args.Add("-c:v");
            args.Add("libx264");
            args.Add("-preset");
            args.Add("veryfast");
            args.Add("-crf");
            args.Add("23");
            args.Add("-pix_fmt");
            args.Add("yuv420p");
            args.Add("-c:a");
            args.Add("aac");
            args.Add("-b:a");
            args.Add("160k");
            args.Add("-movflags");
            args.Add("+faststart");
            args.Add(outputPath);
            return RunFfmpeg(args, out err);
        }

        private string BuildConcatList(List<string> inputPaths)
        {
            var builder = new StringBuilder();
            foreach (string path in inputPaths)
            {
                builder.Append("file '");
                builder.Append(path.Replace("\\", "/").Replace("'", "'\\''"));
                builder.AppendLine("'");
            }
            return builder.ToString();
        }

        private int RunFfmpeg(List<string> args, out string err)
        {
            var psi = new ProcessStartInfo
            {
                FileName = GetFfmpegPath(),
                Arguments = JoinArgs(args),
                WorkingDirectory = supportDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true
            };
            var proc = Process.Start(psi);
            err = proc.StandardError.ReadToEnd();
            proc.WaitForExit();
            return proc.ExitCode;
        }

        private void CleanupSegments(RecordingState recording, List<string> extraPaths)
        {
            foreach (RecordingSegment segment in recording.Segments)
            {
                TryDelete(segment.VideoPath);
                TryDelete(segment.AudioPath);
            }
            if (extraPaths != null)
            {
                foreach (string path in extraPaths) TryDelete(path);
            }
            try
            {
                if (!string.IsNullOrEmpty(recording.TempDirectory) && Directory.Exists(recording.TempDirectory))
                {
                    Directory.Delete(recording.TempDirectory, true);
                }
            }
            catch { }
        }

        private void MarkRecordingSaved(string finalPath)
        {
            Dispatcher.BeginInvoke(new Action(delegate
            {
                outputText.Text = "저장 파일: " + finalPath;
                AppendLog("저장 완료: " + finalPath);
            }));
        }

        private void BringTargetFront()
        {
            if (!EnsureTarget()) return;
            if (targetWindowHidden)
            {
                RestoreTargetWindowFromHidden();
                return;
            }
            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_RESTORE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_TOP, 0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_SHOWWINDOW);
            NativeMethods.SetForegroundWindow(selectedWindow.Handle);
            SetStatus("창을 앞으로 가져왔어요");
        }

        private void HideTargetWindowForRecording()
        {
            if (targetWindowHidden) return;
            if (!EnsureTarget())
            {
                if (windowVisibilityToggle != null) windowVisibilityToggle.IsChecked = false;
                return;
            }

            savedBounds = NativeWindows.GetBounds(selectedWindow.Handle);
            var area = WinForms.Screen.PrimaryScreen.WorkingArea;
            int visibleStrip = Math.Min(140, Math.Max(80, savedBounds.Value.Width / 8));
            int x = area.Right - visibleStrip;
            int y = Math.Max(area.Top, Math.Min(savedBounds.Value.Top, area.Bottom - Math.Min(160, savedBounds.Value.Height)));

            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_SHOWNOACTIVATE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_TOPMOST, x, y, savedBounds.Value.Width, savedBounds.Value.Height,
                NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
            targetWindowHidden = true;
            if (windowVisibilityToggle != null) windowVisibilityToggle.Content = "녹화창 다시 띄우기";
            SetStatus("녹화창을 가장자리에 뒀어요");
        }

        private void RestoreTargetWindowFromHidden()
        {
            if (!targetWindowHidden)
            {
                if (windowVisibilityToggle != null) windowVisibilityToggle.Content = "녹화창 숨기기 (다른 작업 가능)";
                return;
            }

            if (selectedWindow == null || !NativeMethods.IsWindow(selectedWindow.Handle))
            {
                targetWindowHidden = false;
                if (windowVisibilityToggle != null)
                {
                    windowVisibilityToggle.Content = "녹화창 숨기기 (다른 작업 가능)";
                    windowVisibilityToggle.IsChecked = false;
                }
                return;
            }

            if (!savedBounds.HasValue)
            {
                targetWindowHidden = false;
                if (windowVisibilityToggle != null)
                {
                    windowVisibilityToggle.Content = "녹화창 숨기기 (다른 작업 가능)";
                    windowVisibilityToggle.IsChecked = false;
                }
                return;
            }

            var b = savedBounds.Value;
            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_RESTORE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_NOTOPMOST, b.Left, b.Top, b.Width, b.Height,
                NativeMethods.SWP_SHOWWINDOW);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_TOP, b.Left, b.Top, b.Width, b.Height,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_SHOWWINDOW);
            NativeMethods.SetForegroundWindow(selectedWindow.Handle);
            targetWindowHidden = false;
            if (windowVisibilityToggle != null)
            {
                windowVisibilityToggle.Content = "녹화창 숨기기 (다른 작업 가능)";
                windowVisibilityToggle.IsChecked = false;
            }
            SetStatus("녹화창을 다시 띄웠어요");
        }

        private void SetMonitoring(bool enabled)
        {
            if (audioProcess == null || audioProcess.HasExited) return;
            try
            {
                audioProcess.StandardInput.WriteLine(enabled ? "monitor on" : "monitor off");
                AppendLog("듣기: " + (enabled ? "켬" : "끔"));
            }
            catch
            {
                AppendLog("듣기 상태를 바꾸지 못했어요");
            }
        }

        private double GetSegmentElapsedSeconds(RecordingState recording)
        {
            if (recording == null || recording.SegmentStartedUtc == DateTime.MinValue) return 0;
            return Math.Max(0, (DateTime.UtcNow - recording.SegmentStartedUtc).TotalSeconds);
        }

        private void BeginPauseRange(RecordingState recording)
        {
            if (recording == null || recording.CurrentSegment == null) return;
            RecordingSegment segment = recording.CurrentSegment;
            if (segment.ActivePauseStartSeconds.HasValue) return;
            segment.ActivePauseStartSeconds = GetSegmentElapsedSeconds(recording);
        }

        private void EndPauseRange(RecordingState recording)
        {
            if (recording == null || recording.CurrentSegment == null) return;
            RecordingSegment segment = recording.CurrentSegment;
            if (!segment.ActivePauseStartSeconds.HasValue) return;

            double start = segment.ActivePauseStartSeconds.Value;
            double end = GetSegmentElapsedSeconds(recording);
            if (end > start + 0.05)
            {
                segment.PauseRanges.Add(new PauseRange
                {
                    StartSeconds = start,
                    EndSeconds = end
                });
            }
            segment.ActivePauseStartSeconds = null;
        }

        private void BeginAudioBoostRange(RecordingState recording)
        {
            if (recording == null || recording.CurrentSegment == null) return;
            RecordingSegment segment = recording.CurrentSegment;
            if (segment.ActiveAudioBoostStartSeconds.HasValue) return;

            segment.ActiveAudioBoostStartSeconds = GetSegmentElapsedSeconds(recording);
            segment.BoostAudio = true;
            segment.AudioGain = ReducedPlaybackGain;
            recording.BoostAudio = true;
            recording.AudioGain = ReducedPlaybackGain;
            recording.CurrentAudioGain = ReducedPlaybackGain;
        }

        private void EndAudioBoostRange(RecordingState recording, bool keepReducedMode)
        {
            if (recording == null)
            {
                return;
            }

            RecordingSegment segment = recording.CurrentSegment;
            if (segment != null && segment.ActiveAudioBoostStartSeconds.HasValue)
            {
                double start = segment.ActiveAudioBoostStartSeconds.Value;
                double end = GetSegmentElapsedSeconds(recording);
                if (end > start + 0.05)
                {
                    segment.AudioBoostRanges.Add(new AudioBoostRange
                    {
                        StartSeconds = start,
                        EndSeconds = end
                    });
                }
                segment.ActiveAudioBoostStartSeconds = null;
            }

            if (!keepReducedMode)
            {
                recording.BoostAudio = false;
                recording.AudioGain = 1.0;
                recording.CurrentAudioGain = 1.0;
            }
        }

        private void ApplySilentPlaybackIfRecording()
        {
            if (activeRecording == null || silentPlaybackToggle == null || silentPlaybackToggle.IsChecked != true)
            {
                return;
            }
            if (activeRecording.Target == null) return;

            if (silentPlaybackApplied && audioSessionSnapshots.Count > 0)
            {
                BeginAudioBoostRange(activeRecording);
                return;
            }

            DateTime now = DateTime.UtcNow;
            if ((now - lastSilentPlaybackAttemptUtc).TotalSeconds < 2)
            {
                return;
            }
            lastSilentPlaybackAttemptUtc = now;

            try
            {
                var alreadyLowered = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (AudioSessionSnapshot snapshot in audioSessionSnapshots)
                {
                    if (!string.IsNullOrEmpty(snapshot.SessionInstanceId)) alreadyLowered.Add(snapshot.SessionInstanceId);
                }

                List<AudioSessionSnapshot> snapshots = AudioSessionVolumeController.LowerMatchingSessions(
                    activeRecording.Target.ProcessId,
                    activeRecording.Target.ProcessName,
                    (float)ReducedPlaybackVolume,
                    alreadyLowered);

                audioSessionSnapshots.AddRange(snapshots);
                silentPlaybackApplied = audioSessionSnapshots.Count > 0;
                if (silentPlaybackApplied)
                {
                    BeginAudioBoostRange(activeRecording);
                    if (snapshots.Count > 0) AppendLog("대상 앱 재생 소리를 줄였어요. 저장할 때 소리는 다시 키웁니다.");
                }
                else
                {
                    AppendLog("낮출 수 있는 대상 앱 소리 세션을 찾지 못했어요.");
                }
            }
            catch (Exception ex)
            {
                AppendLog("대상 앱 소리를 낮추지 못했어요: " + ex.Message);
            }
        }

        private void ChangeReducedAudioMode(bool enabled)
        {
            lastSilentPlaybackAttemptUtc = DateTime.MinValue;

            if (activeRecording == null)
            {
                if (!enabled) RestorePlaybackVolume();
                return;
            }

            if (enabled)
            {
                ApplySilentPlaybackIfRecording();
                SetStatus(silentPlaybackApplied ? "소리 줄이기 켬" : "녹화 중");
            }
            else
            {
                EndAudioBoostRange(activeRecording, false);
                RestorePlaybackVolume();
                SetStatus(activeRecording == null ? "준비됨" : "녹화 중");
            }
        }

        private void RestorePlaybackVolume()
        {
            if (!silentPlaybackApplied && audioSessionSnapshots.Count == 0) return;

            try
            {
                AudioSessionVolumeController.Restore(audioSessionSnapshots);
            }
            catch { }
            finally
            {
                audioSessionSnapshots.Clear();
                silentPlaybackApplied = false;
                try
                {
                    Dispatcher.BeginInvoke(new Action(delegate { AppendLog("대상 앱 재생 소리를 원래대로 되돌렸어요."); }));
                }
                catch { }
            }
        }

        private void BrowseFolder()
        {
            string selectedPath;
            try
            {
                if (ModernFolderPicker.TryPickFolder(this, saveFolderBox.Text, out selectedPath))
                {
                    saveFolderBox.Text = selectedPath;
                    SaveSettings(selectedPath);
                }
                return;
            }
            catch (Exception ex)
            {
                AppendLog("새 폴더 선택 창을 열지 못해서 기본 창으로 열어요: " + ex.Message);
            }

            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "녹화 파일을 저장할 폴더를 선택하세요";
                dialog.ShowNewFolderButton = true;
                if (Directory.Exists(saveFolderBox.Text)) dialog.SelectedPath = saveFolderBox.Text;
                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    saveFolderBox.Text = dialog.SelectedPath;
                    SaveSettings(dialog.SelectedPath);
                }
            }
        }

        private void OpenSaveFolder()
        {
            string folder = GetSaveFolder();
            if (folder == null) return;
            try
            {
                Directory.CreateDirectory(folder);
                Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                AppendLog("저장 폴더를 열지 못했어요: " + ex.Message);
            }
        }

        private void OpenSoundMixer()
        {
            Process.Start(new ProcessStartInfo("ms-settings:apps-volume") { UseShellExecute = true });
        }

        private void OpenUserGuide()
        {
            string guidePath = Path.Combine(appDir, "사용설명서.html");
            if (!File.Exists(guidePath))
            {
                guidePath = Path.Combine(supportDir, "사용설명서.html");
            }

            if (!File.Exists(guidePath))
            {
                SetStatus("사용설명서 파일을 찾을 수 없어요");
                AppendLog("사용설명서 파일을 찾을 수 없어요");
                return;
            }

            Process.Start(new ProcessStartInfo(guidePath) { UseShellExecute = true });
        }

        private void OnProcessTimer(object sender, EventArgs e)
        {
            if (isStopping)
            {
                busyTick = (busyTick + 1) % 4;
                SetStatus("녹화 종료 처리 중" + new string('.', busyTick));
                return;
            }

            if (activeRecording == null && (DateTime.UtcNow - lastWindowScanUtc).TotalSeconds >= 2)
            {
                lastWindowScanUtc = DateTime.UtcNow;
                RefreshWindows(true);
            }

            if (isPausing)
            {
                busyTick = (busyTick + 1) % 4;
                string dots = new string('.', busyTick);
                SetStatus("일시정지 중" + dots);
                startButton.Content = "일시정지 중" + dots;
                return;
            }

            if (activeRecording != null && silentPlaybackToggle != null && silentPlaybackToggle.IsChecked == true)
            {
                ApplySilentPlaybackIfRecording();
            }

            if (ffmpegProcess != null && ffmpegProcess.HasExited)
            {
                int code = ffmpegProcess.ExitCode;
                AppendLog("화면 녹화 프로세스 종료: " + code.ToString(CultureInfo.InvariantCulture));
                StopRecording();
            }
        }

        private void OnClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if ((ffmpegProcess != null && !ffmpegProcess.HasExited) || (audioProcess != null && !audioProcess.HasExited))
            {
                var result = System.Windows.MessageBox.Show(this, "녹화가 진행 중입니다. 녹화를 끝내고 닫을까요?", "백그라운드 영상 녹화 프로그램", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result != MessageBoxResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
                if (targetWindowHidden) RestoreTargetWindowFromHidden();
                StopRecording();
            }
            else
            {
                if (targetWindowHidden) RestoreTargetWindowFromHidden();
                RestorePlaybackVolume();
            }
        }

        private bool EnsureTarget()
        {
            if (selectedWindow == null)
            {
                SetStatus("먼저 녹화할 창을 선택해주세요");
                return false;
            }
            if (!NativeMethods.IsWindow(selectedWindow.Handle))
            {
                SetStatus("선택한 창을 찾을 수 없어요");
                RefreshWindows();
                return false;
            }
            return true;
        }

        private string GetSaveFolder()
        {
            string value = (saveFolderBox.Text ?? "").Trim();
            if (value.Length == 0)
            {
                SetStatus("저장 폴더를 선택해주세요");
                return null;
            }
            return Path.GetFullPath(value);
        }

        private double GetAudioSyncDelaySeconds()
        {
            if (audioSyncSlider == null) return DefaultAudioSyncDelaySeconds;
            double value = audioSyncSlider.Value / 1000.0;
            if (double.IsNaN(value) || double.IsInfinity(value)) return DefaultAudioSyncDelaySeconds;
            return Math.Max(MinAudioSyncDelaySeconds, Math.Min(MaxAudioSyncDelaySeconds, value));
        }

        private void UpdateAudioSyncValueText()
        {
            if (audioSyncValueText == null || audioSyncSlider == null) return;
            int ms = (int)Math.Round(audioSyncSlider.Value);
            audioSyncValueText.Text = ms.ToString("+0;-0;0", CultureInfo.InvariantCulture) + " ms";
        }

        private void SaveCurrentSettings()
        {
            if (saveFolderBox == null) return;
            string saveDir = (saveFolderBox.Text ?? "").Trim();
            if (saveDir.Length == 0) return;
            SaveSettings(saveDir);
        }

        private void LoadSettings()
        {
            string fallback = GetDefaultRecordingsPath();
            string selectedPath = fallback;
            double syncDelayMs = DefaultAudioSyncDelaySeconds * 1000.0;
            try
            {
                if (File.Exists(settingsPath))
                {
                    var dict = json.Deserialize<Dictionary<string, object>>(File.ReadAllText(settingsPath, Encoding.UTF8));
                    object value;
                    bool hasCurrentSettings = dict.ContainsKey("SettingsVersion");
                    int settingsVersion = 0;
                    if (dict.TryGetValue("SettingsVersion", out value) && value != null)
                    {
                        int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out settingsVersion);
                    }
                    if (dict.TryGetValue("RecordingsDir", out value) && value != null)
                    {
                        string savedPath = value.ToString();
                        if (!string.IsNullOrWhiteSpace(savedPath) && (hasCurrentSettings || !IsLegacyDefaultPath(savedPath)))
                        {
                            selectedPath = savedPath;
                        }
                    }
                    if (dict.TryGetValue("AudioSyncDelayMs", out value) && value != null)
                    {
                        double parsed;
                        if (double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
                        {
                            syncDelayMs = Math.Max(MinAudioSyncDelaySeconds * 1000.0, Math.Min(MaxAudioSyncDelaySeconds * 1000.0, parsed));
                        }
                    }
                    if (settingsVersion < 3 && Math.Abs(syncDelayMs - 450.0) < 0.5)
                    {
                        syncDelayMs = DefaultAudioSyncDelaySeconds * 1000.0;
                    }
                }
            }
            catch { }

            saveFolderBox.Text = selectedPath;
            if (audioSyncSlider != null)
            {
                audioSyncSlider.Value = syncDelayMs;
                UpdateAudioSyncValueText();
            }
            try { Directory.CreateDirectory(selectedPath); } catch { }
        }

        private string GetDefaultRecordingsPath()
        {
            return Path.Combine(appDir, DefaultRecordingsFolderName);
        }

        private bool IsLegacyDefaultPath(string path)
        {
            string videos = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            return IsSamePath(path, Path.Combine(videos, DefaultRecordingsFolderName))
                || IsSamePath(path, Path.Combine(videos, OldDefaultRecordingsFolderName))
                || IsSamePath(path, Path.Combine(documents, DefaultRecordingsFolderName))
                || IsSamePath(path, Path.Combine(documents, OldDefaultRecordingsFolderName))
                || IsSamePath(path, desktop);
        }

        private static bool IsSamePath(string a, string b)
        {
            try
            {
                return string.Equals(Path.GetFullPath(a).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    Path.GetFullPath(b).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        private void SaveSettings(string saveDir)
        {
            try
            {
                var dict = new Dictionary<string, object>();
                dict["SettingsVersion"] = 3;
                dict["RecordingsDir"] = saveDir;
                dict["AudioSyncDelayMs"] = (int)Math.Round(GetAudioSyncDelaySeconds() * 1000.0);
                File.WriteAllText(settingsPath, json.Serialize(dict), Encoding.UTF8);
            }
            catch { }
        }

        private string PrepareRecordingTempDirectory()
        {
            string root = Path.Combine(supportDir, TempFolderName);
            Directory.CreateDirectory(root);

            try
            {
                foreach (string oldDir in Directory.GetDirectories(root))
                {
                    try
                    {
                        var info = new DirectoryInfo(oldDir);
                        if (info.LastWriteTimeUtc < DateTime.UtcNow.AddDays(-2)) Directory.Delete(oldDir, true);
                    }
                    catch { }
                }
            }
            catch { }

            string tempDir = Path.Combine(root, DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss_", CultureInfo.InvariantCulture) + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(tempDir);
            return tempDir;
        }

        private AudioCaptureMode ResolveAudioCaptureMode()
        {
            if (IsProcessAudioCaptureSupported()) return AudioCaptureMode.Process;
            if (IsSystemAudioCaptureAvailable()) return AudioCaptureMode.SystemLoopback;
            return AudioCaptureMode.None;
        }

        private bool IsAudioCaptureAvailable()
        {
            return HasProcessAudioRecorder() || IsSystemAudioCaptureAvailable();
        }

        private bool HasProcessAudioRecorder()
        {
            if (!string.IsNullOrEmpty(GetProcessAudioHelperPath())) return true;

            string processScript = Path.Combine(supportDir, "process_audio_recorder.py");
            if (File.Exists(processScript)) return true;

            processScript = Path.Combine(appDir, "process_audio_recorder.py");
            return File.Exists(processScript);
        }

        private bool IsSystemAudioCaptureAvailable()
        {
            if (!string.IsNullOrEmpty(GetAudioHelperPath())) return true;

            string script = Path.Combine(supportDir, "loopback_audio_recorder.py");
            if (File.Exists(script)) return true;

            script = Path.Combine(appDir, "loopback_audio_recorder.py");
            return File.Exists(script);
        }

        private bool IsProcessAudioCaptureSupported()
        {
            if (processAudioSupportChecked) return processAudioSupported;
            processAudioSupportChecked = true;
            processAudioSupported = false;

            if (!HasProcessAudioRecorder()) return false;

            string tempPath = Path.Combine(Path.GetTempPath(), "window-back-recorder-process-audio-test-" + Guid.NewGuid().ToString("N") + ".wav");
            Process proc = null;
            try
            {
                proc = StartProcessAudioProbe(tempPath);
                if (proc.WaitForExit(1200))
                {
                    processAudioSupported = false;
                    return false;
                }

                processAudioSupported = true;
                try { proc.StandardInput.WriteLine("q"); } catch { }
                if (!proc.WaitForExit(1500))
                {
                    try { proc.Kill(); } catch { }
                }
                return true;
            }
            catch
            {
                processAudioSupported = false;
                return false;
            }
            finally
            {
                try
                {
                    if (proc != null && !proc.HasExited) proc.Kill();
                }
                catch { }
                TryDelete(tempPath);
            }
        }

        private Process StartProcessAudioProbe(string outputPath)
        {
            string fileName;
            var args = new List<string>();
            string helperPath = GetProcessAudioHelperPath();
            if (!string.IsNullOrEmpty(helperPath))
            {
                fileName = helperPath;
            }
            else
            {
                string script = Path.Combine(supportDir, "process_audio_recorder.py");
                if (!File.Exists(script)) script = Path.Combine(appDir, "process_audio_recorder.py");
                fileName = "python";
                args.Add(script);
            }

            args.Add(outputPath);
            args.Add("--pid");
            args.Add(Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture));

            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = JoinArgs(args),
                WorkingDirectory = supportDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            psi.StandardOutputEncoding = Encoding.UTF8;
            psi.StandardErrorEncoding = Encoding.UTF8;
            ApplyPythonUtf8Environment(psi);

            var process = new Process();
            process.StartInfo = psi;
            if (!process.Start()) throw new InvalidOperationException("process audio probe start failed");
            return process;
        }

        private List<AudioSourceInfo> GetLoopbackSources()
        {
            var results = new List<AudioSourceInfo>();
            string fileName;
            var args = new List<string>();
            string helperPath = GetAudioHelperPath();
            if (!string.IsNullOrEmpty(helperPath))
            {
                fileName = helperPath;
            }
            else
            {
                string script = Path.Combine(supportDir, "loopback_audio_recorder.py");
                if (!File.Exists(script)) script = Path.Combine(appDir, "loopback_audio_recorder.py");
                if (!File.Exists(script)) return results;
                fileName = "python";
                args.Add(script);
            }
            args.Add("--list-json");

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = JoinArgs(args),
                    WorkingDirectory = supportDir,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                psi.StandardOutputEncoding = Encoding.UTF8;
                psi.StandardErrorEncoding = Encoding.UTF8;
                ApplyPythonUtf8Environment(psi);
                var proc = Process.Start(psi);
                string output = proc.StandardOutput.ReadToEnd();
                proc.WaitForExit();
                if (proc.ExitCode != 0 || string.IsNullOrWhiteSpace(output)) return results;

                var items = json.Deserialize<List<Dictionary<string, object>>>(output);
                foreach (var item in items)
                {
                    object name;
                    if (item.TryGetValue("name", out name) && name != null)
                    {
                        object isDefaultValue;
                        bool isDefault = item.TryGetValue("default", out isDefaultValue) && isDefaultValue is bool && (bool)isDefaultValue;
                        results.Add(new AudioSourceInfo(name.ToString(), isDefault));
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("소리 장치를 확인하지 못했어요: " + ex.Message);
            }

            return results;
        }

        private bool TestFfmpegFilter(string filterName)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = GetFfmpegPath(),
                    Arguments = "-hide_banner -h filter=" + filterName,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                var proc = Process.Start(psi);
                string text = proc.StandardOutput.ReadToEnd() + proc.StandardError.ReadToEnd();
                proc.WaitForExit();
                return proc.ExitCode == 0 && text.IndexOf(filterName, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
        }

        private string UniquePath(string directory, string baseName, string extension)
        {
            string path = Path.Combine(directory, baseName + extension);
            if (!File.Exists(path)) return path;
            for (int i = 2; i < 1000; i++)
            {
                path = Path.Combine(directory, string.Format(CultureInfo.InvariantCulture, "{0}-{1:00}{2}", baseName, i, extension));
                if (!File.Exists(path)) return path;
            }
            return Path.Combine(directory, baseName + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + extension);
        }

        private void TryDelete(string path)
        {
            try { if (!string.IsNullOrEmpty(path) && File.Exists(path)) File.Delete(path); } catch { }
        }

        private void SetStatus(string text)
        {
            statusText.Text = text;
        }

        private void AppendLog(string text)
        {
            if (logBox == null) return;
            logLineCount++;
            if (logLineCount > 260)
            {
                logBox.Clear();
                logBox.AppendText(DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + "  이전 로그를 정리했어요." + Environment.NewLine);
                logLineCount = 1;
            }
            logBox.AppendText(DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + "  " + text + Environment.NewLine);
            logBox.ScrollToEnd();
        }

        private void QueueUiLog(string text)
        {
            Dispatcher.BeginInvoke(new Action(delegate { AppendLog(text); }));
        }

        private void QueueProcessLog(string logPrefix, string text)
        {
            if (string.Equals(logPrefix, "[화면] ", StringComparison.Ordinal) && text.IndexOf("frame=", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                DateTime now = DateTime.UtcNow;
                if ((now - lastFfmpegProgressLogUtc).TotalSeconds < 5)
                {
                    return;
                }
                lastFfmpegProgressLogUtc = now;
            }

            Dispatcher.BeginInvoke(new Action(delegate { AppendLog(logPrefix + text); }));
        }

        private static SolidColorBrush Brush(string hex)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        }

        private string GetFfmpegPath()
        {
            string bundled = Path.Combine(supportDir, "bin", "ffmpeg.exe");
            if (File.Exists(bundled)) return bundled;

            bundled = Path.Combine(supportDir, "ffmpeg.exe");
            if (File.Exists(bundled)) return bundled;

            return "ffmpeg";
        }

        private string GetAudioHelperPath()
        {
            string bundled = Path.Combine(supportDir, "bin", "loopback_audio_recorder.exe");
            if (File.Exists(bundled)) return bundled;

            bundled = Path.Combine(supportDir, "loopback_audio_recorder.exe");
            if (File.Exists(bundled)) return bundled;

            return null;
        }

        private string GetProcessAudioHelperPath()
        {
            string bundled = Path.Combine(supportDir, "bin", "process_audio_recorder.exe");
            if (File.Exists(bundled)) return bundled;

            bundled = Path.Combine(supportDir, "process_audio_recorder.exe");
            if (File.Exists(bundled)) return bundled;

            return null;
        }

        private static void ApplyPythonUtf8Environment(ProcessStartInfo psi)
        {
            if (!string.Equals(Path.GetFileNameWithoutExtension(psi.FileName), "python", StringComparison.OrdinalIgnoreCase)) return;
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            psi.EnvironmentVariables["PYTHONUTF8"] = "1";
        }

        private static string JoinArgs(IEnumerable<string> args)
        {
            var builder = new StringBuilder();
            foreach (var arg in args)
            {
                if (builder.Length > 0) builder.Append(' ');
                builder.Append(QuoteArg(arg));
            }
            return builder.ToString();
        }

        private static string QuoteArg(string arg)
        {
            if (arg == null) return "\"\"";
            if (arg.Length == 0) return "\"\"";
            if (!Regex.IsMatch(arg, "[\\s\"]")) return arg;
            return "\"" + arg.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }
    }

    public static class ModernFolderPicker
    {
        private const int Cancelled = unchecked((int)0x800704C7);

        public static bool TryPickFolder(Window owner, string initialPath, out string selectedPath)
        {
            selectedPath = null;
            IFileOpenDialog dialog = null;
            IShellItem initialItem = null;
            IShellItem resultItem = null;
            IntPtr pathPointer = IntPtr.Zero;

            try
            {
                dialog = (IFileOpenDialog)new FileOpenDialogCom();

                uint options;
                Marshal.ThrowExceptionForHR(dialog.GetOptions(out options));
                options |= (uint)(
                    FileOpenOptions.PickFolders |
                    FileOpenOptions.ForceFileSystem |
                    FileOpenOptions.PathMustExist |
                    FileOpenOptions.NoChangeDir);
                Marshal.ThrowExceptionForHR(dialog.SetOptions(options));
                Marshal.ThrowExceptionForHR(dialog.SetTitle("녹화 파일을 저장할 폴더를 선택하세요"));
                Marshal.ThrowExceptionForHR(dialog.SetOkButtonLabel("이 폴더 선택"));

                if (!string.IsNullOrWhiteSpace(initialPath) && Directory.Exists(initialPath))
                {
                    Guid shellItemGuid = typeof(IShellItem).GUID;
                    if (SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref shellItemGuid, out initialItem) == 0 && initialItem != null)
                    {
                        dialog.SetFolder(initialItem);
                    }
                }

                IntPtr ownerHandle = owner == null ? IntPtr.Zero : new WindowInteropHelper(owner).Handle;
                int showResult = dialog.Show(ownerHandle);
                if (showResult == Cancelled) return false;
                Marshal.ThrowExceptionForHR(showResult);

                Marshal.ThrowExceptionForHR(dialog.GetResult(out resultItem));
                Marshal.ThrowExceptionForHR(resultItem.GetDisplayName(ShellItemDisplayName.FileSystemPath, out pathPointer));
                selectedPath = Marshal.PtrToStringUni(pathPointer);
                return !string.IsNullOrWhiteSpace(selectedPath);
            }
            finally
            {
                if (pathPointer != IntPtr.Zero) Marshal.FreeCoTaskMem(pathPointer);
                if (resultItem != null) Marshal.ReleaseComObject(resultItem);
                if (initialItem != null) Marshal.ReleaseComObject(initialItem);
                if (dialog != null) Marshal.ReleaseComObject(dialog);
            }
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string path,
            IntPtr bindContext,
            ref Guid riid,
            out IShellItem shellItem);
    }

    [Flags]
    public enum FileOpenOptions : uint
    {
        PickFolders = 0x00000020,
        ForceFileSystem = 0x00000040,
        NoChangeDir = 0x00000008,
        PathMustExist = 0x00000800
    }

    public enum ShellItemDisplayName : uint
    {
        FileSystemPath = 0x80058000
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DialogFilterSpec
    {
        public string Name;
        public string Spec;
    }

    [ComImport]
    [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    public class FileOpenDialogCom
    {
    }

    [ComImport]
    [Guid("D57C7288-D4AD-4768-BE02-9D969532D960")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        [PreserveSig] int SetFileTypes(uint fileTypesCount, [MarshalAs(UnmanagedType.LPArray)] DialogFilterSpec[] fileTypes);
        [PreserveSig] int SetFileTypeIndex(uint fileTypeIndex);
        [PreserveSig] int GetFileTypeIndex(out uint fileTypeIndex);
        [PreserveSig] int Advise(IntPtr events, out uint cookie);
        [PreserveSig] int Unadvise(uint cookie);
        [PreserveSig] int SetOptions(uint options);
        [PreserveSig] int GetOptions(out uint options);
        [PreserveSig] int SetDefaultFolder(IShellItem shellItem);
        [PreserveSig] int SetFolder(IShellItem shellItem);
        [PreserveSig] int GetFolder(out IShellItem shellItem);
        [PreserveSig] int GetCurrentSelection(out IShellItem shellItem);
        [PreserveSig] int SetFileName([MarshalAs(UnmanagedType.LPWStr)] string fileName);
        [PreserveSig] int GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string fileName);
        [PreserveSig] int SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
        [PreserveSig] int SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string text);
        [PreserveSig] int SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
        [PreserveSig] int GetResult(out IShellItem shellItem);
        [PreserveSig] int AddPlace(IShellItem shellItem, uint fileDialogAddPlace);
        [PreserveSig] int SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string defaultExtension);
        [PreserveSig] int Close(int result);
        [PreserveSig] int SetClientGuid(ref Guid guid);
        [PreserveSig] int ClearClientData();
        [PreserveSig] int SetFilter(IntPtr filter);
        [PreserveSig] int GetResults(out IntPtr shellItemArray);
        [PreserveSig] int GetSelectedItems(out IntPtr shellItemArray);
    }

    [ComImport]
    [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellItem
    {
        [PreserveSig] int BindToHandler(IntPtr bindContext, ref Guid handlerId, ref Guid interfaceId, out IntPtr interfacePointer);
        [PreserveSig] int GetParent(out IShellItem shellItem);
        [PreserveSig] int GetDisplayName(ShellItemDisplayName displayName, out IntPtr name);
        [PreserveSig] int GetAttributes(uint attributeMask, out uint attributes);
        [PreserveSig] int Compare(IShellItem shellItem, uint hint, out int order);
    }

    public sealed class AudioSourceInfo
    {
        public string Name { get; private set; }
        public bool IsDefault { get; private set; }

        public AudioSourceInfo(string name, bool isDefault)
        {
            Name = name;
            IsDefault = isDefault;
        }
    }

    public sealed class AudioOption
    {
        public string Label { get; private set; }
        public string SourceName { get; private set; }

        public AudioOption(string label, string sourceName)
        {
            Label = label;
            SourceName = sourceName;
        }
    }

    public enum AudioCaptureMode
    {
        None,
        Process,
        SystemLoopback
    }

    public sealed class RecordingState
    {
        public string FinalPath;
        public string TempDirectory;
        public readonly List<RecordingSegment> Segments = new List<RecordingSegment>();
        public WindowInfo Target;
        public int Fps;
        public bool DrawCursor;
        public AudioCaptureMode AudioMode;
        public bool HasLoopbackAudio;
        public bool BoostAudio;
        public double AudioGain = 1.0;
        public double CurrentAudioGain = 1.0;
        public double AudioSyncDelaySeconds;
        public RecordingSegment CurrentSegment;
        public DateTime SegmentStartedUtc = DateTime.MinValue;
        public int SegmentIndex;
    }

    public sealed class RecordingSegment
    {
        public string VideoPath;
        public string AudioPath;
        public bool HasLoopbackAudio;
        public bool BoostAudio;
        public double AudioGain = 1.0;
        public double AudioSyncDelaySeconds;
        public int Fps;
        public DateTime VideoStartedUtc = DateTime.MinValue;
        public DateTime AudioStartedUtc = DateTime.MinValue;
        public double TrimStartSeconds;
        public double? ActivePauseStartSeconds;
        public double? ActiveAudioBoostStartSeconds;
        public readonly List<PauseRange> PauseRanges = new List<PauseRange>();
        public readonly List<AudioBoostRange> AudioBoostRanges = new List<AudioBoostRange>();
    }

    public sealed class PauseRange
    {
        public double StartSeconds;
        public double EndSeconds;
    }

    public sealed class AudioBoostRange
    {
        public double StartSeconds;
        public double EndSeconds;
    }

    public sealed class AudioSessionSnapshot
    {
        public string SessionInstanceId;
        public ISimpleAudioVolume Volume;
        public float MasterVolume;
        public bool Muted;
    }

    public static class AudioSessionVolumeController
    {
        private static readonly Guid AudioSessionManager2Guid = new Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F");
        private static readonly Guid EventContext = new Guid("7F4F7F73-9A07-4F8B-8E16-3F3E32DE6C4A");

        public static List<AudioSessionSnapshot> LowerMatchingSessions(int processId, string processName, float volume, ISet<string> alreadyLowered)
        {
            var snapshots = new List<AudioSessionSnapshot>();
            IMMDeviceEnumerator enumerator = null;

            try
            {
                enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
                LowerMatchingSessionsForRole(enumerator, ERole.eConsole, processId, processName, volume, alreadyLowered, snapshots);
                LowerMatchingSessionsForRole(enumerator, ERole.eMultimedia, processId, processName, volume, alreadyLowered, snapshots);
                LowerMatchingSessionsForRole(enumerator, ERole.eCommunications, processId, processName, volume, alreadyLowered, snapshots);
            }
            finally
            {
                if (enumerator != null) Marshal.ReleaseComObject(enumerator);
            }

            return snapshots;
        }

        private static void LowerMatchingSessionsForRole(IMMDeviceEnumerator enumerator, ERole role, int processId, string processName, float volume, ISet<string> alreadyLowered, List<AudioSessionSnapshot> snapshots)
        {
            IMMDevice device = null;
            IAudioSessionManager2 manager = null;
            IAudioSessionEnumerator sessions = null;

            try
            {
                if (enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, role, out device) != 0 || device == null) return;

                object managerObject;
                Guid iid = AudioSessionManager2Guid;
                Marshal.ThrowExceptionForHR(device.Activate(ref iid, CLSCTX.CLSCTX_ALL, IntPtr.Zero, out managerObject));
                manager = (IAudioSessionManager2)managerObject;
                Marshal.ThrowExceptionForHR(manager.GetSessionEnumerator(out sessions));

                int count;
                Marshal.ThrowExceptionForHR(sessions.GetCount(out count));
                string normalizedProcessName = (processName ?? "").Trim();
                Guid context = EventContext;

                for (int i = 0; i < count; i++)
                {
                    IAudioSessionControl control = null;
                    try
                    {
                        Marshal.ThrowExceptionForHR(sessions.GetSession(i, out control));
                        var control2 = control as IAudioSessionControl2;
                        var simpleVolume = control as ISimpleAudioVolume;
                        if (control2 == null || simpleVolume == null) continue;

                        int sessionPid;
                        Marshal.ThrowExceptionForHR(control2.GetProcessId(out sessionPid));
                        if (!MatchesProcess(sessionPid, processId, normalizedProcessName)) continue;

                        string sessionInstanceId = "";
                        try { control2.GetSessionInstanceIdentifier(out sessionInstanceId); } catch { }
                        if (!string.IsNullOrEmpty(sessionInstanceId) && alreadyLowered != null && alreadyLowered.Contains(sessionInstanceId)) continue;

                        float originalVolume;
                        bool originalMute;
                        Marshal.ThrowExceptionForHR(simpleVolume.GetMasterVolume(out originalVolume));
                        Marshal.ThrowExceptionForHR(simpleVolume.GetMute(out originalMute));

                        snapshots.Add(new AudioSessionSnapshot
                        {
                            SessionInstanceId = sessionInstanceId,
                            Volume = simpleVolume,
                            MasterVolume = originalVolume,
                            Muted = originalMute
                        });
                        if (!string.IsNullOrEmpty(sessionInstanceId) && alreadyLowered != null) alreadyLowered.Add(sessionInstanceId);

                        Marshal.ThrowExceptionForHR(simpleVolume.SetMute(false, ref context));
                        Marshal.ThrowExceptionForHR(simpleVolume.SetMasterVolume(volume, ref context));
                    }
                    catch
                    {
                        if (control != null) Marshal.ReleaseComObject(control);
                    }
                }
            }
            catch { }
            finally
            {
                if (sessions != null) Marshal.ReleaseComObject(sessions);
                if (manager != null) Marshal.ReleaseComObject(manager);
                if (device != null) Marshal.ReleaseComObject(device);
            }
        }

        public static void Restore(List<AudioSessionSnapshot> snapshots)
        {
            Guid context = EventContext;
            foreach (AudioSessionSnapshot snapshot in snapshots)
            {
                try
                {
                    if (snapshot.Volume != null)
                    {
                        snapshot.Volume.SetMasterVolume(snapshot.MasterVolume, ref context);
                        snapshot.Volume.SetMute(snapshot.Muted, ref context);
                    }
                }
                catch { }
                finally
                {
                    try
                    {
                        if (snapshot.Volume != null) Marshal.ReleaseComObject(snapshot.Volume);
                    }
                    catch { }
                }
            }
        }

        private static bool MatchesProcess(int sessionPid, int selectedPid, string selectedProcessName)
        {
            if (sessionPid == selectedPid) return true;
            if (string.IsNullOrWhiteSpace(selectedProcessName)) return false;

            try
            {
                using (Process process = Process.GetProcessById(sessionPid))
                {
                    return string.Equals(process.ProcessName, selectedProcessName, StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                return false;
            }
        }
    }

    public enum EDataFlow
    {
        eRender = 0,
        eCapture = 1,
        eAll = 2
    }

    public enum ERole
    {
        eConsole = 0,
        eMultimedia = 1,
        eCommunications = 2
    }

    [Flags]
    public enum CLSCTX : uint
    {
        CLSCTX_INPROC_SERVER = 0x1,
        CLSCTX_INPROC_HANDLER = 0x2,
        CLSCTX_LOCAL_SERVER = 0x4,
        CLSCTX_REMOTE_SERVER = 0x10,
        CLSCTX_ALL = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumerator
    {
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        [PreserveSig] int EnumAudioEndpoints(EDataFlow dataFlow, uint dwStateMask, out IntPtr ppDevices);
        [PreserveSig] int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
        [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
        [PreserveSig] int RegisterEndpointNotificationCallback(IntPtr pClient);
        [PreserveSig] int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        [PreserveSig] int Activate(ref Guid iid, CLSCTX dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        [PreserveSig] int OpenPropertyStore(uint stgmAccess, out IntPtr ppProperties);
        [PreserveSig] int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
        [PreserveSig] int GetState(out uint pdwState);
    }

    [ComImport]
    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionManager2
    {
        [PreserveSig] int GetAudioSessionControl(IntPtr AudioSessionGuid, uint StreamFlags, out IAudioSessionControl SessionControl);
        [PreserveSig] int GetSimpleAudioVolume(IntPtr AudioSessionGuid, uint StreamFlags, out ISimpleAudioVolume AudioVolume);
        [PreserveSig] int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);
        [PreserveSig] int RegisterSessionNotification(IntPtr SessionNotification);
        [PreserveSig] int UnregisterSessionNotification(IntPtr SessionNotification);
        [PreserveSig] int RegisterDuckNotification(string sessionID, IntPtr duckNotification);
        [PreserveSig] int UnregisterDuckNotification(IntPtr duckNotification);
    }

    [ComImport]
    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionEnumerator
    {
        [PreserveSig] int GetCount(out int SessionCount);
        [PreserveSig] int GetSession(int SessionCount, out IAudioSessionControl Session);
    }

    [ComImport]
    [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionControl
    {
        [PreserveSig] int GetState(out int pRetVal);
        [PreserveSig] int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        [PreserveSig] int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
        [PreserveSig] int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        [PreserveSig] int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
        [PreserveSig] int GetGroupingParam(out Guid pRetVal);
        [PreserveSig] int SetGroupingParam(ref Guid Override, ref Guid EventContext);
        [PreserveSig] int RegisterAudioSessionNotification(IntPtr NewNotifications);
        [PreserveSig] int UnregisterAudioSessionNotification(IntPtr NewNotifications);
    }

    [ComImport]
    [Guid("BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioSessionControl2
    {
        [PreserveSig] int GetState(out int pRetVal);
        [PreserveSig] int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        [PreserveSig] int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
        [PreserveSig] int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        [PreserveSig] int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
        [PreserveSig] int GetGroupingParam(out Guid pRetVal);
        [PreserveSig] int SetGroupingParam(ref Guid Override, ref Guid EventContext);
        [PreserveSig] int RegisterAudioSessionNotification(IntPtr NewNotifications);
        [PreserveSig] int UnregisterAudioSessionNotification(IntPtr NewNotifications);
        [PreserveSig] int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        [PreserveSig] int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        [PreserveSig] int GetProcessId(out int pRetVal);
        [PreserveSig] int IsSystemSoundsSession();
        [PreserveSig] int SetDuckingPreference([MarshalAs(UnmanagedType.Bool)] bool optOut);
    }

    [ComImport]
    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISimpleAudioVolume
    {
        [PreserveSig] int SetMasterVolume(float fLevel, ref Guid EventContext);
        [PreserveSig] int GetMasterVolume(out float pfLevel);
        [PreserveSig] int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid EventContext);
        [PreserveSig] int GetMute([MarshalAs(UnmanagedType.Bool)] out bool pbMute);
    }

    public sealed class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public int ProcessId { get; set; }
        public string Title { get; set; }
        public string ProcessName { get; set; }
        public IntRect Bounds { get; set; }

        public string SizeText
        {
            get { return Bounds.Width.ToString(CultureInfo.InvariantCulture) + "x" + Bounds.Height.ToString(CultureInfo.InvariantCulture); }
        }

        public string HandleText
        {
            get { return "0x" + Handle.ToInt64().ToString("X", CultureInfo.InvariantCulture); }
        }
    }

    public struct IntRect
    {
        public int Left;
        public int Top;
        public int Width;
        public int Height;
    }

    public static class NativeWindows
    {
        public static List<WindowInfo> GetOpenWindows()
        {
            var windows = new List<WindowInfo>();
            NativeMethods.EnumWindows(delegate(IntPtr hwnd, IntPtr lParam)
            {
                if (!NativeMethods.IsWindowVisible(hwnd)) return true;
                string title = GetWindowTitle(hwnd);
                if (string.IsNullOrWhiteSpace(title)) return true;

                IntRect bounds = GetBounds(hwnd);
                if (bounds.Width < 80 || bounds.Height < 60) return true;

                string processName = "";
                uint pid;
                NativeMethods.GetWindowThreadProcessId(hwnd, out pid);
                try { processName = Process.GetProcessById((int)pid).ProcessName; } catch { }

                windows.Add(new WindowInfo
                {
                    Handle = hwnd,
                    ProcessId = (int)pid,
                    Title = title,
                    ProcessName = processName,
                    Bounds = bounds
                });
                return true;
            }, IntPtr.Zero);

            windows.Sort(delegate(WindowInfo a, WindowInfo b)
            {
                int c = string.Compare(a.ProcessName, b.ProcessName, StringComparison.OrdinalIgnoreCase);
                if (c != 0) return c;
                return string.Compare(a.Title, b.Title, StringComparison.OrdinalIgnoreCase);
            });

            return windows;
        }

        public static IntRect GetBounds(IntPtr hwnd)
        {
            NativeMethods.RECT rect;
            NativeMethods.GetWindowRect(hwnd, out rect);
            return new IntRect
            {
                Left = rect.Left,
                Top = rect.Top,
                Width = rect.Right - rect.Left,
                Height = rect.Bottom - rect.Top
            };
        }

        private static string GetWindowTitle(IntPtr hwnd)
        {
            int length = NativeMethods.GetWindowTextLength(hwnd);
            if (length <= 0) return "";
            var builder = new StringBuilder(length + 1);
            NativeMethods.GetWindowText(hwnd, builder, builder.Capacity);
            return builder.ToString().Trim();
        }
    }

    public static class NativeMethods
    {
        public const int SW_RESTORE = 9;
        public const int SW_SHOWNOACTIVATE = 4;
        public const uint SWP_NOSIZE = 0x0001;
        public const uint SWP_NOMOVE = 0x0002;
        public const uint SWP_NOACTIVATE = 0x0010;
        public const uint SWP_SHOWWINDOW = 0x0040;
        public static readonly IntPtr HWND_TOP = IntPtr.Zero;
        public static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowTextLength(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hwnd, int command);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}

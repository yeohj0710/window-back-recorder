using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
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
        private ToggleButton cursorToggle;
        private Slider fpsSlider;

        private Process ffmpegProcess;
        private Process audioProcess;
        private RecordingState activeRecording;
        private WindowInfo selectedWindow;
        private IntRect? savedBounds;
        private bool gfxCaptureAvailable;
        private DateTime lastWindowScanUtc = DateTime.MinValue;
        private string lastWindowSnapshot = "";

        private const string SupportFolderName = "프로그램 구성 파일";
        private const string NoAudioLabel = "소리 없이 녹화";
        private const string DeveloperLabel = "developed by yeohj0710";
        private const string DefaultRecordingsFolderName = "녹화 완료된 동영상";
        private const string OldDefaultRecordingsFolderName = "녹화 완료 영상";

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
            Width = Math.Min(880, Math.Max(760, workArea.Width - 160));
            Height = Math.Min(740, Math.Max(640, workArea.Height - 80));
            MinWidth = 760;
            MinHeight = 640;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = workArea.Left + 40;
            Top = workArea.Top + 40;
            Background = Brush("#0b0f14");
            FontFamily = new FontFamily("Segoe UI");
            Foreground = Brush("#e8edf2");

            BuildUi();
            LoadSettings();
            RefreshCapabilities();
            RefreshWindows();
            RefreshAudioSources();

            processTimer.Interval = TimeSpan.FromMilliseconds(800);
            processTimer.Tick += OnProcessTimer;
            processTimer.Start();

            Closing += OnClosing;
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
                Background = Brush("#10161d"),
                BorderBrush = Brush("#1c2733"),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(22, 18, 22, 16)
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
                FontSize = 23,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brush("#f5f7fa"),
                VerticalAlignment = VerticalAlignment.Center
            };
            titleRow.Children.Add(title);
            titleRow.Children.Add(CreateDeveloperBadge());

            statusText = new TextBlock
            {
                Text = "준비됨",
                Margin = new Thickness(0, 6, 0, 0),
                FontSize = 13,
                Foreground = Brush("#8da2b8")
            };
            titleStack.Children.Add(statusText);

            var headerButtons = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(headerButtons, 1);
            headerGrid.Children.Add(headerButtons);

            headerButtons.Children.Add(CreateButton("사용법", OpenUserGuide, "#16212b", "#263546"));
            headerButtons.Children.Add(Spacer(8, 1));
            headerButtons.Children.Add(CreateButton("소리 확인", OpenSoundMixer, "#16212b", "#263546"));
            headerButtons.Children.Add(Spacer(8, 1));
            headerButtons.Children.Add(CreateButton("새로고침", delegate { RefreshWindows(); RefreshAudioSources(); }, "#16212b", "#263546"));

            var main = new Grid { Margin = new Thickness(12) };
            main.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            main.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(270) });
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
                Foreground = Brush("#f2f5f8"),
                VerticalAlignment = VerticalAlignment.Center
            });

            targetText = new TextBlock
            {
                Text = "창을 선택해주세요",
                FontSize = 12,
                Foreground = Brush("#7f91a5"),
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
                Background = Brush("#0d131a"),
                Foreground = Brush("#dfe7ef"),
                BorderBrush = Brush("#202c38"),
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

            var controls = new StackPanel();
            rightPanel.Child = controls;

            controls.Children.Add(SectionLabel("녹화 설정"));
            controls.Children.Add(FormLabel("저장 폴더"));

            var folderRow = new Grid { Margin = new Thickness(0, 4, 0, 14) };
            folderRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            folderRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(54) });
            controls.Children.Add(folderRow);

            saveFolderBox = CreateTextBox();
            Grid.SetColumn(saveFolderBox, 0);
            folderRow.Children.Add(saveFolderBox);

            var browse = CreateSmallButton("변경", BrowseFolder);
            browse.Margin = new Thickness(6, 0, 0, 0);
            Grid.SetColumn(browse, 1);
            folderRow.Children.Add(browse);

            controls.Children.Add(FormLabel("소리"));
            audioStatusText = MetaText("선택한 앱 소리만 자동 녹음");
            audioStatusText.Margin = new Thickness(0, 4, 0, 14);
            audioStatusText.Foreground = Brush("#cfe2f5");
            controls.Children.Add(audioStatusText);

            controls.Children.Add(FormLabel("화면 부드러움"));
            var fpsRow = new Grid { Margin = new Thickness(0, 4, 0, 16) };
            fpsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            fpsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(52) });
            controls.Children.Add(fpsRow);

            fpsSlider = new Slider { Minimum = 5, Maximum = 60, Value = 30, TickFrequency = 5, IsSnapToTickEnabled = true };
            Grid.SetColumn(fpsSlider, 0);
            fpsRow.Children.Add(fpsSlider);
            var fpsValue = new TextBlock { Foreground = Brush("#9fb1c4"), VerticalAlignment = VerticalAlignment.Center, TextAlignment = TextAlignment.Right };
            fpsValue.SetBinding(TextBlock.TextProperty, new Binding("Value") { Source = fpsSlider, StringFormat = "{0:0} fps" });
            Grid.SetColumn(fpsValue, 1);
            fpsRow.Children.Add(fpsValue);

            cursorToggle = CreateToggle("마우스 커서도 녹화", true);
            controls.Children.Add(cursorToggle);

            listenToggle = CreateToggle("내 스피커로 듣기", false);
            listenToggle.Visibility = Visibility.Collapsed;

            var buttonGrid = new Grid { Margin = new Thickness(0, 18, 0, 14) };
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            controls.Children.Add(buttonGrid);

            startButton = CreateButton("녹화 시작", StartRecording, "#1e6bff", "#3d82ff");
            startButton.Height = 44;
            Grid.SetColumn(startButton, 0);
            buttonGrid.Children.Add(startButton);

            stopButton = CreateButton("녹화 종료", StopRecordingFromUi, "#3a2027", "#60313a");
            stopButton.Height = 44;
            stopButton.IsEnabled = false;
            Grid.SetColumn(stopButton, 2);
            buttonGrid.Children.Add(stopButton);

            controls.Children.Add(SectionLabel("창 보기"));
            var viewGrid = new UniformGrid { Columns = 2, Margin = new Thickness(0, 8, 0, 12) };
            controls.Children.Add(viewGrid);
            viewGrid.Children.Add(CreateSmallButton("뒤로 보내기", SendTargetBack));
            viewGrid.Children.Add(CreateSmallButton("앞으로 보기", BringTargetFront));
            viewGrid.Children.Add(CreateSmallButton("작게 두기", CompactTarget));
            viewGrid.Children.Add(CreateSmallButton("크기 복원", RestoreTarget));

            controls.Children.Add(SectionLabel("상태"));
            engineText = MetaText("화면 캡처: 확인 중");
            outputText = MetaText("저장 파일: -");
            controls.Children.Add(engineText);
            controls.Children.Add(outputText);

            var footer = new Border
            {
                Background = Brush("#080c10"),
                BorderBrush = Brush("#1c2733"),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(20, 12, 20, 12)
            };
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            logBox = new TextBox
            {
                Height = 82,
                Background = Brush("#080c10"),
                Foreground = Brush("#8fa3b8"),
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
                Background = Brush("#10161d"),
                BorderBrush = Brush("#202b37"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(16)
            };
        }

        private Border CreateDeveloperBadge()
        {
            return new Border
            {
                Margin = new Thickness(14, 3, 0, 0),
                Padding = new Thickness(10, 4, 10, 5),
                Background = Brush("#11272b"),
                BorderBrush = Brush("#2a7b72"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = DeveloperLabel,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brush("#87eadc"),
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
        }

        private TextBlock SectionLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = Brush("#f2f5f8"),
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 10)
            };
        }

        private TextBlock FormLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = Brush("#93a6ba"),
                FontSize = 12,
                Margin = new Thickness(0, 0, 0, 2)
            };
        }

        private TextBlock MetaText(string text)
        {
            return new TextBlock
            {
                Text = text,
                Foreground = Brush("#9fb1c4"),
                FontSize = 12,
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
                Background = Brush("#0d131a"),
                Foreground = Brush("#e8edf2"),
                BorderBrush = Brush("#263545"),
                Padding = new Thickness(10, 6, 10, 6)
            };
        }

        private ToggleButton CreateToggle(string text, bool isChecked)
        {
            var toggle = new ToggleButton
            {
                Content = text,
                IsChecked = isChecked,
                Margin = new Thickness(0, 0, 0, 8),
                Height = 34,
                Foreground = Brush("#dce5ee"),
                Background = Brush("#111a23"),
                BorderBrush = Brush("#263545"),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            return toggle;
        }

        private Button CreateSmallButton(string text, Action action)
        {
            var button = CreateButton(text, action, "#16212b", "#263546");
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
                Background = Brush(background),
                Foreground = Brush("#f5f7fa"),
                BorderBrush = Brush(hover),
                BorderThickness = new Thickness(1),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            button.Click += delegate { action(); };
            return button;
        }

        private FrameworkElement Spacer(double width, double height)
        {
            return new Border { Width = width, Height = height };
        }

        private GridViewColumn CreateColumn(string header, string path, double width)
        {
            var textBlock = new FrameworkElementFactory(typeof(TextBlock));
            textBlock.SetBinding(TextBlock.TextProperty, new Binding(path));
            textBlock.SetValue(TextBlock.ForegroundProperty, Brush("#e8f5ff"));
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
            itemStyle.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#0d131a")));
            itemStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#dfe7ef")));
            itemStyle.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#0d131a")));
            itemStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));

            var selectedTrigger = new Trigger { Property = ListViewItem.IsSelectedProperty, Value = true };
            selectedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#173a52")));
            selectedTrigger.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#ffffff")));
            selectedTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#45b8e8")));
            itemStyle.Triggers.Add(selectedTrigger);

            var mouseOverTrigger = new Trigger { Property = ListViewItem.IsMouseOverProperty, Value = true };
            mouseOverTrigger.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#132434")));
            mouseOverTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#2f5f7d")));
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
            bool available = IsAudioCaptureAvailable();
            if (audioStatusText != null)
            {
                audioStatusText.Text = available ? "선택한 앱 소리만 자동 녹음" : "소리 녹음 준비 안 됨";
                audioStatusText.Foreground = available ? Brush("#cfe2f5") : Brush("#ffb7c3");
            }

            AppendLog(available ? "소리 녹음: 선택한 앱 우선 사용" : "소리 녹음 준비 안 됨");
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

        private void StartRecording()
        {
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

            bool hasLoopback = IsAudioCaptureAvailable();
            if (!hasLoopback)
            {
                SetStatus("소리 녹음 준비 안 됨");
                AppendLog("소리 녹음 파일을 찾지 못했어요. 프로그램 구성 파일을 그대로 둔 상태에서 다시 실행해주세요.");
                return;
            }

            string baseName = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss", CultureInfo.InvariantCulture);
            string finalPath = UniquePath(saveDir, baseName, ".mp4");
            string videoPath = hasLoopback ? Path.Combine(saveDir, Path.GetFileNameWithoutExtension(finalPath) + ".video.mkv") : finalPath;
            string audioPath = hasLoopback ? Path.Combine(saveDir, Path.GetFileNameWithoutExtension(finalPath) + ".audio.wav") : null;

            savedBounds = selectedWindow.Bounds;

            try
            {
                ffmpegProcess = StartFfmpegCapture(selectedWindow, videoPath, (int)fpsSlider.Value, cursorToggle.IsChecked == true, !hasLoopback);
                if (hasLoopback)
                {
                    audioProcess = StartTargetAudio(audioPath, selectedWindow.ProcessId);
                }

                activeRecording = new RecordingState
                {
                    FinalPath = finalPath,
                    VideoPath = videoPath,
                    AudioPath = audioPath,
                    HasLoopbackAudio = hasLoopback
                };

                startButton.IsEnabled = false;
                stopButton.IsEnabled = true;
                saveFolderBox.IsEnabled = false;
                outputText.Text = "저장 파일: " + finalPath;
                SetStatus("녹화 중");
                AppendLog("녹화 시작: " + finalPath);
            }
            catch (Exception ex)
            {
                AppendLog("녹화를 시작하지 못했어요: " + ex.Message);
                StopProcessesOnly();
                SetStatus("시작 실패");
            }
        }

        private Process StartFfmpegCapture(WindowInfo target, string outputPath, int fps, bool drawCursor, bool finalOutput)
        {
            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-y");

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

            return StartProcess(GetFfmpegPath(), args, "[화면] ");
        }

        private Process StartLoopbackAudio(string audioPath, string sourceName, bool monitorOn)
        {
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
            var process = StartProcess(fileName, args, "[소리] ");
            if (process.WaitForExit(900))
            {
                throw new InvalidOperationException("소리 녹음이 바로 종료됐어요. Windows 기본 출력 장치와 녹화할 앱의 음소거 상태를 확인해주세요.");
            }
            return process;
        }

        private Process StartTargetAudio(string audioPath, int processId)
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
                if (!File.Exists(script))
                {
                    AppendLog("선택한 앱 소리 녹음 helper가 없어 PC 전체 소리 녹음으로 대신합니다.");
                    return StartLoopbackAudio(audioPath, null, false);
                }
                fileName = "python";
                args.Add(script);
            }

            args.Add(audioPath);
            args.Add("--pid");
            args.Add(processId.ToString(CultureInfo.InvariantCulture));

            var process = StartProcess(fileName, args, "[소리] ");
            if (process.WaitForExit(900))
            {
                throw new InvalidOperationException("선택한 앱 소리 녹음이 바로 종료됐어요. 앱에서 실제로 소리가 나오는지 확인해주세요.");
            }
            return process;
        }

        private Process StartProcess(string fileName, List<string> args, string logPrefix)
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
                if (!string.IsNullOrEmpty(e.Data)) Dispatcher.BeginInvoke(new Action(delegate { AppendLog(logPrefix + e.Data); }));
            };
            process.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs e)
            {
                if (!string.IsNullOrEmpty(e.Data)) Dispatcher.BeginInvoke(new Action(delegate { AppendLog(logPrefix + e.Data); }));
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
            var recording = activeRecording;
            StopProcessesOnly();

            startButton.IsEnabled = true;
            stopButton.IsEnabled = false;
            saveFolderBox.IsEnabled = true;

            if (recording != null && recording.HasLoopbackAudio)
            {
                MuxRecording(recording);
            }

            activeRecording = null;
            SetStatus("준비됨");
        }

        private void StopProcessesOnly()
        {
            StopProcessNicely(ffmpegProcess);
            StopProcessNicely(audioProcess);
            ffmpegProcess = null;
            audioProcess = null;
        }

        private void StopProcessNicely(Process process)
        {
            if (process == null) return;
            try
            {
                if (!process.HasExited)
                {
                    try { process.StandardInput.WriteLine("q"); } catch { }
                    if (!process.WaitForExit(6000))
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

        private void MuxRecording(RecordingState recording)
        {
            try
            {
                AppendLog("영상과 소리를 합치는 중...");
                var args = new List<string>();
                args.Add("-hide_banner");
                args.Add("-y");
                args.Add("-i");
                args.Add(recording.VideoPath);
                args.Add("-i");
                args.Add(recording.AudioPath);
                args.Add("-map");
                args.Add("0:v:0");
                args.Add("-map");
                args.Add("1:a:0");
                args.Add("-c:v");
                args.Add("copy");
                args.Add("-c:a");
                args.Add("aac");
                args.Add("-b:a");
                args.Add("160k");
                args.Add("-shortest");
                args.Add(recording.FinalPath);

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
                string err = proc.StandardError.ReadToEnd();
                proc.WaitForExit();

                if (proc.ExitCode == 0)
                {
                    TryDelete(recording.VideoPath);
                    TryDelete(recording.AudioPath);
                    outputText.Text = "저장 파일: " + recording.FinalPath;
                    AppendLog("저장 완료: " + recording.FinalPath);
                }
                else
                {
                    AppendLog("파일을 합치지 못했어요: " + err);
                }
            }
            catch (Exception ex)
            {
                AppendLog("파일을 합치지 못했어요: " + ex.Message);
            }
        }

        private void SendTargetBack()
        {
            if (!EnsureTarget()) return;
            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_SHOWNOACTIVATE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_BOTTOM, 0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
            SetStatus("창을 뒤로 보냈어요");
        }

        private void BringTargetFront()
        {
            if (!EnsureTarget()) return;
            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_RESTORE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_TOP, 0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_SHOWWINDOW);
            NativeMethods.SetForegroundWindow(selectedWindow.Handle);
            SetStatus("창을 앞으로 가져왔어요");
        }

        private void CompactTarget()
        {
            if (!EnsureTarget()) return;
            savedBounds = NativeWindows.GetBounds(selectedWindow.Handle);
            var area = WinForms.Screen.PrimaryScreen.WorkingArea;
            int width = Math.Max(360, Math.Min(560, area.Width / 3));
            int height = Math.Max(230, Math.Min(360, area.Height / 3));
            int x = area.Right - width - 20;
            int y = area.Bottom - height - 20;
            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_SHOWNOACTIVATE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_BOTTOM, x, y, width, height,
                NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
            SetStatus("창을 작게 두었어요");
        }

        private void RestoreTarget()
        {
            if (!EnsureTarget()) return;
            if (!savedBounds.HasValue)
            {
                SetStatus("저장된 창 크기가 없어요");
                return;
            }
            var b = savedBounds.Value;
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_BOTTOM, b.Left, b.Top, b.Width, b.Height,
                NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
            SetStatus("창 크기를 복원했어요");
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

        private void BrowseFolder()
        {
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
            if (activeRecording == null && (DateTime.UtcNow - lastWindowScanUtc).TotalSeconds >= 2)
            {
                lastWindowScanUtc = DateTime.UtcNow;
                RefreshWindows(true);
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
                StopRecording();
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

        private void LoadSettings()
        {
            string videos = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
            if (string.IsNullOrWhiteSpace(videos))
            {
                videos = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            string fallback = GetDefaultRecordingsPath(videos);
            string previousDefault = Path.Combine(videos, DefaultRecordingsFolderName);
            string oldDefault = Path.Combine(videos, OldDefaultRecordingsFolderName);
            string selectedPath = fallback;
            try
            {
                if (File.Exists(settingsPath))
                {
                    var dict = json.Deserialize<Dictionary<string, object>>(File.ReadAllText(settingsPath, Encoding.UTF8));
                    object value;
                    if (dict.TryGetValue("RecordingsDir", out value) && value != null)
                    {
                        string savedPath = value.ToString();
                        if (!string.IsNullOrWhiteSpace(savedPath) && !IsSamePath(savedPath, previousDefault) && !IsSamePath(savedPath, oldDefault))
                        {
                            selectedPath = savedPath;
                        }
                    }
                }
            }
            catch { }

            saveFolderBox.Text = selectedPath;
            try { Directory.CreateDirectory(selectedPath); } catch { }
        }

        private string GetDefaultRecordingsPath(string videosPath)
        {
            if (!string.Equals(supportDir, appDir, StringComparison.OrdinalIgnoreCase))
            {
                return Path.Combine(appDir, DefaultRecordingsFolderName);
            }

            return Path.Combine(videosPath, DefaultRecordingsFolderName);
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
                dict["RecordingsDir"] = saveDir;
                File.WriteAllText(settingsPath, json.Serialize(dict), Encoding.UTF8);
            }
            catch { }
        }

        private bool IsAudioCaptureAvailable()
        {
            if (!string.IsNullOrEmpty(GetProcessAudioHelperPath())) return true;

            string processScript = Path.Combine(supportDir, "process_audio_recorder.py");
            if (File.Exists(processScript)) return true;

            processScript = Path.Combine(appDir, "process_audio_recorder.py");
            if (File.Exists(processScript)) return true;

            if (!string.IsNullOrEmpty(GetAudioHelperPath())) return true;

            string script = Path.Combine(supportDir, "loopback_audio_recorder.py");
            if (File.Exists(script)) return true;

            script = Path.Combine(appDir, "loopback_audio_recorder.py");
            return File.Exists(script);
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
            logBox.AppendText(DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + "  " + text + Environment.NewLine);
            logBox.ScrollToEnd();
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

    public sealed class RecordingState
    {
        public string FinalPath;
        public string VideoPath;
        public string AudioPath;
        public bool HasLoopbackAudio;
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

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
        private readonly string settingsPath;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly DispatcherTimer processTimer = new DispatcherTimer();

        private ListView windowList;
        private ComboBox audioCombo;
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

        private const string NoAudioLabel = "No audio";

        public MainWindow()
        {
            appDir = AppDomain.CurrentDomain.BaseDirectory;
            settingsPath = Path.Combine(appDir, "settings.json");

            Title = "Window Back Recorder";
            Width = 1180;
            Height = 760;
            MinWidth = 980;
            MinHeight = 640;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
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

            var title = new TextBlock
            {
                Text = "Window Back Recorder",
                FontSize = 25,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brush("#f5f7fa")
            };
            titleStack.Children.Add(title);

            statusText = new TextBlock
            {
                Text = "Ready",
                Margin = new Thickness(0, 6, 0, 0),
                FontSize = 13,
                Foreground = Brush("#8da2b8")
            };
            titleStack.Children.Add(statusText);

            var headerButtons = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(headerButtons, 1);
            headerGrid.Children.Add(headerButtons);

            headerButtons.Children.Add(CreateButton("Sound mixer", OpenSoundMixer, "#16212b", "#263546"));
            headerButtons.Children.Add(Spacer(10, 1));
            headerButtons.Children.Add(CreateButton("Refresh", delegate { RefreshWindows(); RefreshAudioSources(); }, "#16212b", "#263546"));

            var main = new Grid { Margin = new Thickness(20) };
            main.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            main.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(360) });
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
            DockPanel.SetDock(listHeader, Dock.Top);
            leftDock.Children.Add(listHeader);

            listHeader.Children.Add(new TextBlock
            {
                Text = "Target window",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brush("#f2f5f8"),
                VerticalAlignment = VerticalAlignment.Center
            });

            targetText = new TextBlock
            {
                Text = "No target selected",
                FontSize = 12,
                Foreground = Brush("#7f91a5"),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(targetText, 1);
            listHeader.Children.Add(targetText);

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
            gridView.Columns.Add(CreateColumn("Title", "Title", 520));
            gridView.Columns.Add(CreateColumn("App", "ProcessName", 150));
            gridView.Columns.Add(CreateColumn("Size", "SizeText", 110));
            gridView.Columns.Add(CreateColumn("HWND", "HandleText", 110));
            windowList.View = gridView;

            var rightPanel = CreatePanel();
            rightPanel.Margin = new Thickness(18, 0, 0, 0);
            Grid.SetColumn(rightPanel, 1);
            main.Children.Add(rightPanel);

            var controls = new StackPanel();
            rightPanel.Child = controls;

            controls.Children.Add(SectionLabel("Recording"));
            controls.Children.Add(FormLabel("Save folder"));

            var folderRow = new Grid { Margin = new Thickness(0, 4, 0, 14) };
            folderRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            folderRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            controls.Children.Add(folderRow);

            saveFolderBox = CreateTextBox();
            Grid.SetColumn(saveFolderBox, 0);
            folderRow.Children.Add(saveFolderBox);

            var browse = CreateSmallButton("Browse", BrowseFolder);
            browse.Margin = new Thickness(8, 0, 0, 0);
            Grid.SetColumn(browse, 1);
            folderRow.Children.Add(browse);

            controls.Children.Add(FormLabel("Audio source"));
            audioCombo = new ComboBox
            {
                Margin = new Thickness(0, 4, 0, 14),
                MinHeight = 34,
                Background = Brush("#0d131a"),
                Foreground = Brush("#e8edf2"),
                BorderBrush = Brush("#263545")
            };
            controls.Children.Add(audioCombo);

            controls.Children.Add(FormLabel("Frame rate"));
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

            cursorToggle = CreateToggle("Record cursor", true);
            controls.Children.Add(cursorToggle);

            listenToggle = CreateToggle("Listen to captured audio", false);
            listenToggle.Checked += delegate { SetMonitoring(true); };
            listenToggle.Unchecked += delegate { SetMonitoring(false); };
            controls.Children.Add(listenToggle);

            var buttonGrid = new Grid { Margin = new Thickness(0, 18, 0, 14) };
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(12) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            controls.Children.Add(buttonGrid);

            startButton = CreateButton("Start", StartRecording, "#1e6bff", "#3d82ff");
            startButton.Height = 44;
            Grid.SetColumn(startButton, 0);
            buttonGrid.Children.Add(startButton);

            stopButton = CreateButton("Stop", StopRecordingFromUi, "#3a2027", "#60313a");
            stopButton.Height = 44;
            stopButton.IsEnabled = false;
            Grid.SetColumn(stopButton, 2);
            buttonGrid.Children.Add(stopButton);

            controls.Children.Add(SectionLabel("View"));
            var viewGrid = new UniformGrid { Columns = 2, Margin = new Thickness(0, 8, 0, 12) };
            controls.Children.Add(viewGrid);
            viewGrid.Children.Add(CreateSmallButton("View off", SendTargetBack));
            viewGrid.Children.Add(CreateSmallButton("View on", BringTargetFront));
            viewGrid.Children.Add(CreateSmallButton("Compact", CompactTarget));
            viewGrid.Children.Add(CreateSmallButton("Restore", RestoreTarget));

            controls.Children.Add(SectionLabel("Status"));
            engineText = MetaText("Engine: checking");
            outputText = MetaText("Output: -");
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
                Padding = new Thickness(14, 7, 14, 7),
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
            return new GridViewColumn
            {
                Header = header,
                Width = width,
                DisplayMemberBinding = new Binding(path)
            };
        }

        private void StyleListView(ListView list)
        {
            var itemStyle = new Style(typeof(ListViewItem));
            itemStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 7, 8, 7)));
            itemStyle.Setters.Add(new Setter(Control.MarginProperty, new Thickness(0, 0, 0, 2)));
            itemStyle.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#0d131a")));
            itemStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brush("#dfe7ef")));
            list.ItemContainerStyle = itemStyle;
        }

        private void RefreshCapabilities()
        {
            gfxCaptureAvailable = TestFfmpegFilter("gfxcapture");
            engineText.Text = gfxCaptureAvailable ? "Engine: gfxcapture" : "Engine: gdigrab fallback";
        }

        private void RefreshWindows()
        {
            var selectedHandle = selectedWindow == null ? IntPtr.Zero : selectedWindow.Handle;
            var windows = NativeWindows.GetOpenWindows();
            windowList.ItemsSource = windows;
            selectedWindow = null;
            targetText.Text = "No target selected";

            foreach (WindowInfo info in windows)
            {
                if (info.Handle == selectedHandle)
                {
                    windowList.SelectedItem = info;
                    break;
                }
            }

            AppendLog("Windows refreshed: " + windows.Count.ToString(CultureInfo.InvariantCulture));
        }

        private void RefreshAudioSources()
        {
            var sources = new List<AudioOption>();
            sources.Add(new AudioOption(NoAudioLabel, null));

            foreach (var name in GetLoopbackSources())
            {
                sources.Add(new AudioOption("Loopback: " + name, name));
            }

            audioCombo.ItemsSource = sources;
            audioCombo.DisplayMemberPath = "Label";
            audioCombo.SelectedIndex = sources.Count > 1 ? 1 : 0;
            AppendLog("Loopback sources: " + Math.Max(0, sources.Count - 1).ToString(CultureInfo.InvariantCulture));
        }

        private void OnWindowSelected(object sender, SelectionChangedEventArgs e)
        {
            selectedWindow = windowList.SelectedItem as WindowInfo;
            if (selectedWindow == null)
            {
                targetText.Text = "No target selected";
                return;
            }
            targetText.Text = selectedWindow.ProcessName + "  " + selectedWindow.SizeText;
        }

        private void StartRecording()
        {
            if (selectedWindow == null)
            {
                SetStatus("Select a target window first");
                return;
            }

            if (!NativeMethods.IsWindow(selectedWindow.Handle))
            {
                SetStatus("Target window is gone");
                RefreshWindows();
                return;
            }

            string saveDir = GetSaveFolder();
            if (saveDir == null) return;

            if (!Directory.Exists(saveDir)) Directory.CreateDirectory(saveDir);
            SaveSettings(saveDir);

            var audio = audioCombo.SelectedItem as AudioOption;
            bool hasLoopback = audio != null && !string.IsNullOrEmpty(audio.SourceName);

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
                    audioProcess = StartLoopbackAudio(audioPath, audio.SourceName, listenToggle.IsChecked == true);
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
                outputText.Text = "Output: " + finalPath;
                SetStatus("Recording");
                AppendLog("Recording started: " + finalPath);
            }
            catch (Exception ex)
            {
                AppendLog("Start failed: " + ex.Message);
                StopProcessesOnly();
                SetStatus("Start failed");
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

            return StartProcess("ffmpeg", args, "[video] ");
        }

        private Process StartLoopbackAudio(string audioPath, string sourceName, bool monitorOn)
        {
            var args = new List<string>();
            args.Add(Path.Combine(appDir, "loopback_audio_recorder.py"));
            args.Add(audioPath);
            args.Add("--source");
            args.Add(sourceName);
            if (monitorOn) args.Add("--monitor-on");
            return StartProcess("python", args, "[audio] ");
        }

        private Process StartProcess(string fileName, List<string> args, string logPrefix)
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = JoinArgs(args),
                WorkingDirectory = appDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            psi.StandardOutputEncoding = Encoding.UTF8;
            psi.StandardErrorEncoding = Encoding.UTF8;

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

            if (!process.Start()) throw new InvalidOperationException("Could not start " + fileName);
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
            SetStatus("Ready");
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
                AppendLog("Muxing final MP4...");
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
                    FileName = "ffmpeg",
                    Arguments = JoinArgs(args),
                    WorkingDirectory = appDir,
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
                    outputText.Text = "Output: " + recording.FinalPath;
                    AppendLog("Saved: " + recording.FinalPath);
                }
                else
                {
                    AppendLog("Mux failed: " + err);
                }
            }
            catch (Exception ex)
            {
                AppendLog("Mux failed: " + ex.Message);
            }
        }

        private void SendTargetBack()
        {
            if (!EnsureTarget()) return;
            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_SHOWNOACTIVATE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_BOTTOM, 0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
            SetStatus("View off");
        }

        private void BringTargetFront()
        {
            if (!EnsureTarget()) return;
            NativeMethods.ShowWindow(selectedWindow.Handle, NativeMethods.SW_RESTORE);
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_TOP, 0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_SHOWWINDOW);
            NativeMethods.SetForegroundWindow(selectedWindow.Handle);
            SetStatus("View on");
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
            SetStatus("Compact view");
        }

        private void RestoreTarget()
        {
            if (!EnsureTarget()) return;
            if (!savedBounds.HasValue)
            {
                SetStatus("No saved size");
                return;
            }
            var b = savedBounds.Value;
            NativeMethods.SetWindowPos(selectedWindow.Handle, NativeMethods.HWND_BOTTOM, b.Left, b.Top, b.Width, b.Height,
                NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
            SetStatus("Restored");
        }

        private void SetMonitoring(bool enabled)
        {
            if (audioProcess == null || audioProcess.HasExited) return;
            try
            {
                audioProcess.StandardInput.WriteLine(enabled ? "monitor on" : "monitor off");
                AppendLog("Listen: " + (enabled ? "on" : "off"));
            }
            catch
            {
                AppendLog("Could not change listening state");
            }
        }

        private void BrowseFolder()
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "Choose recording folder";
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

        private void OnProcessTimer(object sender, EventArgs e)
        {
            if (ffmpegProcess != null && ffmpegProcess.HasExited)
            {
                int code = ffmpegProcess.ExitCode;
                AppendLog("Video process exited: " + code.ToString(CultureInfo.InvariantCulture));
                StopRecording();
            }
        }

        private void OnClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if ((ffmpegProcess != null && !ffmpegProcess.HasExited) || (audioProcess != null && !audioProcess.HasExited))
            {
                var result = System.Windows.MessageBox.Show(this, "A recording is running. Stop and close?", "Window Back Recorder", MessageBoxButton.YesNo, MessageBoxImage.Warning);
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
                SetStatus("Select a target first");
                return false;
            }
            if (!NativeMethods.IsWindow(selectedWindow.Handle))
            {
                SetStatus("Target window is gone");
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
                SetStatus("Choose a save folder");
                return null;
            }
            return Path.GetFullPath(value);
        }

        private void LoadSettings()
        {
            string fallback = Path.Combine(appDir, "recordings");
            saveFolderBox.Text = fallback;
            try
            {
                if (File.Exists(settingsPath))
                {
                    var dict = json.Deserialize<Dictionary<string, object>>(File.ReadAllText(settingsPath, Encoding.UTF8));
                    object value;
                    if (dict.TryGetValue("RecordingsDir", out value) && value != null)
                    {
                        saveFolderBox.Text = value.ToString();
                    }
                }
            }
            catch { }
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

        private List<string> GetLoopbackSources()
        {
            var results = new List<string>();
            string script = Path.Combine(appDir, "loopback_audio_recorder.py");
            if (!File.Exists(script)) return results;

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = JoinArgs(new List<string> { script, "--list-json" }),
                    WorkingDirectory = appDir,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                psi.StandardOutputEncoding = Encoding.UTF8;
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
                        results.Add(name.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("Audio probe failed: " + ex.Message);
            }

            return results;
        }

        private bool TestFfmpegFilter(string filterName)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "ffmpeg",
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
        public IntPtr Handle;
        public string Title;
        public string ProcessName;
        public IntRect Bounds;

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

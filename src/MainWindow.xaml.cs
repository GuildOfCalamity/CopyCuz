using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Con = System.Diagnostics.Debug;

namespace CopyCuz
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// <remarks>
    //  We must include INotifyPropertyChanged as an inheritance for the class, or OnPropertyChanged() calls will not work.
    /// </remarks>
    public partial class MainWindow : Window, INotifyPropertyChanged 
    {
        #region [Properties]
        public event PropertyChangedEventHandler PropertyChanged;
        System.Windows.Forms.NotifyIcon notifyIcon;
        const int MENU1 = 0; 
        const int MENU2 = 1; 
        const int MENU3 = 2;

        string textColor = "#DDDDDD";
        public string TextColor
        {
            get => textColor;
            set
            {
                if (textColor != value)
                {
                    textColor = value;
                    OnPropertyChanged();
                }
            }
        }

        string fuseColor = "#DDDDDD";
        public string FuseColor
        {
            get => fuseColor;
            set
            {
                if (fuseColor != value)
                {
                    fuseColor = value;
                    OnPropertyChanged();
                }
            }
        }

        string versionBrush = "#85151520";
        public string VersionBrush
        {
            get => versionBrush;
            set
            {
                if (versionBrush != value)
                {
                    versionBrush = value;
                    OnPropertyChanged();
                }
            }
        }

        bool dataPulse = false;
        public bool DataPulse
        {
            get => dataPulse;
            set
            {
                if (dataPulse != value)
                {
                    dataPulse = value;
                    OnPropertyChanged();
                }
            }
        }

        bool isBusy = false;
        public bool IsBusy
        {
            get => isBusy;
            set
            {
                if (isBusy != value)
                {
                    isBusy = value;
                    OnPropertyChanged();
                }
            }
        }

        bool isAnimated = false;
        public bool IsAnimated
        {
            get => isAnimated;
            set
            {
                if (isAnimated != value)
                {
                    isAnimated = value;
                    OnPropertyChanged();
                }
            }
        }

        string status = "";
        public string Status
        {
            get => status;
            set
            {
                if (status != value)
                {
                    status = value;
                    OnPropertyChanged();
                }
            }
        }

        double textWidth = 400;
        public double TextWidth
        {
            get => textWidth;
            set
            {
                if (textWidth != value)
                {
                    textWidth = value;
                    OnPropertyChanged();
                }
            }
        }

        string header1 = "Copy Group 1";
        public string Header1
        {
            get => header1;
            set
            {
                if (header1 != value)
                {
                    header1 = value;
                    OnPropertyChanged();
                    btnCopy1.Content = header1.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{header1}" : $"Copy {header1}";
                }
            }
        }

        string header2 = "Copy Group 2";
        public string Header2
        {
            get => header2;
            set
            {
                if (header2 != value)
                {
                    header2 = value;
                    OnPropertyChanged();
                    btnCopy2.Content = header2.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{header2}" : $"Copy {header2}";
                }
            }
        }

        string header3 = "Copy Group 3";
        public string Header3
        {
            get => header3;
            set
            {
                if (header3 != value)
                {
                    header3 = value;
                    OnPropertyChanged();
                    btnCopy3.Content = header3.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{header3}" : $"Copy {header3}";
                }
            }
        }

        string header4 = "Copy Group 4";
        public string Header4
        {
            get => header4;
            set
            {
                if (header4 != value)
                {
                    header4 = value;
                    OnPropertyChanged();
                    btnCopy4.Content = header4.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{header4}" : $"Copy {header4}";
                }
            }
        }

        bool is1Expanded = false;
        public bool Is1Expanded
        {
            get => is1Expanded; 
            set
            {
                is1Expanded = value;
                OnPropertyChanged(nameof(Is1Expanded));
                if (is1Expanded)
                {
                    Is2Expanded = Is3Expanded = Is4Expanded = false;
                    OnPropertyChanged(nameof(Is2Expanded));
                    OnPropertyChanged(nameof(Is3Expanded));
                    OnPropertyChanged(nameof(Is4Expanded));
                }
            }
        }

        bool is2Expanded = false;
        public bool Is2Expanded
        {
            get => is2Expanded;
            set
            {
                is2Expanded = value;
                OnPropertyChanged(nameof(Is2Expanded));
                if (is2Expanded)
                {
                    Is1Expanded = Is3Expanded = Is4Expanded = false;
                    OnPropertyChanged(nameof(Is1Expanded));
                    OnPropertyChanged(nameof(Is3Expanded));
                    OnPropertyChanged(nameof(Is4Expanded));
                }
            }
        }

        bool is3Expanded = false;
        public bool Is3Expanded
        {
            get => is3Expanded;
            set
            {
                is3Expanded = value;
                OnPropertyChanged(nameof(Is3Expanded));
                if (is3Expanded)
                {
                    Is1Expanded = Is2Expanded = Is4Expanded = false;
                    OnPropertyChanged(nameof(Is1Expanded));
                    OnPropertyChanged(nameof(Is2Expanded));
                    OnPropertyChanged(nameof(Is4Expanded));
                }
            }
        }

        bool is4Expanded = false;
        public bool Is4Expanded
        {
            get => is4Expanded;
            set
            {
                is4Expanded = value;
                OnPropertyChanged(nameof(Is4Expanded));
                if (is4Expanded)
                {
                    Is1Expanded = Is2Expanded = Is3Expanded = false;
                    OnPropertyChanged(nameof(Is1Expanded));
                    OnPropertyChanged(nameof(Is2Expanded));
                    OnPropertyChanged(nameof(Is4Expanded));
                }
            }
        }
        #endregion

        #region [Configs]
        bool _firstRun = true;
        bool _topMost = false;
        bool _runOnStartup = false;
        double _windowLeft = 0;
        double _windowTop = 0;
        string _folderFrom1 = string.Empty;
        string _folderFrom2 = string.Empty;
        string _folderFrom3 = string.Empty;
        string _folderFrom4 = string.Empty;
        string _folderTo1 = string.Empty;
        string _folderTo2 = string.Empty;
        string _folderTo3 = string.Empty;
        string _folderTo4 = string.Empty;
        string _lastCopy = string.Empty;
        HashSet<string> _exList = new HashSet<string>();
        #endregion

        public MainWindow()
        {
            InitializeComponent();

            this.ShowInTaskbar = false; // hide in task bar
            this.SizeToContent = SizeToContent.Height; // grow vertically for expander controls

            #region [NotifyIcon]
            try
            {
                notifyIcon = new System.Windows.Forms.NotifyIcon();
                notifyIcon.Icon = CopyCuz.Properties.Resources.new_logo;
                notifyIcon.Click += NotifyIcon_Click;
                notifyIcon.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
                notifyIcon.ContextMenuStrip.Items.Add("Repeat last copy", CopyCuz.Properties.Resources.repeat, OnRepeatClicked); // MENU1
                notifyIcon.ContextMenuStrip.Items.Add("Stay on top", CopyCuz.Properties.Resources.x, OnTopmostClicked);       // MENU2
                notifyIcon.ContextMenuStrip.Items.Add("Exit application", CopyCuz.Properties.Resources.exit, OnExitClicked);  // MENU3
                notifyIcon.Visible = true;
                notifyIcon.Text = "CopyCuz";
            }
            catch (Exception ex)
            {
                App.ShowDialog($"NotifyIcon creation error: {ex.Message}", "Warning", assetName: "error.png");
            }
            #endregion

            this.DataContext = this; // ⇦ very important for INotifyPropertyChanged!

            Status = $"v{App.GetCurrentAssemblyVersion()}";

            var exFile = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Exclusions.txt");
            if (File.Exists(exFile))
            {
                _exList = Extensions.ReadLines(exFile);
                Status = $"{(_exList.Count == 1 ? "1 exclusion" : $"{_exList.Count} exclusions")} loaded";
            }
        }

        #region [Events]
        void ComboWindow_Activated(object sender, EventArgs e) => IsAnimated = true;

        void ComboWindow_Deactivated(object sender, EventArgs e) => IsAnimated = false;

        void ComboWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (e.NewSize.IsInvalidOrZero())
                return;

            //Status = $"New size: {e.NewSize.Width:N0},{e.NewSize.Height:N0}";
            // Add in some margin
            bkgnd.Width = e.NewSize.Width - 8;
            bkgnd.Height = e.NewSize.Height - 8;
        }

        async void Copy_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                var btn = sender as System.Windows.Controls.Button;
                _lastCopy = btn.Name;
                //var tag = btn.Tag;
                switch (_lastCopy)
                {
                    case "btnCopy1":
                        if (string.IsNullOrEmpty(tbFrom1.Text) || string.IsNullOrEmpty(tbTo1.Text))
                        {
                            Status = "Fix the from/to path";
                            return;
                        }
                        // Disable UI elements while the operation is running
                        spinner1.Visibility = Visibility.Visible;
                        btnCopy1.IsEnabled = tbFrom1.IsEnabled = tbTo1.IsEnabled = false;
                        btnCopy1.Content = "";
                        // Await the asynchronous operation
                        await CopyFilesAsync(tbFrom1.Text, tbTo1.Text);
                        // Re-enable UI elements and update status once complete (back on the UI thread)
                        await Dispatcher.InvokeAsync(async () =>
                        {
                            await Task.Delay(200);
                            spinner1.Visibility = Visibility.Hidden;
                            btnCopy1.IsEnabled = tbFrom1.IsEnabled = tbTo1.IsEnabled = true;
                            btnCopy1.Content = Header1.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{Header1}" : $"Copy {Header1}";
                        }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    case "btnCopy2":
                        if (string.IsNullOrEmpty(tbFrom2.Text) || string.IsNullOrEmpty(tbTo2.Text))
                        {
                            Status = "Fix the from/to path";
                            return;
                        }
                        // Disable UI elements while the operation is running
                        spinner2.Visibility = Visibility.Visible;
                        btnCopy2.IsEnabled = tbFrom2.IsEnabled = tbTo2.IsEnabled = false;
                        btnCopy2.Content = "";
                        // Await the asynchronous operation
                        await CopyFilesAsync(tbFrom2.Text, tbTo2.Text);
                        // Re-enable UI elements and update status once complete (back on the UI thread)
                        await Dispatcher.InvokeAsync(async () =>
                        {
                            await Task.Delay(200);
                            spinner2.Visibility = Visibility.Hidden;
                            btnCopy2.IsEnabled = tbFrom2.IsEnabled = tbTo2.IsEnabled = true;
                            btnCopy2.Content = Header2.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{Header2}" : $"Copy {Header2}";
                        }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    case "btnCopy3":
                        if (string.IsNullOrEmpty(tbFrom3.Text) || string.IsNullOrEmpty(tbTo3.Text))
                        {
                            Status = "Fix the from/to path";
                            return;
                        }
                        // Disable UI elements while the operation is running
                        spinner3.Visibility = Visibility.Visible;
                        btnCopy3.IsEnabled = tbFrom3.IsEnabled = tbTo3.IsEnabled = false;
                        btnCopy3.Content = "";
                        // Await the asynchronous operation
                        await CopyFilesAsync(tbFrom3.Text, tbTo3.Text);
                        // Re-enable UI elements and update status once complete (back on the UI thread)
                        await Dispatcher.InvokeAsync(async () =>
                        {
                            await Task.Delay(200);
                            spinner3.Visibility = Visibility.Hidden;
                            btnCopy3.IsEnabled = tbFrom3.IsEnabled = tbTo3.IsEnabled = true;
                            btnCopy3.Content = Header3.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{Header3}" : $"Copy {Header3}";
                        }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    case "btnCopy4":
                        if (string.IsNullOrEmpty(tbFrom4.Text) || string.IsNullOrEmpty(tbTo4.Text))
                        {
                            Status = "Fix the from/to path";
                            return;
                        }
                        // Disable UI elements while the operation is running
                        spinner4.Visibility = Visibility.Visible;
                        btnCopy4.IsEnabled = tbFrom4.IsEnabled = tbTo4.IsEnabled = false;
                        btnCopy4.Content = "";
                        // Await the asynchronous operation
                        await CopyFilesAsync(tbFrom4.Text, tbTo4.Text);
                        // Re-enable UI elements and update status once complete (back on the UI thread)
                        await Dispatcher.InvokeAsync(async () =>
                        {
                            await Task.Delay(200);
                            spinner4.Visibility = Visibility.Hidden;
                            btnCopy4.IsEnabled = tbFrom4.IsEnabled = tbTo4.IsEnabled = true;
                            btnCopy4.Content = Header4.StartsWith("copy", StringComparison.OrdinalIgnoreCase) ? $"{Header4}" : $"Copy {Header4}";
                        }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    default:
                        
                        break;
                }
            }
            else
            {
                Status = "Copy sender is null!";
            }
        }

        void NotifyIcon_Click(object sender, EventArgs e)
        {
            if (e is System.Windows.Forms.MouseEventArgs)
            {
                if ((e as System.Windows.Forms.MouseEventArgs).Button == System.Windows.Forms.MouseButtons.Left)
                {
                    //Application.Current.MainWindow.WindowState = WindowState.Minimized;
                    this.Activate();
                    this.Focus();
                }
                else if ((e as System.Windows.Forms.MouseEventArgs).Button == System.Windows.Forms.MouseButtons.Right)
                {
                    //Application.Current.MainWindow.WindowState = WindowState.Normal;
                    //Application.Current.Shutdown();
                }
            }
        }

        async void OnRepeatClicked(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastCopy))
            {
                switch (_lastCopy)
                {
                    case "btnCopy1": Is1Expanded = true;
                        await Dispatcher.InvokeAsync(async () => { Copy_Click(btnCopy1, new RoutedEventArgs()); }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    case "btnCopy2": Is2Expanded = true;
                        await Dispatcher.InvokeAsync(async () => { Copy_Click(btnCopy2, new RoutedEventArgs()); }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    case "btnCopy3": Is3Expanded = true;
                        await Dispatcher.InvokeAsync(async () => { Copy_Click(btnCopy3, new RoutedEventArgs()); }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    case "btnCopy4": Is4Expanded = true;
                        await Dispatcher.InvokeAsync(async () => { Copy_Click(btnCopy4, new RoutedEventArgs()); }, System.Windows.Threading.DispatcherPriority.Background);
                        break;
                    default: Status = $"Undefined switch '{_lastCopy}'";
                        break;
                }

            }
            else
            {
                Status = "No previous copy action to repeat";
            }
        }

        void OnTopmostClicked(object sender, EventArgs e)
        {
            if (this.Topmost)
            {
                this.Topmost = false;
                // swap icon image
                notifyIcon.ContextMenuStrip.Items[MENU2].Image = CopyCuz.Properties.Resources.x;
            }
            else
            {
                this.Topmost = true;
                // swap icon image
                notifyIcon.ContextMenuStrip.Items[MENU2].Image = CopyCuz.Properties.Resources.check;
            }

            ConfigManager.Set("WindowOnTop", value: this.Topmost);
        }

        void OnExitClicked(object sender, EventArgs e) => System.Windows.Application.Current.Shutdown();

        void ComboWindow_Loaded(object sender, RoutedEventArgs e)
        {
            spinner1.Visibility = spinner2.Visibility = spinner3.Visibility = spinner4.Visibility = Visibility.Hidden;

            #region [Load Configs]
            _topMost = ConfigManager.Get("WindowOnTop", defaultValue: false);
            _firstRun = ConfigManager.Get("FirstRun", defaultValue: true);
            _runOnStartup = ConfigManager.Get("RunOnStartup", defaultValue: false);
            _windowLeft = ConfigManager.Get("WindowLeft", defaultValue: 250D);
            _windowTop = ConfigManager.Get("WindowTop", defaultValue: 200D);
            _folderFrom1 = ConfigManager.Get("FolderFrom1", defaultValue: string.Empty);
            _folderTo1 = ConfigManager.Get("FolderTo1", defaultValue: string.Empty);
            _folderFrom2 = ConfigManager.Get("FolderFrom2", defaultValue: string.Empty);
            _folderTo2 = ConfigManager.Get("FolderTo2", defaultValue: string.Empty);
            _folderFrom3 = ConfigManager.Get("FolderFrom3", defaultValue: string.Empty);
            _folderTo3 = ConfigManager.Get("FolderTo3", defaultValue: string.Empty);
            _folderFrom4 = ConfigManager.Get("FolderFrom4", defaultValue: string.Empty);
            _folderTo4 = ConfigManager.Get("FolderTo4", defaultValue: string.Empty);
            Header1 = ConfigManager.Get("Header1", defaultValue: "Copy Group 1");
            Header2 = ConfigManager.Get("Header2", defaultValue: "Copy Group 2");
            Header3 = ConfigManager.Get("Header3", defaultValue: "Copy Group 3");
            Header4 = ConfigManager.Get("Header4", defaultValue: "Copy Group 4");
            if (!string.IsNullOrEmpty(_folderFrom1)) { tbFrom1.Text = _folderFrom1; }
            if (!string.IsNullOrEmpty(_folderTo1)) { tbTo1.Text = _folderTo1; }
            if (!string.IsNullOrEmpty(_folderFrom2)) { tbFrom2.Text = _folderFrom2; }
            if (!string.IsNullOrEmpty(_folderTo2)) { tbTo2.Text = _folderTo2; }
            if (!string.IsNullOrEmpty(_folderFrom3)) { tbFrom3.Text = _folderFrom3; }
            if (!string.IsNullOrEmpty(_folderTo3)) { tbTo3.Text = _folderTo3; }
            if (!string.IsNullOrEmpty(_folderFrom4)) { tbFrom4.Text = _folderFrom4; }
            if (!string.IsNullOrEmpty(_folderTo4)) { tbTo4.Text = _folderTo4; }

            // Check if position is on any screen
            this.RestorePosition(_windowLeft, _windowTop);

            if (notifyIcon != null && _firstRun)
                notifyIcon.ShowBalloonTip(3000, "Notice", $"I'll be in the tray if you need me.", System.Windows.Forms.ToolTipIcon.Info);

            // Are we required in startup?
            RegistryStartupHelper.SetStartup("CopyCuz", _runOnStartup);
            #endregion

            //if (Debugger.IsAttached)
            //{
            //    tbFrom1.Text = @"D:\Cache\Folder1";
            //    tbTo1.Text = @"D:\Cache\Folder2";
            //}
        }

        void ComboWindow_Closing(object sender, CancelEventArgs e)
        {
            #region [Save Configs]
            ConfigManager.Set("FolderFrom1", value: tbFrom1.Text);
            ConfigManager.Set("FolderTo1", value: tbTo1.Text);
            ConfigManager.Set("FolderFrom2", value: tbFrom2.Text);
            ConfigManager.Set("FolderTo2", value: tbTo2.Text);
            ConfigManager.Set("FolderFrom3", value: tbFrom3.Text);
            ConfigManager.Set("FolderTo3", value: tbTo3.Text);
            ConfigManager.Set("FolderFrom4", value: tbFrom4.Text);
            ConfigManager.Set("FolderTo4", value: tbTo4.Text);
            ConfigManager.Set("Header1", value: Header1);
            ConfigManager.Set("Header2", value: Header2);
            ConfigManager.Set("Header3", value: Header3);
            ConfigManager.Set("Header4", value: Header4);
            ConfigManager.Set("LastUse", value: DateTime.Now);
            ConfigManager.Set("FirstRun", value: false);
            ConfigManager.Set("WindowOnTop", value: _topMost);
            ConfigManager.Set("WindowLeft", value: this.Left.IsInvalid() ? 250D : this.Left);
            ConfigManager.Set("WindowTop", value: this.Top.IsInvalid() ? 200D : this.Top);
            #endregion

            if (notifyIcon != null)
            {
                notifyIcon.Visible = false;
                notifyIcon.Icon.Dispose();
                notifyIcon.Dispose();
                notifyIcon = null;
            }
        }

        void SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
             tbTo1.Text = "value copied to clipboard";
             DataPulse = true; //trigger our fade in/out style
             System.Windows.Clipboard.SetText(tbTo1.Text);
             DataPulse = false; //trigger our fade in/out style
        }

        void cmbCustomColors_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.ComboBox)
            {
                var tmp = sender as System.Windows.Controls.ComboBox;
                if (tmp.SelectedItem != null)
                {
                    //Application.Current.Dispatcher.Invoke(() => ComboSelected = tmp.SelectedItem.ToString());
                    //Con.WriteLine($"[CustomSelectionChanged]: {tmp.SelectedItem}");
                    //Con.WriteLine($"[CustomBindingProperty]: {ComboSelected}");
                }
            }
        }

        public void OnPropertyChanged([CallerMemberName] string propertyName = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        void ComboWindow_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                Cursor = System.Windows.Input.Cursors.Hand;
                DragMove();
            }
            Cursor = System.Windows.Input.Cursors.Arrow;
        }

        void ComboWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                App.Current.Shutdown();
        }

        void Expander_ExpandedOrCollapsed(object sender, RoutedEventArgs e)
        {
            // The SizeToContent="Height" property handles the automatic resizing.
            // We only need to check if the new size makes the window go off-screen.

            // This code must run after the layout system has updated the window size.
            // Dispatcher.InvokeAsync with a low priority ensures the layout pass completes first.
            Dispatcher.InvokeAsync(() => { AdjustWindowPosition(); }, System.Windows.Threading.DispatcherPriority.Loaded);
        }
        #endregion

        #region [Helper Methods]
        /// <summary>
        /// Checks the given <paramref name="exclusions"/> to see if the <paramref name="fileName"/> matches any of the patterns.
        /// </summary>
        static bool IsExcluded(string fileName, HashSet<string> exclusions)
        {
            if (string.IsNullOrWhiteSpace(fileName)) 
                return false;
            
            if (exclusions == null || exclusions.Count == 0) 
                return false;

            foreach (var pattern in exclusions)
            {
                if (fileName.MatchesPattern(pattern))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Checks the given <paramref name="exclusions"/> to see if the <paramref name="fileName"/> matches any of the patterns.
        /// </summary>
        static bool IsExcluded(string fileName, List<string> exclusions)
        {
            if (string.IsNullOrWhiteSpace(fileName)) 
                return false;

            if (exclusions == null || exclusions.Count == 0) 
                return false;

            foreach (var pattern in exclusions)
            {
                if (fileName.MatchesPattern(pattern))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Copies files from the <paramref name="source"/> to the <paramref name="destination"/> using <see cref="System.IO.File.Copy(string, string)"/>.
        /// </summary>
        async Task CopyFilesAsync(string source, string destination)
        {
            int estCount = 0;
            int cpyCount = 0;
            try
            {
                Status = "Starting copy…";

                // Wrap the entire synchronous operation in Task.Run
                // Delegates all file I/O work to background thread pool thread.
                await Task.Run(() =>
                {
                    try
                    {
                        if (!source.EndsWith("\\")) { source += "\\"; }

                        // These operations now run safely in the background
                        foreach (var p in Directory.GetDirectories(source, "*", SearchOption.AllDirectories))
                        {
                            Directory.CreateDirectory(System.IO.Path.Combine(destination, p.Substring(source.Length)));
                        }

                        // A little inefficient, but we need to know how much to copy vs what was copied.
                        IEnumerable<string> filesToCount = Directory.GetFiles(source, "*", SearchOption.AllDirectories).Where(fileName => !IsExcluded(System.IO.Path.GetFileName(fileName), _exList));
                        foreach (string fn in filesToCount)
                        {
                            estCount++;
                        }

                        // Now copy the file listing to the destination.
                        IEnumerable<string> filesToProcess = Directory.GetFiles(source, "*", SearchOption.AllDirectories).Where(fileName => !IsExcluded(System.IO.Path.GetFileName(fileName), _exList));
                        foreach (string fn in filesToProcess)
                        {
                            try
                            {
                                System.IO.File.Copy(fn, System.IO.Path.Combine(destination, fn.Substring(source.Length)), true);
                                cpyCount++;
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[WARNING] CopyFilesAsync: {ex.Message}");
                                $"[WARNING] File copy failed ⇒ {ex.Message}".WriteToLog();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[WARNING] CopyFilesAsync(Task.Run): {ex.Message}");
                        App.ShowDialog($" File copy failed! \r\n {ex.Message} ", "Warning", assetName: "warning.png");
                    }
                });

                // Any code here will run on the UI thread once the file copying is complete
                Status = $"Copy attempt completed ({estCount}/{cpyCount} files)";

            }
            catch (Exception ex)
            {
                // Exception handling can occur here back on the UI thread context
                Status = "Copy attempt failed!";
                Debug.WriteLine($"[ERROR] CopyFilesAsync: {ex.Message}");
                App.ShowDialog($" CopyFilesAsync Error! \r\n {ex.Message} ", "Warning", assetName: "error.png");
            }
        }

        /// <summary>
        /// Copies files from the <paramref name="source"/> to the <paramref name="destination"/> using <see cref="System.IO.File.Copy(string, string)"/>.
        /// </summary>
        void CopyFiles(string source, string destination)
        {
            int estCount = 0;
            int cpyCount = 0;
            try
            {
                Status = "Starting copy…";

                if (!source.EndsWith("\\")) { source += "\\"; }

                foreach (var p in Directory.GetDirectories(source, "*", SearchOption.AllDirectories))
                {
                    Directory.CreateDirectory(System.IO.Path.Combine(destination, p.Substring(source.Length)));
                }
                // A little inefficient, but we need to know how much to copy vs what was copied.
                foreach (var f in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
                {
                    estCount++;
                }
                foreach (var f in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
                {
                    System.IO.File.Copy(f, System.IO.Path.Combine(destination, f.Substring(source.Length)), true);
                    cpyCount++;
                }

                Status = $"Copy completed successfully ({estCount}/{cpyCount} files)";
            }
            catch (Exception ex)
            {
                Status = "File copy failed!";
                Debug.WriteLine($"[ERROR] CopyFiles: {ex.Message}");
                App.ShowDialog($" File copy failed! \r\n {ex.Message} ", "Warning", assetName: "error.png");
            }
        }

        double originalWindowTop = double.NaN;
        void AdjustWindowPosition()
        {
            try
            {
                // Get screen information using the Windows Forms Screen class (more reliable for work area)
                WindowInteropHelper wih = new WindowInteropHelper(this);
                if (wih.Handle == IntPtr.Zero)
                    return;

                Screen currentScreen = Screen.FromHandle(wih.Handle);
                System.Drawing.Rectangle workingArea = currentScreen.WorkingArea;

                // WPF uses logical pixels, System.Windows.Forms uses physical pixels.
                // For standard DPI settings (100%), they match. 

                double screenHeight = workingArea.Height;
                double screenTop = workingArea.Top;

                // Calculate the bottom edge of the window's current position
                double windowBottom = this.Top + this.ActualHeight;

                // Check if the bottom of the window is below the screen's working area
                if (windowBottom > screenHeight + screenTop)
                {
                    // [CASE A]: Window is expanding and going off the bottom of the screen

                    if (double.IsNaN(originalWindowTop))
                    {
                        // Store the position from before the window had to be forced up
                        originalWindowTop = this.Top;
                    }

                    // Calculate how much the window goes off the bottom
                    double offset = windowBottom - (screenHeight + screenTop);

                    // Reposition the window upwards
                    this.Top -= offset;

                    // Ensure the top doesn't go off screen during repositioning
                    if (this.Top < screenTop)
                    {
                        // If even the top can't fit, you might need a ScrollViewer in the Expander content
                        this.Top = screenTop;
                    }
                }
                else if (windowBottom < screenHeight + screenTop)
                {
                    // [CASE B]: Window is collapsing or already fits entirely on screen

                    if (!double.IsNaN(originalWindowTop))
                    {
                        // Restore the window to its original position before the expansion forced it up
                        this.Top = originalWindowTop;
                        originalWindowTop = double.NaN; // Clear the stored position
                    }
                    else
                    {
                        // If it was never forced up, you can optionally recenter it vertically
                        // this.Top = screenTop + (screenHeight - this.ActualHeight) / 2;
                    }
                }
            }
            catch (Exception ex)
            {
                App.ShowDialog($" AdjustWindowPosition Error! \r\n {ex.Message} ", "Warning", assetName: "error.png");
            }
        }

        string CombineColors(string color1, string color2, bool mult = false)
        {
            var tmp1 = (Color)ColorConverter.ConvertFromString(color1);
            var tmp2 = (Color)ColorConverter.ConvertFromString(color2);
            if (!mult)
            {
                var newColor = Color.Add(tmp1, tmp2);
                return newColor.ToString();
            }
            else
            {
                var newColor = Color.Multiply(tmp1, 0.5F);
                return newColor.ToString();
            }
        }

        string BrightenColor(string color)
        {
            var tmp1 = (Color)ColorConverter.ConvertFromString(color);
            var tmp2 = (Color)ColorConverter.ConvertFromString("#AAAAAA");
            var newColor = Color.Add(tmp1, tmp2);
            return newColor.ToString();
        }

       /// <summary>Blends the provided two colors together.</summary>
       /// <param name="foreColor">Color to blend onto the background color.</param>
       /// <param name="backColor">Color to blend the other color onto.</param>
       /// <param name="amount">How much of <paramref name="foreColor"/> to keep, on top of <paramref name="backColor"/>.</param>
       /// <returns>The blended colors.</returns>
       /// <remarks>The alpha channel is not altered.</remarks>
       Color ColorBlend(Color foreColor, Color backColor, double amount = 0.3)
       {
           byte r = (byte)(foreColor.R * amount + backColor.R * (1 - amount));
           byte g = (byte)(foreColor.G * amount + backColor.G * (1 - amount));
           byte b = (byte)(foreColor.B * amount + backColor.B * (1 - amount));
           return Color.FromArgb(255, r, g, b);
       }

        string MakeRandomColor()
        {
            string retVal = "#";
            Random rnd = new Random();
            byte[] b = new byte[3];
            rnd.NextBytes(b); //use the NextBytes() method to fill an array of bytes with random byte values
            for (int i = 0; i <= b.GetUpperBound(0); i++)
            {
                Console.WriteLine("{0}: {1:X2}", i, b[i]);
                retVal += String.Format("{0:X2}", b[i]);
            }
            return retVal;
        }

        #endregion
    }

    //https://stackoverflow.com/questions/1263001/wpf-text-fade-out-then-in-effect
    /*
        Here is an implementation that automatically does the fade-out, switch value, fade in.
        To use (after setting xmlns:l to the correct namespace:
            
            <Label l:AnimatedSwitch.Property="Content" l:AnimatedSwitch.Binding="{Binding SomeProp}"/>
     */
    public class AnimatedSwitch : DependencyObject
    {
        #region [Attached properties]
        public static DependencyProperty BindingProperty = DependencyProperty.RegisterAttached(
            "Binding", 
            typeof(object), 
            typeof(AnimatedSwitch),
            new PropertyMetadata(BindingChanged));
        public static DependencyProperty PropertyProperty = DependencyProperty.RegisterAttached(
            "Property", 
            typeof(string), 
            typeof(AnimatedSwitch));
        #endregion

        public static object GetBinding(DependencyObject e) => e.GetValue(BindingProperty);
        public static void SetBinding(DependencyObject e, object value) =>  e.SetValue(BindingProperty, value);
        public static string GetProperty(DependencyObject e) => (string)e.GetValue(PropertyProperty);
        public static void SetProperty(DependencyObject e, string value) => e.SetValue(PropertyProperty, value);

        /// <summary>
        /// When the value changes do the fadeout-switch-fadein.
        /// </summary>
        static void BindingChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            Storyboard fadeout = new Storyboard();
            var fadeoutAnim = new DoubleAnimation() 
            { 
                To = 0, 
                Duration = new Duration(TimeSpan.FromSeconds(0.3)) 
            };
            Storyboard.SetTarget(fadeoutAnim, d);
            Storyboard.SetTargetProperty(fadeoutAnim, new PropertyPath("Opacity"));
            fadeout.Children.Add(fadeoutAnim);
            fadeout.Completed += (d1, d2) =>
            {
                d.GetType().GetProperty(GetProperty(d)).SetValue(d, GetBinding(d), null);

                Storyboard fadein = new Storyboard();
                var fadeinAnim = new DoubleAnimation()
                { 
                    To = 1, 
                    Duration = new Duration(TimeSpan.FromSeconds(0.3)) 
                };
                Storyboard.SetTarget(fadeinAnim, d);
                Storyboard.SetTargetProperty(fadeinAnim, new PropertyPath("Opacity"));
                fadein.Children.Add(fadeinAnim);
                fadein.Begin();
            };
            fadeout.Begin();
        }
    }
}

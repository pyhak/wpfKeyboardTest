using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using FlaUI.Core.Input;
using Keyboard = FlaUI.Core.Input.Keyboard;
using Window = FlaUI.Core.AutomationElements.Window;
using System.Windows.Forms; 
using System.Drawing;
using TextBox = FlaUI.Core.AutomationElements.TextBox;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Documents;
using System.Windows.Threading;
using System.Windows.Controls;
using TabItem = FlaUI.Core.AutomationElements.TabItem;
using System.Windows.Input;

namespace WpfKeyboard
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            InitializeTimer();

            _notifyIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application, // Siia tuleb lisada oma ikooni fail
                Visible = false,
                Text = "My Application",
                ContextMenuStrip = new ContextMenuStrip()
            };
            _notifyIcon.ContextMenuStrip.Items.Add("Restore", null, RestoreWindow);
            _notifyIcon.ContextMenuStrip.Items.Add("Exit", null, ExitApplication);

            _notifyIcon.DoubleClick += RestoreWindow;
        }

        private bool _textStarted = false;
        private Dictionary<string, string> selected_tab = new();
        private NotifyIcon _notifyIcon;
        private DispatcherTimer _timer;
        private const int DelayMilliseconds = 5000; // 5 sekundit


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = FindWindow("google");
            mainWindow ??= GetActiveWindow();

            var tabItem = FindTabByName("google", mainWindow);
            if (tabItem == null) return;

            tabItem.Focus();
            var textBox = FindTextBoxById("APjFqb", mainWindow);
            if (textBox == null) return;

            textBox.Focus();
            textBox.Enter("Hello world");
        }

 

        private async void Button_Click_1Async(object sender, RoutedEventArgs e)
        {
            Random random = new Random();
            textArea.Focus();
            var text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vulputate, nisl nec volutpat consequatinar nec. Duis eget nulla sed justo placerat hendrerit quis vitae sem. Aliquam at libero at ex feugiat fermentum. Quisque orci dolo.";
            foreach(char c in text) {
                
                int sleepTime = random.Next(10, 1001);
                await Task.Delay(sleepTime);
                textArea.AppendText(c.ToString());
            }           
        }

        private static FlaUI.Core.AutomationElements.Window? GetActiveWindow()
        {
            using var automation = new UIA3Automation();

            var edgeWindows = automation.GetDesktop().FindAllChildren(cf => cf.ByClassName("Chrome_WidgetWin_1"));
            var activeWindow = edgeWindows.FirstOrDefault().AsWindow();
            return activeWindow;
        }

        private static Window? FindWindow(string name)
        {
            using var automation = new UIA3Automation();

            var edgeWindows = automation.GetDesktop().FindAllChildren(cf => cf.ByClassName("Chrome_WidgetWin_1"));
            var searchWindow = edgeWindows.Where(x => x.Name.Contains(name)).FirstOrDefault().AsWindow();
            if (searchWindow == null) { return null; }
            return searchWindow;
        }

        private static TabItem? FindTabByName(String name, Window window)
        {
            var tabItems = window.FindAllDescendants(cf => cf.ByControlType(ControlType.TabItem));

            foreach (var tab in tabItems)
            {
                var tabItem = tab.AsTabItem();

                if (tabItem == null) continue;
                if (tabItem.Name.ToLower().Contains(name.ToLower())) return tabItem;
            }
            return null;
        }

        private TextBox? FindTextBoxByName(String name)
        {
            var mainWindow = GetActiveWindow();
            if (mainWindow == null) return null;

            var tab = FindTabByName(name, mainWindow);
            if (tab == null) return null;
            
            tab.Select();
            var textBox = mainWindow.FindFirstDescendant(cf => cf.ByName(name)).AsTextBox();
            if (textBox == null) return null;

            return textBox;
        }

        private TextBox? FindTextBoxById(String id, Window window)
        {
            var searchBox1 = window.FindFirstDescendant(cf => cf.ByAutomationId(id))?.AsTextBox(); 
            if (searchBox1 == null) return null;

            return searchBox1;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var mainWindow = GetActiveWindow();
            var tabItem = FindTabByName("google", mainWindow);
            if (tabItem == null) return;

                    // Otsime lehelt otsinguvälja
                    //var searchBox = mainWindow.FindFirstDescendant(cf => cf.ByName("q"))?.AsTextBox();
                    var searchBox1 = mainWindow.FindFirstDescendant(cf => cf.ByControlType(ControlType.ComboBox))?.AsTextBox();
                    //var searchBox2 = mainWindow.FindFirstDescendant(cf => cf.ByClassName("gLFyf"))?.AsTextBox();

                    if (searchBox1 != null)
                    {
                        // Sisestame teksti otsinguväljale
                        searchBox1.Enter("Hello, world!!!");
                        
                        // Simuleerime Enter-klahvi, kui soovime otsingu käivitada
                        // Keyboard.Press(VirtualKeyShort.ENTER);

                        Console.WriteLine($"Tekst sisestatud vahekaardile: {tabItem.Name}");
                    }
                }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var window = GetActiveWindow();
            if (window == null) return;

            window.Focus();

            // Saatke Ctrl + Tab, et vahetada vahekaarti
            Keyboard.Press(VirtualKeyShort.CONTROL);
            Keyboard.Press(VirtualKeyShort.TAB);
            Keyboard.Release(VirtualKeyShort.TAB);
            Keyboard.Release(VirtualKeyShort.CONTROL);
        }

        private void OpenUrl(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch
            {
                // hack because of this: https://github.com/dotnet/corefx/issues/10361
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    throw;
                }
            }
        }

        private void InitializeTimer()
        {
            _timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(DelayMilliseconds)
            };
            _timer.Tick += Timer_Tick;
        }

        private void TextArea_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_timer == null || textArea == null)
            {
                return;
            }

            _timer.Stop();
            _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            // Peatage ajastamine ja kutsuge tegevus, kui tekst ei muutu 5 sekundi jooksul
            _timer.Stop();

            var window = FindWindow("Editpad");
            window ??= GetActiveWindow();
            if (window == null) return;
            //window.Focus();

            var tabItem = FindTabByName("Editpad", window);
            if (tabItem == null)
            {
                OpenUrl("https://www.editpad.org/");
                Thread.Sleep(1000);
                tabItem = FindTabByName("Editpad", window);
            }
            

            if (!tabItem.IsSelected) tabItem.Focus();
            var textBox = FindTextBoxById("textarea__editor", window);
            if (textBox == null) return;

            // Hankige RichTextBox-ilt FlowDocument
            FlowDocument document = textArea.Document;

            // Hankige kogu tekst FlowDocument-ilt
            string text = new TextRange(document.ContentStart, document.ContentEnd).Text;

            //textBox.Focus();
            textBox.Text = text;
            

        }

        private void SendTextToTextArea(string input)
        {
            TextPointer caretPosition = textArea.CaretPosition;

            // Liigume kursori kohta, mis vastab algharu algusele (näiteks rea algusesse)
            caretPosition = caretPosition.GetInsertionPosition(LogicalDirection.Forward);

            // Leidke kaugus dokumendi algusest
            int cursorOffset = new TextRange(textArea.Document.ContentStart, caretPosition).Text.Length;
            Console.WriteLine($"Kursor asub indeksil: {cursorOffset}");

            // Näiteks, kui tahad kursori asukoha teada ja sisestada teksti just sinna:
            textArea.CaretPosition = caretPosition;
            textArea.CaretPosition.InsertTextInRun(input);
            textArea.Focus();
            var nextPosition = caretPosition.GetNextInsertionPosition(LogicalDirection.Forward);

            if (nextPosition != null)
            {
                textArea.CaretPosition = nextPosition;
            }
            textArea.Focus();
        }

        private void BtnC_Click(object sender, RoutedEventArgs e)
        {
            SendTextToTextArea("C");
            //Keyboard.Press(VirtualKeyShort.KEY_C);
        }

        private void BtnG_Click(object sender, RoutedEventArgs e)
        {
            SendTextToTextArea("G");
            //Keyboard.Press(VirtualKeyShort.KEY_G);
        }

        private void BtnI_Click(object sender, RoutedEventArgs e)
        {
            SendTextToTextArea("I");
            //Keyboard.Press(VirtualKeyShort.KEY_I);
        }

        #region menubar

        private void MinimizeWindow(object sender, RoutedEventArgs e)
        {
            this.Hide();
            _notifyIcon.Visible = true;
        }

        private void RestoreWindow(object sender, EventArgs e)
        {
            // Taastame akna ja peidame NotifyIconi
            this.Show();
            this.WindowState = WindowState.Normal;
            _notifyIcon.Visible = false;
        }

        private void ExitApplication(object sender, EventArgs e)
        {
            // Sulgeme rakenduse
            _notifyIcon.Dispose();
            System.Windows.Application.Current.Shutdown();
        }

        protected override void OnStateChanged(EventArgs e)
        {
            // Kontrollime, kas aken minimeeriti
            if (WindowState == WindowState.Minimized)
            {
                MinimizeWindow(null, null);
            }
            base.OnStateChanged(e);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // Kui rakendus suletakse, puhastame NotifyIconi
            _notifyIcon.Dispose();
            base.OnClosing(e);
        }

        private void OnTitleBarMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Allow the user to drag the window
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
            {
                DragMove();
            }
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        #endregion

        private void MinimizeToTray(object sender, RoutedEventArgs e)
        {
            this.Hide();
            _notifyIcon.Visible = true;
        }
    }
}


/*
            var ips = new InputSimulator();
            Thread.Sleep(1200);
            //ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
            //ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.F1);
            //ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_C);
            //ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_E);
            string url = "http://www.google.com";
            //Process.Start(new ProcessStartInfo("cmd", $"/c start microsoft-edge:"+url) { CreateNoWindow = true });
            OpenUrl(url);
            Thread.Sleep(1200);
            ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_H);
            ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_E);
            ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_L);
            ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_L);
            ips.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_O);
            Thread.Sleep(1200);
            ips.Keyboard.ModifiedKeyStroke(
                WindowsInput.Native.VirtualKeyCode.RCONTROL,
                WindowsInput.Native.VirtualKeyCode.TAB);
            ips.Keyboard.ModifiedKeyStroke(
                WindowsInput.Native.VirtualKeyCode.RCONTROL,
                WindowsInput.Native.VirtualKeyCode.TAB);
            ips.Keyboard.ModifiedKeyStroke(
                WindowsInput.Native.VirtualKeyCode.RCONTROL,
                WindowsInput.Native.VirtualKeyCode.TAB);
            
            textArea.AppendText("Sended tab");
            */
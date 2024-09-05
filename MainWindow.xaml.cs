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
using System;
using System.Windows.Media;

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

        private bool sendingTextStarted = false;

        private Dictionary<string, string> selected_tab = new();
        private NotifyIcon _notifyIcon;


        private async void Button_Click_1Async(object sender, RoutedEventArgs e)
        {
            Random random = new Random();
            textArea.Focus();
            var text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vulputate, nisl nec volutpat consequatinar nec. Duis eget nulla sed justo placerat hendrerit quis vitae sem. Aliquam at libero at ex feugiat fermentum. Quisque orci dolo.";
            foreach (char c in text)
            {

                int sleepTime = random.Next(10, 1001);
                await Task.Delay(sleepTime);
                textArea.AppendText(c.ToString());
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



        private void TextArea_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }

        //DEPRICATED
        private void Timer_Tick(object sender, EventArgs e)
        {
            // Peatage ajastamine ja kutsuge tegevus, kui tekst ei muutu 5 sekundi jooksul

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



        //          SendTextToTextArea("C");
        //Keyboard.Press(VirtualKeyShort.KEY_C);


        #region helpers
        /// <summary>
        /// Check if there's an active field to send texts
        /// </summary>
        /// <param name="window"></param>
        /// <returns></returns>
        private static bool IsTextInputActive(Window window)
        {
            var focusedElement = window.Automation.FocusedElement();
            if (focusedElement != null)
            {
                if (focusedElement.ControlType == FlaUI.Core.Definitions.ControlType.Edit ||
                    focusedElement.ControlType == FlaUI.Core.Definitions.ControlType.Document)
                {
                    var elementName = focusedElement.Properties.Name.ValueOrDefault;
                    if (!string.IsNullOrEmpty(elementName) && elementName.Contains("Address"))
                    {
                        return false;
                    }

                    var controlType = focusedElement.Properties.LocalizedControlType;
                    if (controlType.Value == "edit" || controlType.Value == "document")
                    {
                        return true;
                    }
                }

                if (focusedElement.ControlType == FlaUI.Core.Definitions.ControlType.Document)
                {
                    var controlType = focusedElement.Properties.ClassName;
                    if (controlType.Value == "RichTextBox")
                    {
                        return true;
                    }
                }
            }
            return false;
        }


        private static Window? GetActiveWindow()
        {
            using var automation = new UIA3Automation();

            var edgeWindows = automation.GetDesktop().FindAllChildren(cf => cf.ByClassName("Chrome_WidgetWin_1"));
            var activeWindow = edgeWindows.FirstOrDefault()?.AsWindow();
            return activeWindow;
        }
        /// <summary>
        /// Find edge windo by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
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

        #endregion helpers

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

        private void MinimizeToTray(object sender, RoutedEventArgs e)
        {
            this.Hide();
            _notifyIcon.Visible = true;
        }


        #endregion

        private async void BtnStartSendText(object sender, RoutedEventArgs e)
        {
            sendingTextStarted = true;

            var text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vulputate, nisl nec volutpat consequatinar nec. Duis eget nulla sed justo placerat hendrerit quis vitae sem. Aliquam at libero at ex feugiat fermentum. Quisque orci dolo.";

            Random random = new Random();

            var window = GetActiveWindow();
            if (window == null) return;

            foreach (char c in text)
            {
                // Kontrollime igal sammul, kas peaksime peatuma
                if (!sendingTextStarted) return;

                int sleepTime = random.Next(10, 1001);
                await Task.Delay(sleepTime);

                // Kontrollime akent uuesti
                var currentWindow = GetActiveWindow();
                if (currentWindow == null) return;

                // Kontrollime, kas tekstiväli on aktiivne
                bool isTextInputActive = IsTextInputActive(currentWindow);

                // Uuendame Tuli.Fill ainult siis, kui olek muutub
                Tuli.Fill = isTextInputActive ? new SolidColorBrush(Colors.Green) : new SolidColorBrush(Colors.Red);

                if (isTextInputActive)
                {
                    // Saadame ainult siis, kui fookus on tekstikastil
                    Keyboard.Type(c);
                }
            }

            // Kui tekst on täielikult saadetud, peatame tsükli
            sendingTextStarted = false;
        }

        private void BtnStopSendText(object sender, RoutedEventArgs e)
        {
            sendingTextStarted = false;
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
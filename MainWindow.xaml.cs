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

        private void SendControlTabToBroweser(object sender, RoutedEventArgs e)
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
                if (focusedElement.ControlType == ControlType.Edit ||
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

        private static TextBox? FindTextBoxById(String id, Window window)
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

        private void BtnStartSendText(object sender, RoutedEventArgs e)
        {
            StartTypeing();
        }

        private async void StartTypeing()
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

        private void SendTextToElement(string windowName, string tabName, string elementId, string text)
        {
            var window = FindWindow(windowName);
            window ??= GetActiveWindow();
            if (window == null) return;

            var tabItem = FindTabByName(tabName, window);
            if (tabItem == null)
            {
                OpenUrl("file:///C:/Dev/forms/WpfKeyboard/info.html"); // Kohanda vastavalt vajadusele
                Thread.Sleep(1000);
                tabItem = FindTabByName(tabName, window);
            }

            if (tabItem != null && !tabItem.IsSelected)
            {
                tabItem.Focus();
                Thread.Sleep(200); // Ootame, kuni tabItem on fookuses
            }

            var element = FindTextBoxById(elementId, window);
            if (element == null) return;

            element.Focus();
            Thread.Sleep(200); // Ootame, kuni element on fookuses

            // Saada tekst elementile
            SendTextToFocusedElement(text);
        }

        private void SendTextToWeb1(object sender, RoutedEventArgs e)
        {
            var text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vulputate, nisl nec volutpat consequatinar nec. Duis eget nulla sed justo placerat hendrerit quis vitae sem. Aliquam at libero at ex feugiat fermentum. Quisque orci dolo.";
            SendTextToElement("Infosüsteem", "Infosüsteem", "input1", text);
        }

        private static void SendTextToFocusedElement(string text)
        {
            foreach (char c in text)
            {
                Thread.Sleep(50); // Väike viivitus iga tähe vahel
                Keyboard.Type(c);
            }
        }

        private void SendTextToWebTextarea(object sender, RoutedEventArgs e)
        {
            var text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vulputate, nisl nec volutpat consequatinar nec. Duis eget nulla sed justo placerat hendrerit quis vitae sem. Aliquam at libero at ex feugiat fermentum. Quisque orci dolo.";
            SendTextToElement("Infosüsteem", "Infosüsteem", "textarea1", text);
        }
    }
}

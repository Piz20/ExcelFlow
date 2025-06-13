using System.Windows;
using WpfControls = System.Windows.Controls;
using WpfMsgBox = System.Windows.MessageBox;
using WpsMsgBoxButton = System.Windows.MessageBoxButton;
using WpsMsgBoxImage = System.Windows.MessageBoxImage;
using ExcelFlow.Views;
using System.Linq;
using ExcelFlow.Utilities; // Pour IClosableView

namespace ExcelFlow
{
    public partial class MainWindow : Window
    {
        private WpfControls.Button? _currentActiveButton;
        private GenerationView? _generationView;
        private SendEmailView? _sendEmailView;

        public MainWindow()
        {
            InitializeComponent();

            _generationView = new GenerationView();
            _sendEmailView = new SendEmailView();

            // Alternative à FindVisualChildren : utiliser FindName ou chercher dans le XAML
            var initialButton = this.FindName("GenerationViewButton") as WpfControls.Button;
            if (initialButton == null)
            {
                // Si FindVisualChildren est défini dans ExcelFlow.Helpers
                initialButton = this.FindVisualChildren<WpfControls.Button>()
                                   .FirstOrDefault(b => b.Tag?.ToString() == "GenerationView");
            }

            NavigateToView("GenerationView", initialButton);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            var currentView = MainContentControl.Content as IClosableView;
            if (currentView?.IsOperationInProgress == true)
            {
                (string message, string title, WpsMsgBoxImage icon) = currentView.GetClosingConfirmation();
                var result = WpfMsgBox.Show(message, title, WpsMsgBoxButton.YesNo, icon);
                if (result == MessageBoxResult.No)
                {
                    e.Cancel = true;
                }
            }

            base.OnClosing(e);
        }

        private void NavigationButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfControls.Button clickedButton)
            {
                string? viewTag = clickedButton.Tag?.ToString();
                if (!string.IsNullOrEmpty(viewTag))
                {
                    NavigateToView(viewTag, clickedButton);
                }
            }
        }

        private void NavigateToView(string viewName, WpfControls.Button? clickedButton)
        {
            if (_currentActiveButton != null && _currentActiveButton != clickedButton)
            {
                _currentActiveButton.Style = (Style)this.FindResource("NavigationButtonStyle");
            }

            switch (viewName)
            {
                case "GenerationView":
                    MainContentControl.Content = _generationView;
                    break;
                case "SendEmailView":
                    MainContentControl.Content = _sendEmailView;
                    break;
                default:
                    WpfMsgBox.Show($"La vue '{viewName}' n'est pas reconnue.", "Erreur de Navigation", WpsMsgBoxButton.OK, WpsMsgBoxImage.Error);
                    return;
            }

            if (clickedButton != null)
            {
                clickedButton.Style = (Style)this.FindResource("ActiveNavigationButtonStyle");
                _currentActiveButton = clickedButton;
            }
        }
    }
}
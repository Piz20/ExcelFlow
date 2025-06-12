using System.Windows;
using WpfControls = System.Windows.Controls;
using WpfMsgBox = System.Windows.MessageBox;
using ExcelFlow.Views;
using System.Linq;

using ExcelFlow.Helpers; // <--- This line is correct

/// <summary>
/// Represents the main window of the ExcelFlow application, providing navigation between different views.
/// </summary>
namespace ExcelFlow
{
    public partial class MainWindow : Window
    {
        private WpfControls.Button? _currentActiveButton;

        // 1. Declare instances of your views
        private GenerationView? _generationView;
        private SendEmailView? _sendEmailView;

        // Constructor for MainWindow
        public MainWindow()
        {
            InitializeComponent();

            // 2. Initialize views once when the MainWindow is created
            _generationView = new GenerationView();
            _sendEmailView = new SendEmailView();

            // Now 'this.FindVisualChildren' should be correctly recognized as an extension method.
            var initialButton = this.FindVisualChildren<WpfControls.Button>()
                                    .FirstOrDefault(b => b.Tag?.ToString() == "GenerationView");

            // Navigate to the initial view (GenerationView)
            NavigateToView("GenerationView", initialButton);
        }

        // Event handler for navigation buttons
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

        /// Method to navigate to the specified view and update the active button style
        private void NavigateToView(string viewName, WpfControls.Button? clickedButton)
        {
            if (_currentActiveButton != null && _currentActiveButton != clickedButton)
            {
                _currentActiveButton.Style = (Style)this.FindResource("NavigationButtonStyle");
            }

            switch (viewName)
            {
                case "GenerationView":
                    // 3. Reuse the existing instance
                    MainContentControl.Content = _generationView;
                    break;
                case "SendEmailView":
                    // 3. Reuse the existing instance
                    MainContentControl.Content = _sendEmailView;
                    break;
                default:
                    WpfMsgBox.Show($"La vue '{viewName}' n'est pas reconnue.", "Erreur de Navigation", MessageBoxButton.OK, MessageBoxImage.Error);
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
// C:\Users\p.eminiant\Desktop\PROJETS\ExcelFlow\frontend\MainWindow.xaml.cs
using System.Windows;
using WpfControls = System.Windows.Controls;
using WpfMsgBox = System.Windows.MessageBox;
using ExcelFlow.Views;
using System.Linq;

using ExcelFlow.Helpers; // <--- ADD THIS LINE! This makes the extension method available

/// <summary>
/// Represents the main window of the ExcelFlow application, providing navigation between different views.
/// </summary>
namespace ExcelFlow
{
    public partial class MainWindow : Window
    {
        private WpfControls.Button? _currentActiveButton;

// Constructor for MainWindow   
        public MainWindow()
        {
            InitializeComponent();

            // Now 'this.FindVisualChildren' should be correctly recognized as an extension method.
            var initialButton = this.FindVisualChildren<WpfControls.Button>()
                                    .FirstOrDefault(b => b.Tag?.ToString() == "GenerationView");

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
                    MainContentControl.Content = new GenerationView();
                    break;
                case "SendEmailView":
                    MainContentControl.Content = new SendEmailView();
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
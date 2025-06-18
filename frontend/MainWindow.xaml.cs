using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using WpfControls = System.Windows.Controls;
using WpfMsgBox = System.Windows.MessageBox;
using WpsMsgBoxButton = System.Windows.MessageBoxButton;
using WpsMsgBoxImage = System.Windows.MessageBoxImage;
using ExcelFlow.Views;
using ExcelFlow.Utilities;
using ExcelFlow.Models;

namespace ExcelFlow
{
    public partial class MainWindow : Window
    {
        private const string SmtpConfigFilePath = "smtpconfig.json";

        private SmtpConfig _smtpConfig = new();

        private GenerationView? _generationView;
        private SendEmailView? _sendEmailView;
        private SmtpSettingsView? _smtpSettingsView;

        private WpfControls.Button? _currentActiveButton;

        public MainWindow()
        {
            InitializeComponent();

            // Charger config au démarrage
            _smtpConfig = LoadSmtpConfig();

            _generationView = new GenerationView();
            _sendEmailView = new SendEmailView();
            _smtpSettingsView = new SmtpSettingsView();

            // S’abonner à l’événement de changement config SMTP
            if (_smtpSettingsView != null)
            {
                _smtpSettingsView.SmtpConfigChanged += OnSmtpConfigChanged;
            }

            ApplySmtpConfigToViews();

            var initialButton = this.FindName("GenerationViewButton") as WpfControls.Button;
            if (initialButton == null)
            {
                initialButton = this.FindVisualChildren<WpfControls.Button>()
                                   .FirstOrDefault(b => b.Tag?.ToString() == "GenerationView");
            }

            NavigateToView("GenerationView", initialButton);
        }

        private void OnSmtpConfigChanged(SmtpConfig newConfig)
        {
            _smtpConfig = newConfig;

            // Propager la config mise à jour aux vues
            ApplySmtpConfigToViews();

            // Sauvegarder dans un fichier JSON
            SaveSmtpConfig();
        }

        private void ApplySmtpConfigToViews()
        {
            if (_sendEmailView != null)
            {
                _sendEmailView.SmtpHost = _smtpConfig.SmtpHost;
                _sendEmailView.SmtpPort = _smtpConfig.SmtpPort ?? 587;  // Valeur par défaut si null
                _sendEmailView.SmtpFromEmail = _smtpConfig.SmtpFromEmail;
            }

            if (_smtpSettingsView != null)
            {
                _smtpSettingsView.SmtpHost = _smtpConfig.SmtpHost;
                _smtpSettingsView.SmtpPort = _smtpConfig.SmtpPort ?? 587;  // Pareil ici
                _smtpSettingsView.SmtpFromEmail = _smtpConfig.SmtpFromEmail;
            }
        }

        private SmtpConfig LoadSmtpConfig()
        {
            try
            {
                if (File.Exists(SmtpConfigFilePath))
                {
                    var json = File.ReadAllText(SmtpConfigFilePath);
                    var config = JsonSerializer.Deserialize<SmtpConfig>(json);
                    if (config != null) return config;
                }
            }
            catch (Exception ex)
            {
                WpfMsgBox.Show($"Erreur lors du chargement de la config SMTP : {ex.Message}", "Erreur", WpsMsgBoxButton.OK, WpsMsgBoxImage.Error);
            }

            // Valeurs par défaut si aucun fichier trouvé ou erreur
            return new SmtpConfig
            {
                SmtpHost = "smtp.example.com",
                SmtpPort = 587,
                SmtpFromEmail = "noreply@example.com"
            };
        }

        private void SaveSmtpConfig()
        {
            try
            {
                var json = JsonSerializer.Serialize(_smtpConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SmtpConfigFilePath, json);
            }
            catch (Exception ex)
            {
                WpfMsgBox.Show($"Erreur lors de la sauvegarde de la config SMTP : {ex.Message}", "Erreur", WpsMsgBoxButton.OK, WpsMsgBoxImage.Error);
            }
        }

        // Gestion clic boutons navigation
        private void NavigationButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfControls.Button clickedButton)
            {
                var viewName = clickedButton.Tag?.ToString();
                if (!string.IsNullOrEmpty(viewName))
                {
                    NavigateToView(viewName, clickedButton);
                }
                else
                {
                    WpfMsgBox.Show("Le bouton n'a pas de Tag défini pour la navigation.", "Erreur", WpsMsgBoxButton.OK, WpsMsgBoxImage.Error);
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
                    ApplySmtpConfigToViews();
                    MainContentControl.Content = _sendEmailView;
                    break;

                case "SmtpSettingsView":
                    ApplySmtpConfigToViews();
                    MainContentControl.Content = _smtpSettingsView;
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

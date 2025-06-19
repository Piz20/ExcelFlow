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
using System.Text.Json;

using ExcelFlow.Models;

namespace ExcelFlow
{
    public partial class MainWindow : Window
    {
        private const string ConfigFilePath = "appconfigs.json";


        private AppConfig _appConfig = new();


        private GenerationView? _generationView;
        private SendEmailView? _sendEmailView;
        private SmtpSettingsView? _smtpSettingsView;

        private WpfControls.Button? _currentActiveButton;
        public MainWindow()
        {
            InitializeComponent();

            // Charger config au démarrage
            _appConfig = LoadAppConfig();

            _generationView = new GenerationView(_appConfig);
            _sendEmailView = new SendEmailView(_appConfig);
            _smtpSettingsView = new SmtpSettingsView();



    


            if (_smtpSettingsView != null)
            {
                _smtpSettingsView.SmtpConfigChanged += OnAppConfigChanged;
            }

            ApplyAppConfigToViews();

            var initialButton = this.FindName("GenerationViewButton") as WpfControls.Button;
            if (initialButton == null)
            {
                initialButton = this.FindVisualChildren<WpfControls.Button>()
                                   .FirstOrDefault(b => b.Tag?.ToString() == "GenerationView");
            }

            NavigateToView("GenerationView", initialButton);
        }

        private void OnAppConfigChanged(SmtpConfig newConfig)
        {
            _appConfig.Smtp = newConfig;

            ApplyAppConfigToViews();
            SaveAppConfig();
        }

        private void ApplyAppConfigToViews()
        {
            if (_sendEmailView != null)
            {
                _sendEmailView.SmtpHost = _appConfig.Smtp.SmtpHost;
                _sendEmailView.SmtpPort = _appConfig.Smtp.SmtpPort ?? 587;
                _sendEmailView.SmtpFromEmail = _appConfig.Smtp.SmtpFromEmail;
            }

            if (_smtpSettingsView != null)
            {
                _smtpSettingsView.SmtpHost = _appConfig.Smtp.SmtpHost;
                _smtpSettingsView.SmtpPort = _appConfig.Smtp.SmtpPort ?? 587;
                _smtpSettingsView.SmtpFromEmail = _appConfig.Smtp.SmtpFromEmail;
            }
        }

        private AppConfig LoadAppConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json = File.ReadAllText(ConfigFilePath);
                    var config = JsonSerializer.Deserialize<AppConfig>(json);
                    if (config != null) return config;
                }
            }
            catch (Exception ex)
            {
                WpfMsgBox.Show($"Erreur lors du chargement de la configuration : {ex.Message}", "Erreur", WpsMsgBoxButton.OK, WpsMsgBoxImage.Error);
            }

            return new AppConfig
            {
                Smtp = new SmtpConfig
                {
                    SmtpHost = "smtp.example.com",
                    SmtpPort = 587,
                    SmtpFromEmail = "noreply@example.com"
                }
            };
        }

        private void SaveAppConfig()
        {
            try
            {
                var json = JsonSerializer.Serialize(_appConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigFilePath, json);
            }
            catch (Exception ex)
            {
                WpfMsgBox.Show($"Erreur lors de la sauvegarde de la configuration : {ex.Message}", "Erreur", WpsMsgBoxButton.OK, WpsMsgBoxImage.Error);
            }
        }

        // Gestion clic boutons navigation
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
                    ApplyAppConfigToViews();
                    MainContentControl.Content = _sendEmailView;
                    break;

                case "SmtpSettingsView":
                    ApplyAppConfigToViews();
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

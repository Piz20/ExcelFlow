using System;
using System.Windows;
using System.Windows.Controls;
using ExcelFlow.Models;
using System.Text.Json;
using System.IO;
using ExcelFlow.Utilities;
using WpfMsgBox = System.Windows.MessageBox;
using WpsMsgBoxImage = System.Windows.MessageBoxImage;
using WpfControls = System.Windows.Controls;

namespace ExcelFlow.Views
{
    public partial class SmtpSettingsView : WpfControls.UserControl, IClosableView
    {
        private readonly AppConfig _appConfig;
        public event Action<SmtpConfig>? SmtpConfigChanged;

        public SmtpSettingsView(AppConfig config)
        {
            InitializeComponent();
            _appConfig = config;

            // Initialiser les champs avec les valeurs existantes
            SmtpHost = _appConfig.Smtp.SmtpHost;
            SmtpPort = _appConfig.Smtp.SmtpPort;
            SmtpFromEmail = _appConfig.Smtp.SmtpFromEmail;
        }

        public string SmtpHost
        {
            get => SmtpHostTextBox.Text.Trim();
            set => SmtpHostTextBox.Text = value ?? string.Empty;
        }

        public int? SmtpPort
        {
            get => int.TryParse(SmtpPortTextBox.Text, out var port) ? port : null;
            set => SmtpPortTextBox.Text = value?.ToString() ?? string.Empty;
        }

        public string SmtpFromEmail
        {
            get => SmtpFromEmailTextBox.Text.Trim();
            set => SmtpFromEmailTextBox.Text = value ?? string.Empty;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (SmtpPort <= 0)
            {
                ShowMsg("Le champ SMTP Port est invalide."); return;
            }

            if (!string.IsNullOrWhiteSpace(SmtpFromEmail) && !IsValidEmail(SmtpFromEmail))
            {
                ShowMsg("L'adresse e-mail de l'expéditeur est invalide."); return;
            }

            _appConfig.Smtp.SmtpHost = SmtpHost;
            _appConfig.Smtp.SmtpPort = SmtpPort;
            _appConfig.Smtp.SmtpFromEmail = SmtpFromEmail;

            SmtpConfigChanged?.Invoke(new SmtpConfig
            {
                SmtpHost = SmtpHost,
                SmtpPort = SmtpPort,
                SmtpFromEmail = SmtpFromEmail
            });

            try
            {
                var json = JsonSerializer.Serialize(_appConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText("appconfigs.json", json);
            }
            catch (Exception ex)
            {
                ShowMsg($"⚠️ Erreur lors de l'enregistrement de la configuration SMTP : {ex.Message}", "Erreur", MessageBoxImage.Warning);
                return;
            }

            ShowMsg("Configuration SMTP sauvegardée avec succès !", "Succès", MessageBoxImage.Information);
        }

        private static bool IsValidEmail(string email)
        {
            try { return new System.Net.Mail.MailAddress(email).Address == email; }
            catch { return false; }
        }

        private static void ShowMsg(string msg, string title = "Erreur", MessageBoxImage icon = MessageBoxImage.Warning) =>
            WpfMsgBox.Show(msg, title, MessageBoxButton.OK, icon);

        public bool IsOperationInProgress => false;

        public (string Message, string Title, WpsMsgBoxImage Icon) GetClosingConfirmation() =>
            ("Une opération est en cours. Voulez-vous vraiment fermer ?", "Confirmation", WpsMsgBoxImage.Warning);
    }
}

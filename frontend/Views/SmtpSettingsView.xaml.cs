using System;
using System.Windows;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Controls;
using ExcelFlow.Models;
using ExcelFlow.Services;
using ExcelFlow.Utilities;
using Microsoft.Win32;
using WpfMsgBox = System.Windows.MessageBox;
using WpsMsgBoxImage = System.Windows.MessageBoxImage;

namespace ExcelFlow.Views
{
    public partial class SmtpSettingsView : UserControl, IClosableView
    {
        private readonly AppConfig _appConfig;
        public event Action<SmtpConfig>? SmtpConfigChanged;

        public SmtpSettingsView(AppConfig config)
        {
            InitializeComponent();
            _appConfig = config;
            SmtpHost = config.SmtpHost;
            SmtpPort = config.SmtpPort;
            SmtpFromEmail = config.SmtpFromEmail;
        }

        public string SmtpHost
        {
            get => SmtpHostTextBox.Text.Trim();
            set => SmtpHostTextBox.Text = value ?? "";
        }

        public int? SmtpPort
        {
            get => int.TryParse(SmtpPortTextBox.Text, out var port) ? port : null;
            set => SmtpPortTextBox.Text = value?.ToString() ?? "";
        }

        public string SmtpFromEmail
        {
            get => SmtpFromEmailTextBox.Text.Trim();
            set => SmtpFromEmailTextBox.Text = value ?? "";
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

            _appConfig.SmtpHost = SmtpHost;
            _appConfig.SmtpPort = SmtpPort;
            _appConfig.SmtpFromEmail = SmtpFromEmail;
            _appConfig.Save();

            SmtpConfigChanged?.Invoke(new SmtpConfig
            {
                SmtpHost = SmtpHost,
                SmtpPort = SmtpPort,
                SmtpFromEmail = SmtpFromEmail
            });

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

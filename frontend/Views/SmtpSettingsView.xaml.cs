using System;
using System.Windows;
using WpfControls = System.Windows.Controls;
using ExcelFlow.Utilities; // Pour IClosableView
using ExcelFlow.Models;
using WinForms = System.Windows.Forms;
using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using System.Collections.Generic;
using System.Linq;
using ExcelFlow.Services;
using WpfMsgBox = System.Windows.MessageBox;
using WpsMsgBoxImage = System.Windows.MessageBoxImage;

using ExcelFlow.Utilities; // Pour IClosableView et AppConst


namespace ExcelFlow.Views
{
    public partial class SmtpSettingsView : WpfControls.UserControl, IClosableView
    {
        public SmtpSettingsView()
        {
            InitializeComponent();
        }

        public string SmtpHost
        {
            get => SmtpHostTextBox.Text.Trim();
            set => SmtpHostTextBox.Text = value ?? string.Empty;
        }

        public int? SmtpPort
        {
            get
            {
                if (int.TryParse(SmtpPortTextBox.Text, out int port))
                    return port;
                return null;
            }
            set => SmtpPortTextBox.Text = value?.ToString() ?? string.Empty;
        }

        public string SmtpFromEmail
        {
            get => SmtpFromEmailTextBox.Text.Trim();
            set => SmtpFromEmailTextBox.Text = value ?? string.Empty;
        }

        // Événement pour notifier les changements SMTP
        public event Action<SmtpConfig>? SmtpConfigChanged;

        // Bouton Sauvegarder : tu dois ajouter ce handler dans le XAML (Click="SaveButton_Click")

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // On ne fait plus de validation obligatoire
            // Mais on peut valider si un champ est rempli, pour éviter les erreurs bêtes

            if (!string.IsNullOrWhiteSpace(SmtpPortTextBox.Text))
            {
                if (!int.TryParse(SmtpPortTextBox.Text, out int port) || port <= 0)
                {
                    WpfMsgBox.Show("Le champ SMTP Port est invalide.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            if (!string.IsNullOrWhiteSpace(SmtpFromEmail) && !IsValidEmail(SmtpFromEmail))
            {
                WpfMsgBox.Show("L'adresse e-mail de l'expéditeur est invalide.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var newConfig = new SmtpConfig
            {
                SmtpHost = this.SmtpHost,                  // Peut être vide
                SmtpPort = this.SmtpPort,                  // Peut être null
                SmtpFromEmail = this.SmtpFromEmail         // Peut être vide
            };

            SmtpConfigChanged?.Invoke(newConfig);

            WpfMsgBox.Show("Configuration SMTP sauvegardée avec succès !", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
        }


        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        // Implémentation IClosableView
        public bool IsOperationInProgress => false;

        public (string Message, string Title, WpsMsgBoxImage Icon) GetClosingConfirmation()
        {
            return ("Une opération est en cours. Voulez-vous vraiment fermer ?", "Confirmation", WpsMsgBoxImage.Warning);
        }

    }



}

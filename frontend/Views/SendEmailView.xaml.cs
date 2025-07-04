using Microsoft.Win32;
using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using WpfControls = System.Windows.Controls;
using WinForms = System.Windows.Forms;
using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using System.Collections.Generic;
using System.Linq;
using ExcelFlow.Models;
using System.Text.Json;
using System.IO;

using ExcelFlow.Services;
using WpfMsgBox = System.Windows.MessageBox;
using ExcelFlow.Utilities; // Pour IClosableView et AppConstants

namespace ExcelFlow.Views
{
    public partial class SendEmailView : WpfControls.UserControl, IClosableView
    {
        private readonly HubConnection _hubConnection;
        private readonly SendEmailService _sendEmailService;
        private CancellationTokenSource? _cts;
        private string _generatedFilesFolderPath = "";
        private string _partnerEmailFilePath = "";

        private AppConfig _appConfig;

        public string SmtpHost { get; set; } = string.Empty;
        public int SmtpPort { get; set; }
        public string SmtpFromEmail { get; set; } = string.Empty;

        public bool IsOperationInProgress => _cts != null && !_cts.IsCancellationRequested;

        public (string Message, string Title, MessageBoxImage Icon) GetClosingConfirmation()
        {
            return (
                "An email sending process is currently running. Closing the application may interrupt the operation. Are you sure you want to exit?",
                "Confirm Exit - Email Sending",
                MessageBoxImage.Warning
            );
        }

        public SendEmailView(AppConfig config)
        {
            InitializeComponent();

            _appConfig = config;

            _sendEmailService = new SendEmailService($"http://localhost:{AppConstants.port}");
            _hubConnection = new HubConnectionBuilder()
                .WithUrl($"http://localhost:{AppConstants.port}/partnerFileHub")
                .WithAutomaticReconnect()
                .Build();

            // Pré-remplissage à partir de la config
            PartnerEmailFilePathTextBox.Text = _appConfig.SendEmail.PartnerEmailFilePath ?? "";
            GeneratedFilesFolderTextBox.Text = _appConfig.SendEmail.GeneratedFilesFolderPath ?? "";
            FromDisplayNameTextBox.Text = _appConfig.SendEmail.FromDisplayName ?? "";
            CcRecipientsTextBox.Text = _appConfig.SendEmail.CcRecipients ?? "";
            BccRecipientsTextBox.Text = _appConfig.SendEmail.BccRecipients ?? "";
            SmtpHost = _appConfig.Smtp.SmtpHost ?? "";
            SmtpPort = _appConfig.Smtp.SmtpPort ?? 0;
            SmtpFromEmail = _appConfig.Smtp.SmtpFromEmail ?? "";

            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressTextBlock.Visibility = Visibility.Collapsed;
            TxtLogs.Text = string.Empty;

            // Configuration des gestionnaires SignalR
            _hubConnection.On<string>("ReceiveMessage", message =>
            {
                Dispatcher.Invoke(() => AppendLog(message));
            });

            _hubConnection.On<string>("ReceiveErrorMessage", message =>
            {
                Dispatcher.Invoke(() => AppendLog($"❌ ERREUR: {message}"));
            });

            _hubConnection.On<ProgressUpdate>("ReceiveProgressUpdate", data =>
            {
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Visibility = Visibility.Visible;
                    ProgressTextBlock.Visibility = Visibility.Visible;
                    ProgressBar.Minimum = 0;
                    ProgressBar.Maximum = data.Total > 0 ? data.Total : 1;
                    ProgressBar.Value = data.Current;
                    ProgressTextBlock.Text = $"{data.Percentage}%";
                    AppendLog(data.Message ?? "");
                });
            });

            _hubConnection.On<List<PartnerInfo>>("ReceiveIdentifiedPartners", partners =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppendLog($"Reçu {partners.Count} partenaires identifiés.");
                });
            });

            _hubConnection.On<SentEmailSummary>("ReceiveSentEmailSummary", summary =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppendLog($"Email envoyé : '{summary.FileName}' à '{summary.PartnerName}'.");
                });
            });

            _hubConnection.On<int>("ReceiveTotalFilesCount", totalFiles =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppendLog($"Total de fichiers détectés : {totalFiles}.");
                });
            });

            _hubConnection.Reconnected += (sender) =>
            {
                Dispatcher.Invoke(() => AppendLog("🔌 Reconnexion au hub réussie."));
                return Task.CompletedTask;
            };
            _hubConnection.Reconnecting += (ex) =>
            {
                Dispatcher.Invoke(() => AppendLog($"🔌 Reconnexion au hub en cours... {ex?.Message}"));
                return Task.CompletedTask;
            };
            _hubConnection.Closed += (ex) =>
            {
                Dispatcher.Invoke(() => AppendLog($"❌ Connexion au hub fermée : {ex?.Message}"));
                return Task.CompletedTask;
            };

            this.Loaded += SendEmailView_Loaded;
            SetUiEnabledState(true);
        }

        private async void SendEmailView_Loaded(object sender, RoutedEventArgs e)
        {
            await StartSignalRConnection();
        }

        private async Task StartSignalRConnection()
        {
            try
            {
                if (_hubConnection.State == HubConnectionState.Disconnected)
                {
                    await _hubConnection.StartAsync();
                    AppendLog("🔌 Connecté au Service d'Envoi de mails.");
                }
            }
            catch (Exception ex)
            {
                AppendLog($"❌ Impossible de se connecter au SignalR Hub: {ex.Message}");
                WpfMsgBox.Show($"Impossible de se connecter au service. Assurez-vous que le backend est en cours d'exécution.\nErreur: {ex.Message}", "Erreur de Connexion", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AppendLog(string message)
        {
            TxtLogs.AppendText(message + Environment.NewLine);
            TxtLogs.ScrollToEnd();
        }

        private void ClearGeneratedFilesFolderButton_Click(object sender, RoutedEventArgs e)
        {
            GeneratedFilesFolderTextBox.Clear();
            _generatedFilesFolderPath = string.Empty;
        }

        private void ClearPartnerEmailFileButton_Click(object sender, RoutedEventArgs e)
        {
            PartnerEmailFilePathTextBox.Clear();
            _partnerEmailFilePath = string.Empty;
        }

        private void ClearFromDisplayNameButton_Click(object sender, RoutedEventArgs e)
        {
            FromDisplayNameTextBox.Clear();
        }

        private void ClearCcRecipientsButton_Click(object sender, RoutedEventArgs e)
        {
            CcRecipientsTextBox.Clear();
        }

        private void ClearBccRecipientsButton_Click(object sender, RoutedEventArgs e)
        {
            BccRecipientsTextBox.Clear();
        }

        private void BrowseGeneratedFilesButton_Click(object sender, RoutedEventArgs e)
        {
            using var dialog = new WinForms.FolderBrowserDialog
            {
                Description = "Sélectionnez le dossier des fichiers partenaires générés",
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                _generatedFilesFolderPath = dialog.SelectedPath;
                GeneratedFilesFolderTextBox.Text = _generatedFilesFolderPath;
            }
        }

        private void BrowsePartnerEmailFileButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog
            {
                Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Tous les fichiers (*.*)|*.*",
                Title = "Sélectionnez le fichier des mails des partenaires"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _partnerEmailFilePath = openFileDialog.FileName;
                PartnerEmailFilePathTextBox.Text = _partnerEmailFilePath;
            }
        }

        private async void StartSendingButton_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
            AppendLog("🚀 Préparation des emails en cours...\n");

            // Récupération des valeurs de l'UI
            _generatedFilesFolderPath = GeneratedFilesFolderTextBox.Text.Trim();
            _partnerEmailFilePath = PartnerEmailFilePathTextBox.Text.Trim();
            string fromDisplayName = FromDisplayNameTextBox.Text.Trim();
            string ccText = CcRecipientsTextBox.Text.Trim();
            string bccText = BccRecipientsTextBox.Text.Trim();
            ;

            if (string.IsNullOrWhiteSpace(_generatedFilesFolderPath) || string.IsNullOrWhiteSpace(_partnerEmailFilePath))
            {
                WpfMsgBox.Show("Veuillez spécifier le dossier des fichiers générés et le fichier des mails des partenaires.", "Champs manquants", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Mise à jour de la config en mémoire
            _appConfig.SendEmail.GeneratedFilesFolderPath = _generatedFilesFolderPath;
            _appConfig.SendEmail.PartnerEmailFilePath = _partnerEmailFilePath;
            _appConfig.SendEmail.FromDisplayName = fromDisplayName;
            _appConfig.SendEmail.CcRecipients = ccText;
            _appConfig.SendEmail.BccRecipients = bccText;

            // Sauvegarde de la config sur disque
            SaveAppConfig();

            SetUiEnabledState(false);

            SmtpHost = _appConfig.Smtp.SmtpHost ?? "";
            SmtpPort = _appConfig.Smtp.SmtpPort ?? 0;
            SmtpFromEmail = _appConfig.Smtp.SmtpFromEmail ?? "";

            var request = new PrepareEmailRequest
            {
                GeneratedFilesFolder = _generatedFilesFolderPath,
                PartnerExcelPath = _partnerEmailFilePath,
                FromDisplayName = fromDisplayName,
                CcRecipients = EmailUtils.ExtractEmails(ccText),
                BccRecipients = EmailUtils.ExtractEmails(bccText),

                SmtpHost = SmtpHost,
                SmtpPort = SmtpPort,
                SmtpFromEmail = SmtpFromEmail
            };

            _cts = new CancellationTokenSource();

            try
            {
                var preparedEmails = await _sendEmailService.PrepareEmailsAsync(request, _cts.Token);

                if (preparedEmails == null || preparedEmails.Count == 0)
                {
                    WpfMsgBox.Show("Aucun email à préparer.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var previewWindow = new EmailPreviewWindow(preparedEmails);
                previewWindow.Show(); // Non bloquant, la fenêtre parente reste active

            }
            catch (Exception ex)
            {
                WpfMsgBox.Show($"❌ Erreur lors de la préparation : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                AppendLog($"❌ Erreur lors de la préparation : {ex.Message}");
            }
            finally
            {
                SetUiEnabledState(true);
                _cts?.Dispose();
                _cts = null;
            }
        }

        private void CancelSendingButton_Click(object sender, RoutedEventArgs e)
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
                AppendLog("🛑 Demande d'annulation envoyée au service.");
            }
        }

        private void ClearLogs_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
        }

        private void SetUiEnabledState(bool enabled)
        {
            StartSendingButton.IsEnabled = enabled;
            BrowseGeneratedFilesButton.IsEnabled = enabled;
            BrowsePartnerEmailFileButton.IsEnabled = enabled;

            ClearGeneratedFilesFolderButton.IsEnabled = enabled;
            ClearPartnerEmailFileButton.IsEnabled = enabled;
            ClearFromDisplayNameButton.IsEnabled = enabled;
            ClearCcRecipientsButton.IsEnabled = enabled;
            ClearBccRecipientsButton.IsEnabled = enabled;

            GeneratedFilesFolderTextBox.IsEnabled = enabled;
            PartnerEmailFilePathTextBox.IsEnabled = enabled;
            FromDisplayNameTextBox.IsEnabled = enabled;
            CcRecipientsTextBox.IsEnabled = enabled;
            BccRecipientsTextBox.IsEnabled = enabled;
            CancelSendingButton.IsEnabled = !enabled;
        }
        private void SaveAppConfig()
        {
            try
            {
                var json = JsonSerializer.Serialize(_appConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText("appconfigs.json", json); // <-- nom du fichier corrigé ici
                AppendLog("✅ Configuration sauvegardée avec succès dans appconfigs.json.");
            }
            catch (Exception ex)
            {
                AppendLog($"❌ Erreur lors de la sauvegarde de la configuration : {ex.Message}");
            }
        }

    }
}

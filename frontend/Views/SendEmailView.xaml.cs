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

// Références vers vos projets de modèles et de services
using ExcelFlow.Models;
using ExcelFlow.Services;
using WpfMsgBox = System.Windows.MessageBox; // Alias for System.Windows.MessageBox
using ExcelFlow.Utilities; // IMPORTANT : Ajoutez cette ligne pour accéder à AppConstants

namespace ExcelFlow.Views
{
    public partial class SendEmailView : WpfControls.UserControl
    {
        private readonly HubConnection _hubConnection;
        private readonly SendEmailService _sendEmailService;
        private CancellationTokenSource? _cts;

        private string _generatedFilesFolderPath = "";
        private string _partnerEmailFilePath = "";

        public SendEmailView()
        {
            InitializeComponent();

            _sendEmailService = new SendEmailService("https://localhost:7274"); // Vérifiez l'URL
            _hubConnection = new HubConnectionBuilder()
                .WithUrl("https://localhost:7274/partnerFileHub") // Vérifiez l'URL du hub
                .WithAutomaticReconnect()
                .Build();

            // >>> MODIFICATION ICI : Pré-remplir le TextBox avec la constante
            FromDisplayNameTextBox.Text = AppConstants.DefaultFromDisplayName;

            // Initialisation des visibilités des éléments de progression
            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressTextBlock.Visibility = Visibility.Collapsed;
            ProgressMessageTextBlock.Visibility = Visibility.Collapsed;
            TxtLogs.Text = string.Empty; // S'assurer que les logs sont vides au démarrage de la vue

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
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "Sélectionnez le dossier des fichiers partenaires générés";
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    _generatedFilesFolderPath = dialog.SelectedPath;
                    GeneratedFilesFolderTextBox.Text = _generatedFilesFolderPath;
                }
            }
        }

        private void BrowsePartnerEmailFileButton_Click(object sender, RoutedEventArgs e)
        {
            WpfOpenFileDialog openFileDialog = new WpfOpenFileDialog
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
            AppendLog("🚀 Début du processus d'envoi d'emails...\n");

            ProgressBar.Value = 0;
            ProgressTextBlock.Text = "0%";
            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressTextBlock.Visibility = Visibility.Collapsed;
            ProgressMessageTextBlock.Visibility = Visibility.Collapsed;

            if (string.IsNullOrWhiteSpace(_generatedFilesFolderPath) || string.IsNullOrWhiteSpace(_partnerEmailFilePath))
            {
                WpfMsgBox.Show("Veuillez spécifier le dossier des fichiers générés et le fichier des mails des partenaires.", "Champs manquants", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string fromDisplayName = FromDisplayNameTextBox.Text.Trim();
            List<string> ccRecipients = CcRecipientsTextBox.Text.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
            List<string> bccRecipients = BccRecipientsTextBox.Text.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();

            SetUiEnabledState(false);

            var request = new EmailSendRequest
            {
                GeneratedFilesFolderPath = _generatedFilesFolderPath,
                PartnerEmailFilePath = _partnerEmailFilePath,
                FromDisplayName = fromDisplayName,
                CcRecipients = ccRecipients,
                BccRecipients = bccRecipients
            };

            _cts = new CancellationTokenSource();

            try
            {
                var resultMessage = await _sendEmailService.StartEmailSendingAsync(request, _cts.Token);
                AppendLog(resultMessage);

                if (resultMessage.StartsWith("❌"))
                {
                    WpfMsgBox.Show(resultMessage, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else if (_cts.IsCancellationRequested)
                {
                    // Déjà géré par OperationCanceledException
                }
                else
                {
                    WpfMsgBox.Show("🎉 Processus d'envoi d'emails terminé avec succès !", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (OperationCanceledException)
            {
                WpfMsgBox.Show("L'opération d'envoi d'emails a été annulée.", "Annulation", MessageBoxButton.OK, MessageBoxImage.Information);
                AppendLog("🛑 Opération annulée par l'utilisateur.");
            }
            catch (Exception ex)
            {
                WpfMsgBox.Show($"Une erreur inattendue est survenue : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                AppendLog($"❌ Erreur lors de l'envoi d'emails : {ex.Message}");
            }
            finally
            {
                SetUiEnabledState(true);
                ProgressBar.Visibility = Visibility.Collapsed;
                ProgressTextBlock.Visibility = Visibility.Collapsed;
                ProgressMessageTextBlock.Visibility = Visibility.Collapsed;
                _cts?.Dispose();
                _cts = null;
                AppendLog("Processus d'envoi d'emails terminé ou annulé.");
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
    }
}
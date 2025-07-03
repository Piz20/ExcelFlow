using Microsoft.Win32;
using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using WpfControls = System.Windows.Controls;
using ExcelFlow.Models;
using WinForms = System.Windows.Forms;
using WpfMsgBox = System.Windows.MessageBox;
using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using System.IO;
using System.Text.Json;

using ExcelFlow.Services;
using ExcelFlow.Utilities; // Contient IClosableView

namespace ExcelFlow.Views
{
    public partial class GenerationView : WpfControls.UserControl, IClosableView
    {
        private readonly HubConnection _hubConnection;
        private readonly GenerationService _generationService;
        private CancellationTokenSource? _cts;


        private AppConfig _appConfig;
        public GenerationView(AppConfig config)
        {
            InitializeComponent();

            _appConfig = config;

            // Pré-remplir les champs s’il y a des valeurs en mémoire
            TxtSourceFilePath.Text = _appConfig.Generation.SourcePath ?? "";
            TxtTemplateFilePath.Text = _appConfig.Generation.TemplatePath ?? "";
            TxtOutputDir.Text = _appConfig.Generation.OutputDir ?? "";

            _generationService = new GenerationService($"http://localhost:{AppConstants.port}");

            _hubConnection = new HubConnectionBuilder()
                .WithUrl($"http://localhost:{AppConstants.port}/partnerFileHub")
                .WithAutomaticReconnect()
                .Build();

            // Initialisation de la visibilité des éléments de progression
            ProgressGeneration.Visibility = Visibility.Collapsed;
            ProgressPercentageText.Visibility = Visibility.Collapsed;
            TxtLogs.Text = string.Empty;

            // Initialisation de l'état de l'UI
            SetUiEnabledState(true);

            _hubConnection.On<string>("ReceiveMessage", message =>
            {
                Dispatcher.Invoke(() => AppendLog(message));
            });

            _hubConnection.On<ProgressUpdate>("ReceiveProgress", data =>
            {
                var current = data.Current;
                var total = data.Total;
                var percentage = data.Percentage;
                var message = data.Message;

                Dispatcher.Invoke(() =>
                {
                    ProgressGeneration.Visibility = Visibility.Visible;
                    ProgressPercentageText.Visibility = Visibility.Visible;

                    ProgressGeneration.Minimum = 0;
                    ProgressGeneration.Maximum = total;
                    ProgressGeneration.Value = current;

                    ProgressPercentageText.Text = $"{percentage}%";

                    AppendLog(message ?? "");
                });
            });

            StartSignalRConnection();
        }

        // Implémentation de l'interface IClosableView
        public bool IsOperationInProgress => _cts != null && !_cts.IsCancellationRequested;

        public (string Message, string Title, MessageBoxImage Icon) GetClosingConfirmation()
        {
            return (
                "A file generation process is currently running. Closing the application may result in incomplete files. Are you sure you want to exit?",
                "Confirm Exit - File Generation",
                MessageBoxImage.Warning
            );
        }

        private async void StartSignalRConnection()
        {
            try
            {
                if (_hubConnection.State == HubConnectionState.Disconnected)
                {
                    await _hubConnection.StartAsync();
                    AppendLog("🔌 Connecté au service de génération");
                }
            }
            catch (Exception ex)
            {
                AppendLog("❌ Erreur de connexion SignalR: " + ex.Message);
                WpfMsgBox.Show($"Impossible de se connecter au service de génération. Assurez-vous que le backend est en cours d'exécution.\nErreur: {ex.Message}", "Erreur de Connexion", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AppendLog(string message)
        {
            TxtLogs.AppendText(message + Environment.NewLine);
            TxtLogs.ScrollToEnd();
        }

        private void BtnSelectSourceFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog { Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                TxtSourceFilePath.Text = openFileDialog.FileName;
            }
        }

        private void BtnSelectTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog { Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                TxtTemplateFilePath.Text = openFileDialog.FileName;
            }
        }

        private void BtnSelectOutputDir_Click(object sender, RoutedEventArgs e)
        {
            using var dialog = new WinForms.FolderBrowserDialog
            {
                Description = "Sélectionner le dossier de sortie",
                ShowNewFolderButton = true
            };

            var result = dialog.ShowDialog();

            if (result == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
            {
                TxtOutputDir.Text = dialog.SelectedPath;
            }
        }

        private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
            AppendLog("🚀 Début de la génération veuillez patienter...\n");

            if (string.IsNullOrEmpty(TxtSourceFilePath.Text) ||
                string.IsNullOrEmpty(TxtTemplateFilePath.Text) ||
                string.IsNullOrEmpty(TxtOutputDir.Text))
            {
                WpfMsgBox.Show("Merci de sélectionner tous les fichiers et dossiers requis.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SetUiEnabledState(false);

            var request = new GenerationRequest
            {
                filePath = TxtSourceFilePath.Text,
                templatePath = TxtTemplateFilePath.Text,
                outputDir = TxtOutputDir.Text,
                sheetName = "Analyse",
                startIndex = int.TryParse(TxtStartIndex.Text, out int si) ? si : 0,
                count = int.TryParse(TxtCount.Text, out int c) ? c : 200
            };

            try
            {
                ProgressGeneration.Minimum = 0;
                ProgressGeneration.Maximum = 1;
                ProgressGeneration.Value = 0;
                ProgressPercentageText.Text = "0%";

                ProgressGeneration.Visibility = Visibility.Visible;
                ProgressPercentageText.Visibility = Visibility.Visible;

                _cts = new CancellationTokenSource();

                var resultMessage = await _generationService.GenerateAsync(request, _cts.Token);

                if (resultMessage.StartsWith("❌"))
                {
                    WpfMsgBox.Show(resultMessage, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    WpfMsgBox.Show("🎉 Génération réussie !", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (OperationCanceledException)
            {
                WpfMsgBox.Show("L'opération a été annulée.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                AppendLog("🛑 Opération annulée.");
            }
            catch (Exception ex)
            {
                WpfMsgBox.Show($"Erreur lors de la requête : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                AppendLog($"❌ Erreur lors de la génération : {ex.Message}");
            }
            finally
            {
                SetUiEnabledState(true);
                ProgressGeneration.Visibility = Visibility.Collapsed;
                ProgressPercentageText.Visibility = Visibility.Collapsed;
                _cts?.Dispose();
                _cts = null;

                // Mise à jour de la config par défaut
                _appConfig.Generation.SourcePath = TxtSourceFilePath.Text;
                _appConfig.Generation.TemplatePath = TxtTemplateFilePath.Text;
                _appConfig.Generation.OutputDir = TxtOutputDir.Text;

                // Sauvegarde dans le fichier
                try
                {
                    var json = JsonSerializer.Serialize(_appConfig, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText("appconfigs.json", json);  // Ou ton chemin config
                }
                catch (Exception ex)
                {
                    AppendLog($"⚠️ Erreur lors de l'enregistrement des préférences : {ex.Message}");
                }

                AppendLog("Processus de génération terminé ou annulé.");
            }
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
                AppendLog("🛑 Annulation demandée au service.");
            }
        }

        private void ClearLogs_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
        }

        private void ClearSourceFileButton_Click(object sender, RoutedEventArgs e)
        {
            TxtSourceFilePath.Text = string.Empty;
        }

        private void ClearTemplateFileButton_Click(object sender, RoutedEventArgs e)
        {
            TxtTemplateFilePath.Text = string.Empty;
        }

        private void ClearOutputDirButton_Click(object sender, RoutedEventArgs e)
        {
            TxtOutputDir.Text = string.Empty;
        }

        private void SetUiEnabledState(bool enabled)
        {
            TxtSourceFilePath.IsEnabled = enabled;
            TxtTemplateFilePath.IsEnabled = enabled;
            TxtOutputDir.IsEnabled = enabled;
            TxtStartIndex.IsEnabled = enabled;
            TxtCount.IsEnabled = enabled;

            BtnSelectSourceFile.IsEnabled = enabled;
            BtnSelectTemplateFile.IsEnabled = enabled;
            BtnSelectOutputDir.IsEnabled = enabled;

            ClearSourceFileButton.IsEnabled = enabled;
            ClearTemplateFileButton.IsEnabled = enabled;
            ClearOutputDirButton.IsEnabled = enabled;

            BtnGenerate.IsEnabled = enabled;
            BtnStop.IsEnabled = !enabled;
        }
    }
}
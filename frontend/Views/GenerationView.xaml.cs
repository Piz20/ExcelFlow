using Microsoft.Win32;
using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows; // This namespace contains MessageBox, MessageBoxButton, MessageBoxImage
using WpfControls = System.Windows.Controls; // This alias is for UserControl and other Controls
using ExcelFlow.Models;
using WinForms = System.Windows.Forms;
using WpfMsgBox = System.Windows.MessageBox; // Alias for System.Windows.MessageBox
using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using System.IO;

namespace ExcelFlow.Views
{
    public partial class GenerationView : WpfControls.UserControl
    {
        private readonly HubConnection _hubConnection;
        private readonly GenerationService _generationService;
        private CancellationTokenSource? _cts;

        private string sourceFilePath = "";
        private string templateFilePath = "";
        private string outputDir = "";


/// <summary>
        /// Constructeur de la vue de génération.
        public GenerationView()
        {
            InitializeComponent();

            _generationService = new GenerationService("http://localhost:5297");

            _hubConnection = new HubConnectionBuilder()
                .WithUrl("http://localhost:5297/partnerFileHub")
                .WithAutomaticReconnect()
                .Build();

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

  // <summary>
        /// Démarre la connexion SignalR au service de génération.
        private async void StartSignalRConnection()
        {
            try
            {
                await _hubConnection.StartAsync();
                AppendLog("🔌 Connecté au service de génération");
            }
            catch (Exception ex)
            {
                AppendLog("❌ Erreur de connexion SignalR: " + ex.Message);
            }
        }

// <summary>
        /// Ajoute un message au journal de logs.
        private void AppendLog(string message)
        {
            TxtLogs.AppendText(message + Environment.NewLine);
            TxtLogs.ScrollToEnd();
        }


// <summary>
        /// Sélectionne le fichier source, le modèle et le dossier de sortie.
        private void BtnSelectSourceFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog { Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                sourceFilePath = openFileDialog.FileName;
                TxtSourceFilePath.Text = sourceFilePath;
            }
        }

// <summary>
        /// Sélectionne le fichier modèle pour la génération.
        private void BtnSelectTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog { Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                templateFilePath = openFileDialog.FileName;
                TxtTemplateFilePath.Text = templateFilePath;
            }
        }

// <summary>
        /// Sélectionne le dossier de sortie pour les fichiers générés.
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
                outputDir = dialog.SelectedPath;
                TxtOutputDir.Text = outputDir;
            }
        }

// <summary>
        /// Gère le clic sur le bouton de génération.
        private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
            AppendLog("🚀 Début de la génération veuillez patienter...\n");

            if (string.IsNullOrEmpty(sourceFilePath) ||
                string.IsNullOrEmpty(templateFilePath) ||
                string.IsNullOrEmpty(outputDir))
            {
                // Removed WpfControls. prefix
                WpfMsgBox.Show("Merci de sélectionner tous les fichiers et dossiers requis.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            BtnGenerate.IsEnabled = false;
            BtnStop.IsEnabled = true;

            var request = new GenerationRequest
            {
                filePath = sourceFilePath,
                templatePath = templateFilePath,
                outputDir = outputDir,
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
                    // Removed WpfControls. prefix
                    WpfMsgBox.Show(resultMessage, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    // Removed WpfControls. prefix
                    WpfMsgBox.Show("🎉 Génération réussie !", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (OperationCanceledException)
            {
                // Removed WpfControls. prefix
                WpfMsgBox.Show("L'opération a été annulée.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                AppendLog("🛑 Opération annulée.");
            }
            catch (Exception ex)
            {
                // Removed WpfControls. prefix
                WpfMsgBox.Show($"Erreur lors de la requête : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                AppendLog($"❌ Erreur lors de la génération : {ex.Message}");
            }
            finally
            {
                BtnGenerate.IsEnabled = true;
                BtnStop.IsEnabled = false;
                ProgressGeneration.Visibility = Visibility.Collapsed;
                ProgressPercentageText.Visibility = Visibility.Collapsed;
                _cts?.Dispose();
                _cts = null;
                AppendLog("Processus de génération terminé ou annulé.");
            }
        }

// <summary>
        /// Gère le clic sur le bouton d'arrêt de la génération.
        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
                AppendLog("🛑 Annulation demandée au service.");
            }
        }
    }
}
using Microsoft.Win32;
using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using ExcelFlow.Models; // Assurez-vous que ce namespace correspond à l'emplacement de votre modèle ProgressUpdate
using WinForms = System.Windows.Forms;
using WpfMsgBox = System.Windows.MessageBox;
using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
// using ExcelFlow.Models; // Décommentez ceci si votre ProgressUpdate est dans ExcelFlow.Models

namespace ExcelFlow
{
    public partial class MainWindow : Window
    {
        private readonly HubConnection _hubConnection;
        private readonly GenerationService _generationService;
        private CancellationTokenSource? _cts;

        private string sourceFilePath = "";
        private string templateFilePath = "";
        private string outputDir = "";

        public MainWindow()
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

            // *** MODIFICATION MAJEURE ICI : Le type attendu est maintenant ProgressUpdate ***
            _hubConnection.On<ProgressUpdate>("ReceiveProgress", data =>
            {
                // Plus besoin de GetProperty().GetValue(), les propriétés sont directement accessibles
                var current = data.Current;
                var total = data.Total;
                var percentage = data.Percentage;
                var message = data.Message;

                Console.WriteLine($"[DEBUG-UI-UPDATE] Attempting UI update (Received values):");
                Console.WriteLine($"  - ProgressBar: Value={current}, Max={total}");
                Console.WriteLine($"  - TextBlock: Text='{percentage}%'");
                Console.WriteLine($"  - Log: Message='{message}'");


                Dispatcher.Invoke(() =>
                {
                    // Rendez les éléments visibles au début de la progression si ce n'est pas déjà fait
                    ProgressGeneration.Visibility = Visibility.Visible;
                    ProgressPercentageText.Visibility = Visibility.Visible;

                    ProgressGeneration.Minimum = 0;
                    ProgressGeneration.Maximum = total;
                    ProgressGeneration.Value = current;

                    ProgressPercentageText.Text = $"{percentage}%";

                    AppendLog(message ?? "");

                    // Optionnel : change la couleur du texte pour confirmer visuellement qu'il est mis à jour
                    // ProgressPercentageText.Foreground = (percentage % 2 == 0) ? System.Windows.Media.Brushes.Blue : System.Windows.Media.Brushes.DarkGreen;
                });
            });

            StartSignalRConnection();
        }

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
                sourceFilePath = openFileDialog.FileName;
                TxtSourceFilePath.Text = sourceFilePath;
            }
        }

        private void BtnSelectTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog { Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                templateFilePath = openFileDialog.FileName;
                TxtTemplateFilePath.Text = templateFilePath;
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
                outputDir = dialog.SelectedPath;
                TxtOutputDir.Text = outputDir;
            }

        }

        private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
            AppendLog("🚀 Début de la génération veuillez patienter...\n");

            if (string.IsNullOrEmpty(sourceFilePath) ||
                string.IsNullOrEmpty(templateFilePath) ||
                string.IsNullOrEmpty(outputDir))
            {
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
                // Initialisation UI au début de la génération
                ProgressGeneration.Minimum = 0;
                ProgressGeneration.Maximum = 1; // Valeur temporaire, sera mise à jour par le hub
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
                BtnGenerate.IsEnabled = true;
                BtnStop.IsEnabled = false;
                ProgressGeneration.Visibility = Visibility.Collapsed;
                ProgressPercentageText.Visibility = Visibility.Collapsed;
                _cts?.Dispose();
                _cts = null;
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
    }

  
   
}
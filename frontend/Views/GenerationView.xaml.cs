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
using ExcelFlow.Services; // Assurez-vous que ce using est présent si GenerationService est dans ce namespace.

namespace ExcelFlow.Views
{
    public partial class GenerationView : WpfControls.UserControl
    {
        private readonly HubConnection _hubConnection;
        private readonly GenerationService _generationService;
        private CancellationTokenSource? _cts;

        /// <summary>
        /// Constructeur de la vue de génération.
        /// </summary>
        public GenerationView()
        {
            InitializeComponent();

            _generationService = new GenerationService("http://localhost:5297");

            _hubConnection = new HubConnectionBuilder()
                .WithUrl("http://localhost:5297/partnerFileHub")
                .WithAutomaticReconnect()
                .Build();

            // Initialisation de la visibilité des éléments de progression
            ProgressGeneration.Visibility = Visibility.Collapsed;
            ProgressPercentageText.Visibility = Visibility.Collapsed;
            TxtLogs.Text = string.Empty; // S'assurer que les logs sont vides au démarrage

            // Initialisation de l'état de l'UI
            SetUiEnabledState(true); // Tous les contrôles sont activés au démarrage

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

        /// <summary>
        /// Démarre la connexion SignalR au service de génération.
        /// </summary>
        private async void StartSignalRConnection()
        {
            try
            {
                if (_hubConnection.State == HubConnectionState.Disconnected) // Vérifier si la connexion n'est pas déjà établie
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

        /// <summary>
        /// Ajoute un message au journal de logs.
        /// </summary>
        private void AppendLog(string message)
        {
            TxtLogs.AppendText(message + Environment.NewLine);
            TxtLogs.ScrollToEnd();
        }


        /// <summary>
        /// Sélectionne le fichier source.
        /// </summary>
        private void BtnSelectSourceFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog { Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                TxtSourceFilePath.Text = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// Sélectionne le fichier modèle pour la génération.
        /// </summary>
        private void BtnSelectTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new WpfOpenFileDialog { Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == true)
            {
                TxtTemplateFilePath.Text = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// Sélectionne le dossier de sortie pour les fichiers générés.
        /// </summary>
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

        /// <summary>
        /// Gère le clic sur le bouton de génération.
        /// </summary>
        private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
            AppendLog("🚀 Début de la génération veuillez patienter...\n");

            // Utilisez directement les TextBoxes pour la validation
            if (string.IsNullOrEmpty(TxtSourceFilePath.Text) ||
                string.IsNullOrEmpty(TxtTemplateFilePath.Text) ||
                string.IsNullOrEmpty(TxtOutputDir.Text))
            {
                WpfMsgBox.Show("Merci de sélectionner tous les fichiers et dossiers requis.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Désactive l'UI au début de l'opération
            SetUiEnabledState(false);

            var request = new GenerationRequest
            {
                filePath = TxtSourceFilePath.Text,
                templatePath = TxtTemplateFilePath.Text,
                outputDir = TxtOutputDir.Text,
                sheetName = "Analyse", // Considérez si cela devrait être configurable
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
                // Réactive l'UI à la fin de l'opération (succès, échec ou annulation)
                SetUiEnabledState(true);
                ProgressGeneration.Visibility = Visibility.Collapsed;
                ProgressPercentageText.Visibility = Visibility.Collapsed;
                _cts?.Dispose();
                _cts = null;
                AppendLog("Processus de génération terminé ou annulé.");
            }
        }

        /// <summary>
        /// Gère le clic sur le bouton d'arrêt de la génération.
        /// </summary>
        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
                AppendLog("🛑 Annulation demandée au service.");
            }
        }

        /// <summary>
        /// Efface le contenu de la TextBox du chemin du fichier source.
        /// </summary>
        private void ClearSourceFileButton_Click(object sender, RoutedEventArgs e)
        {
            TxtSourceFilePath.Text = string.Empty;
        }

        /// <summary>
        /// Efface le contenu de la TextBox du chemin du fichier modèle.
        /// </summary>
        private void ClearTemplateFileButton_Click(object sender, RoutedEventArgs e)
        {
            TxtTemplateFilePath.Text = string.Empty;
        }

        /// <summary>
        /// Efface le contenu de la TextBox du chemin du dossier de sortie.
        /// </summary>
        private void ClearOutputDirButton_Click(object sender, RoutedEventArgs e)
        {
            TxtOutputDir.Text = string.Empty;
        }

        /// <summary>
        /// Active ou désactive les éléments de l'UI en fonction de l'état d'une opération.
        /// </summary>
        /// <param name="enabled">True pour activer les contrôles, False pour les désactiver.</param>
        private void SetUiEnabledState(bool enabled)
        {
            // Champs de saisie
            TxtSourceFilePath.IsEnabled = enabled;
            TxtTemplateFilePath.IsEnabled = enabled;
            TxtOutputDir.IsEnabled = enabled;
            TxtStartIndex.IsEnabled = enabled;
            TxtCount.IsEnabled = enabled;

            // Boutons de sélection de fichiers/dossiers
            BtnSelectSourceFile.IsEnabled = enabled;
            BtnSelectTemplateFile.IsEnabled = enabled;
            BtnSelectOutputDir.IsEnabled = enabled;

            // Boutons de suppression (croix)
            ClearSourceFileButton.IsEnabled = enabled;
            ClearTemplateFileButton.IsEnabled = enabled;
            ClearOutputDirButton.IsEnabled = enabled;

            // Boutons d'action principaux
            BtnGenerate.IsEnabled = enabled; // Le bouton Générer est activé si enabled est vrai
            BtnStop.IsEnabled = !enabled;    // Le bouton Stop est activé si enabled est faux (quand la génération est en cours)
        }
    }
}
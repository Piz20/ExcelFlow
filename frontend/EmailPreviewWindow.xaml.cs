using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.AspNetCore.SignalR.Client;
using ExcelFlow.Utilities;
using ExcelFlow.Services;
using ExcelFlow.Models;
using WpfMsgBox = System.Windows.MessageBox;
using wpfCheckBox = System.Windows.Controls.CheckBox;

namespace ExcelFlow
{
    public partial class EmailPreviewWindow : Window
    {
        private readonly HubConnection _hubConnection;
        private CancellationTokenSource? _cts;

        public class EmailToSendViewModel : INotifyPropertyChanged
        {
            public EmailToSend Email { get; }

            private bool _isSelected = true;
            public bool IsSelected
            {
                get => _isSelected;
                set
                {
                    if (_isSelected != value)
                    {
                        _isSelected = value;
                        OnPropertyChanged(nameof(IsSelected));
                    }
                }
            }

            private bool _isSending;
            public bool IsSending
            {
                get => _isSending;
                set
                {
                    if (_isSending != value)
                    {
                        _isSending = value;
                        OnPropertyChanged(nameof(IsSending));
                    }
                }
            }

            private bool _isSuccess;
            public bool IsSuccess
            {
                get => _isSuccess;
                set
                {
                    if (_isSuccess != value)
                    {
                        _isSuccess = value;
                        OnPropertyChanged(nameof(IsSuccess));
                    }
                }
            }

            private bool _isFailure;
            public bool IsFailure
            {
                get => _isFailure;
                set
                {
                    if (_isFailure != value)
                    {
                        _isFailure = value;
                        OnPropertyChanged(nameof(IsFailure));
                    }
                }
            }

            private string _selectedPartner;
            public string SelectedPartner
            {
                get => _selectedPartner;
                set
                {
                    if (_selectedPartner != value)
                    {
                        _selectedPartner = value;
                        OnPropertyChanged(nameof(SelectedPartner));
                    }
                }
            }

            private string _selectedAttachment;
            public string SelectedAttachment
            {
                get => _selectedAttachment;
                set
                {
                    if (_selectedAttachment != value)
                    {
                        _selectedAttachment = value;
                        OnPropertyChanged(nameof(SelectedAttachment));
                    }
                }
            }

            private string _selectedRecipient;
            public string SelectedRecipient
            {
                get => _selectedRecipient;
                set
                {
                    if (_selectedRecipient != value)
                    {
                        _selectedRecipient = value;
                        OnPropertyChanged(nameof(SelectedRecipient));
                    }
                }
            }


            private bool _isReadOnly;
            public bool IsReadOnly
            {
                get => _isReadOnly;
                set
                {
                    if (_isReadOnly != value)
                    {
                        _isReadOnly = value;
                        OnPropertyChanged(nameof(IsReadOnly));
                    }
                }
            }

            public List<string> PartnerNameList => new List<string> { Email.PartnerName ?? "(Partenaire inconnu)" };
            public List<string> AttachmentFilePaths => Email.AttachmentFilePaths?.Count > 0
                ? Email.AttachmentFilePaths
                : new List<string> { "(Aucun fichier)" };
            public List<string> ToRecipients => Email.ToRecipients ?? new List<string>();

            public EmailToSendViewModel(EmailToSend email)
            {
                Email = email;
                _selectedPartner = PartnerNameList.FirstOrDefault() ?? string.Empty;
                _selectedAttachment = AttachmentFilePaths.FirstOrDefault() ?? string.Empty;
                _selectedRecipient = ToRecipients.FirstOrDefault() ?? string.Empty;

            }

            public event PropertyChangedEventHandler? PropertyChanged;
            protected void OnPropertyChanged(string propertyName) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private readonly ObservableCollection<EmailToSendViewModel> _emailViewModels = new();

        private readonly SendEmailService _sendEmailService;

        public EmailPreviewWindow(List<EmailToSend> preparedEmails)
        {
            InitializeComponent();

            _hubConnection = new HubConnectionBuilder()
               .WithUrl($"http://localhost:{AppConstants.port}/partnerFileHub")
               .WithAutomaticReconnect()
               .Build();

            _sendEmailService = new SendEmailService($"http://localhost:{AppConstants.port}");

            foreach (var email in preparedEmails)
            {
                var vm = new EmailToSendViewModel(email);
                vm.PropertyChanged += EmailVM_PropertyChanged;
                _emailViewModels.Add(vm);
            }

            EmailsDataGrid.ItemsSource = _emailViewModels;
            UpdateSelectedEmailsCount();

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

            this.Loaded += EmailPreviewWindow_Loaded;
            this.Closing += EmailPreviewWindow_Closing;

        }

        private async void EmailPreviewWindow_Loaded(object sender, RoutedEventArgs e)
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

        private void EmailVM_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(EmailToSendViewModel.IsSelected))
            {
                UpdateSelectedEmailsCount();
            }
        }

        private void UpdateSelectedEmailsCount()
        {
            if (SelectedEmailsCountTextBlock == null)
                return;

            int selectedCount = _emailViewModels.Count(vm => vm.IsSelected);
            SelectedEmailsCountTextBlock.Text = $"Partenaires sélectionnés : {selectedCount}";

            if (SelectAllCheckBox == null)
                return;

            if (_emailViewModels.All(vm => vm.IsSelected))
            {
                SelectAllCheckBox.IsChecked = true;
                SelectAllCheckBox.Content = "Ne rien sélectionner";
            }
            else if (_emailViewModels.All(vm => !vm.IsSelected))
            {
                SelectAllCheckBox.IsChecked = false;
                SelectAllCheckBox.Content = "Tout sélectionner";
            }
            else
            {
                SelectAllCheckBox.IsChecked = null; // état indéterminé
                SelectAllCheckBox.Content = "Tout sélectionner";
            }

            if (SendSelectedButton != null)
            {
                SendSelectedButton.IsEnabled = selectedCount > 0;
            }
        }

        private void SelectAllCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is wpfCheckBox cb)
            {
                bool isChecked = cb.IsChecked ?? false;

                foreach (var emailVM in _emailViewModels)
                {
                    emailVM.IsSelected = isChecked;
                }

                UpdateSelectedEmailsCount();
            }
        }

        private async void SendSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            var toSend = _emailViewModels.Where(vm => vm.IsSelected).ToList();

            if (!toSend.Any())
            {
                WpfMsgBox.Show("Veuillez sélectionner au moins un email à envoyer.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _cts = new CancellationTokenSource();

            EmailsDataGrid.IsReadOnly = true; // Protection supplémentaire
            SetInteractionLock(true);         // <- Désactive les checkbox

            SelectAllCheckBox.IsEnabled = false;
            SendSelectedButton.IsEnabled = false;
            StopButton.IsEnabled = true;

            try
            {
                for (int i = 0; i < toSend.Count; i++)
                {
                    var vm = toSend[i];

                    AppendLog($"\n\n📤 Envoi de l'email {i + 1}/{toSend.Count} à : {string.Join(", ", vm.Email.ToRecipients)}");

                    vm.IsSending = true;
                    vm.IsSuccess = false;
                    vm.IsFailure = false;

                    try
                    {
                        var singleEmailList = new List<EmailToSend> { vm.Email };
                        var results = await _sendEmailService.SendPreparedEmailsAsync(singleEmailList, _cts.Token);
                        var result = results.FirstOrDefault();

                        if (result != null && result.Success)
                        {
                            vm.IsSuccess = true;
                            vm.IsFailure = false;
                        }
                        else
                        {
                            vm.IsSuccess = false;
                            vm.IsFailure = true;
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        vm.IsSending = false;
                        throw;
                    }
                    catch (Exception)
                    {
                        vm.IsSuccess = false;
                        vm.IsFailure = true;
                    }
                    finally
                    {
                        vm.IsSending = false;
                    }
                }

                foreach (var vm in toSend)
                {
                    if (vm.IsSuccess)
                    {
                        vm.IsSelected = false;
                    }
                }

                WpfMsgBox.Show("Envoi terminé.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                WpfMsgBox.Show("L'envoi a été annulé.", "Annulé", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                _cts.Dispose();
                _cts = null;

                SetInteractionLock(false);            // <- Réactive les checkbox
                EmailsDataGrid.IsReadOnly = false;

                SelectAllCheckBox.IsEnabled = true;
                SendSelectedButton.IsEnabled = true;
                StopButton.IsEnabled = false;

                UpdateSelectedEmailsCount();
            }
        }

        private void SetInteractionLock(bool isLocked)
        {
            foreach (var vm in _emailViewModels)
            {
                vm.IsReadOnly = isLocked;
            }
        }


        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;

            while (source != null && !(source is System.Windows.Controls.TextBox) &&
        !(source is System.Windows.Controls.ComboBox) &&
        !(source is System.Windows.Controls.PasswordBox))
            {
                source = VisualTreeHelper.GetParent(source);
            }

            if (source == null)
            {
                Keyboard.ClearFocus();
                FocusManager.SetFocusedElement(this, this);
            }
        }


        private void EmailPreviewWindow_Closing(object? sender, CancelEventArgs e)
        {
            bool sendingInProgress = _emailViewModels.Any(vm => vm.IsSending);

            if (sendingInProgress)
            {
                WpfMsgBox.Show("Veuillez attendre que tous les envois d’emails soient terminés avant de fermer la fenêtre.",
                    "Envoi en cours",
                    MessageBoxButton.OK,
                    MessageBoxImage.Exclamation);
                e.Cancel = true;
                return;
            }

            var result = WpfMsgBox.Show(
                " URGENT : Êtes-vous sûr d’avoir envoyé tous les mails dont vous avez besoin ?",
                "Confirmation de fermeture",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning
            );

            if (result != MessageBoxResult.Yes)
            {
                e.Cancel = true;
            }
        }

        private void ClearLogs_Click(object sender, RoutedEventArgs e)
        {
            TxtLogs.Clear();
        }

        private void AppendLog(string message)
        {
            TxtLogs.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
            TxtLogs.ScrollToEnd();
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
            }
        }
    }
}

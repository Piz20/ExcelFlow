using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelFlow.Services;
using ExcelFlow.Models;
using WpfMsgBox = System.Windows.MessageBox;
using wpfCheckBox = System.Windows.Controls.CheckBox;

namespace ExcelFlow
{
    public partial class EmailPreviewWindow : Window
    {
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
                get => _selectedAttachment; // Corrigé pour retourner _selectedAttachment
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

            public List<string> PartnerNameList => new List<string> { Email.PartnerName ?? "(Partenaire inconnu)" };
            public List<string> AttachmentFilePaths => Email.AttachmentFilePaths?.Count > 0
                ? Email.AttachmentFilePaths
                : new List<string> { "(Aucun fichier)" };
            public List<string> ToRecipients => Email.ToRecipients ?? new List<string>();

            public EmailToSendViewModel(EmailToSend email)
            {
                Email = email;
                _selectedPartner = PartnerNameList.FirstOrDefault();
                _selectedAttachment = AttachmentFilePaths.FirstOrDefault();
                _selectedRecipient = ToRecipients.FirstOrDefault();
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

            // Initialise le service avec l'URL de ton backend (ajuste si nécessaire)
            _sendEmailService = new SendEmailService("https://localhost:7274");

            foreach (var email in preparedEmails)
            {
                var vm = new EmailToSendViewModel(email);
                vm.PropertyChanged += EmailVM_PropertyChanged;
                _emailViewModels.Add(vm);
            }

            EmailsDataGrid.ItemsSource = _emailViewModels;
            UpdateSelectedEmailsCount();
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
            SelectedEmailsCountTextBlock.Text = $"Emails sélectionnés : {selectedCount}";

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

            // Désactiver UI au début
            EmailsDataGrid.IsEnabled = false;
            SelectAllCheckBox.IsEnabled = false;
            SendSelectedButton.IsEnabled = false;
            StopButton.IsEnabled = true;

            try
            {
                foreach (var vm in toSend)
                {
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

                            // Logging dans la console
                            Console.WriteLine($"✔ Email envoyé à {result.To}");
                        }
                        else
                        {
                            vm.IsSuccess = false;
                            vm.IsFailure = true;

                            // Logging dans la console avec erreur
                            Console.WriteLine($"✘ Échec de l'envoi à {result?.To ?? "inconnu"} : {result?.ErrorMessage ?? "Erreur inconnue"}");
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        vm.IsSending = false;
                        vm.IsSuccess = false;
                        vm.IsFailure = false;
                        Console.WriteLine("⏹️ Envoi annulé par l'utilisateur.");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        vm.IsSuccess = false;
                        vm.IsFailure = true;
                        Console.WriteLine($"✘ Exception lors de l'envoi : {ex.Message}");
                    }
                    finally
                    {
                        vm.IsSending = false;
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

                EmailsDataGrid.IsEnabled = true;
                SelectAllCheckBox.IsEnabled = true;
                SendSelectedButton.IsEnabled = true;
                StopButton.IsEnabled = false;

                UpdateSelectedEmailsCount();
            }
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
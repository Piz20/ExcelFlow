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

        public EmailPreviewWindow(List<EmailToSend> preparedEmails)
        {
            InitializeComponent();

            // Initialiser explicitement la couleur du bouton Stop à rouge pâle
            StopButton.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FF9999"));

            foreach (var email in preparedEmails)
            {
                var vm = new EmailToSendViewModel(email);
                vm.PropertyChanged += EmailVM_PropertyChanged; // pour suivre la sélection
                _emailViewModels.Add(vm);
            }

            EmailsDataGrid.ItemsSource = _emailViewModels;

            UpdateSelectAllCheckBox();
        }

        private void EmailVM_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(EmailToSendViewModel.IsSelected))
            {
                UpdateSelectAllCheckBox();
                UpdateSendButtonEnabled();
            }
        }

        private void UpdateSelectAllCheckBox()
        {
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
        }

        private void UpdateSendButtonEnabled()
        {
            if (SendSelectedButton == null)
                return;

            SendSelectedButton.IsEnabled = _emailViewModels.Any(vm => vm.IsSelected);
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

                cb.Content = isChecked ? "Ne rien sélectionner" : "Tout sélectionner";

                UpdateSendButtonEnabled();
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

            // Afficher le nombre d'emails sélectionnés
            WpfMsgBox.Show($"Nombre d'emails à envoyer : {toSend.Count}", "Information", MessageBoxButton.OK, MessageBoxImage.Information);

            _cts = new CancellationTokenSource();

            // Nettoyer la colonne Statut pour les emails sélectionnés
            foreach (var vm in toSend)
            {
                vm.IsSending = false;
                vm.IsSuccess = false;
                vm.IsFailure = false;
            }

            // Bloquer les contrôles sauf le bouton Stop
            EmailsDataGrid.IsEnabled = false;
            SelectAllCheckBox.IsEnabled = false;
            SendSelectedButton.IsEnabled = false;
            StopButton.IsEnabled = true;
            StopButton.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F44336"));

            try
            {
                foreach (var vm in toSend)
                {
                    _cts.Token.ThrowIfCancellationRequested();

                    vm.IsSending = true;

                    await Task.Delay(1000, _cts.Token); // Simulation, à remplacer par l'envoi réel

                    bool success = true; // succès simulé

                    vm.IsSending = false;
                    vm.IsSuccess = success;
                    vm.IsFailure = !success;
                }
            }
            catch (OperationCanceledException)
            {
                WpfMsgBox.Show("L'envoi a été annulé.", "Annulé", MessageBoxButton.OK, MessageBoxImage.Information);

                foreach (var vm in _emailViewModels)
                {
                    if (vm.IsSending)
                    {
                        vm.IsSending = false;
                        vm.IsSuccess = false;
                        vm.IsFailure = false;
                    }
                }
            }
            finally
            {
                _cts.Dispose();
                _cts = null;

                // Réactiver les contrôles et rétablir la couleur du bouton Stop
                EmailsDataGrid.IsEnabled = true;
                SelectAllCheckBox.IsEnabled = true;
                StopButton.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FF9999"));
                UpdateSendButtonEnabled();
                StopButton.IsEnabled = false;
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
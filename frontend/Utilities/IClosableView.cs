
using WpsMsgBoxImage = System.Windows.MessageBoxImage;
namespace ExcelFlow.Utilities

{
    // 1. Interface pour les vues personnalisables lors de la fermeture
    public interface IClosableView
    {
        bool IsOperationInProgress { get; }
        (string Message, string Title, WpsMsgBoxImage Icon) GetClosingConfirmation();
    }
}
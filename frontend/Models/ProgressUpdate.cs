// Fichier : Backend/Models/ProgressUpdate.cs
namespace ExcelFlow.Models
{
    public class ProgressUpdate
    {
        public int Current { get; set; }
        public int Total { get; set; }
        public int Percentage { get; set; }
        public string? Message { get; set; } // Utilisez string? pour la nullabilité
    }
}
using System.ComponentModel.DataAnnotations;

namespace ExcelFlow.Models;

public class ShowStructureRequest
{
    [Required(ErrorMessage = "Le chemin du fichier est requis.")]
    public string FilePath { get; set; } = string.Empty;

}

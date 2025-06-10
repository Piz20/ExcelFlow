using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using System.ComponentModel.DataAnnotations;

namespace ExcelFlow.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ShowStructureController : ControllerBase
    {
        [HttpGet("afficher-structure")]
        public IActionResult AfficherStructureDepuisQuery([FromQuery][Required] string filePath)
        {
            if (!System.IO.File.Exists(filePath))
                return NotFound($"❌ Fichier introuvable : {filePath}");

            try
            {
                using var workbook = new XLWorkbook(filePath);

                foreach (var worksheet in workbook.Worksheets)
                {
                    Console.WriteLine($"\n=== Structure de la feuille : {worksheet.Name} ===");
                    ExcelUtils.AfficherStructureColonnes(workbook, worksheet.Name);
                }

                return Ok("✅ Structure de toutes les feuilles affichée dans la console.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"❌ Erreur lors de la lecture du fichier Excel : {ex.Message}");
            }
        }
    }
}

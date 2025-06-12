using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using ClosedXML.Excel;
using ExcelFlow.Services;
using ExcelFlow.Models;

using ExcelFlow.Hubs;

namespace ExcelFlow.Controllers
{
    [ApiController]
    [Route("api/partnerfile")]
    public class PartnerFileController : ControllerBase
    {
        private readonly PartnerFileGenerator _generator;
        private readonly IHubContext<PartnerFileHub> _hubContext;

        public PartnerFileController(PartnerFileGenerator generator, IHubContext<PartnerFileHub> hubContext)
        {
            _generator = generator;
            _hubContext = hubContext;
        }

        [HttpPost("generate")]
        public async Task<IActionResult> GenerateFiles([FromBody] PartnerFileRequest request, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(request.FilePath) || !System.IO.File.Exists(request.FilePath))
                return BadRequest("❌ Fichier source introuvable ou manquant.");

            if (string.IsNullOrEmpty(request.TemplatePath) || !System.IO.File.Exists(request.TemplatePath))
                return BadRequest("❌ Fichier template introuvable ou manquant.");

            if (string.IsNullOrEmpty(request.OutputDir) || !Directory.Exists(request.OutputDir))
                return BadRequest("❌ Dossier de sortie inexistant ou manquant.");

            try
            {
                await _hubContext.Clients.All.SendAsync("ReceiveLog", "Début de la génération des fichiers...");

                using var stream = System.IO.File.Open(request.FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workbook = new XLWorkbook(stream);
// Cherche dans toutes les feuilles une qui a un nom égal au SheetName trimé
var worksheet = workbook.Worksheets
    .FirstOrDefault(ws => ws.Name.Trim().Equals(request.SheetName?.Trim(), StringComparison.OrdinalIgnoreCase));

if (worksheet == null)
{
    await _hubContext.Clients.All.SendAsync("ReceiveLog", $"❌ La feuille '{request.SheetName}' est introuvable.");
    return BadRequest($"La feuille '{request.SheetName}' est introuvable dans le fichier Excel.");
}


                // On passe le token d'annulation
                await _generator.GeneratePartnerFilesAsync(
                    worksheet,
                    request.TemplatePath,
                    request.OutputDir,
                    request.StartIndex,
                    request.Count,
                    cancellationToken
                );

                await _hubContext.Clients.All.SendAsync("ReceiveLog", "✅ Fichiers générés avec succès !");
                return Ok("✅ Fichiers générés avec succès !");
            }
            catch (OperationCanceledException)
            {
                await _hubContext.Clients.All.SendAsync("ReceiveLog", "⚠️ Génération annulée par l'utilisateur.");
                return StatusCode(499, "⚠️ Génération annulée par l'utilisateur."); // 499 = Client Closed Request (non officiel)
            }
            catch (Exception ex)
            {
                await _hubContext.Clients.All.SendAsync("ReceiveLog", $"❌ Erreur lors de la génération : {ex.Message}");
                return StatusCode(500, $"❌ Erreur lors de la génération : {ex.Message}");
            }
        }
      

    }
}

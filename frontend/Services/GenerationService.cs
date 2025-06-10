using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading;
using System.Threading.Tasks;

public class GenerationService
{
    private readonly HttpClient _client;

    public GenerationService(string baseUrl)
    {
        _client = new HttpClient { BaseAddress = new Uri(baseUrl) };
    }

    public async Task<string> GenerateAsync(GenerationRequest request, CancellationToken cancellationToken)
    {
        try
        {
            var response = await _client.PostAsJsonAsync("/api/partnerfile/generate", request, cancellationToken);
            var result = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception(result);

            return result; // ✅ Fichiers générés avec succès !
        }
        catch (OperationCanceledException)
        {
            // Propager l'annulation vers l'appelant pour gestion spécifique
            throw;
        }
        catch (Exception ex)
        {
            return $"❌ Erreur lors de la génération : {ex.Message}";
        }
    }

}

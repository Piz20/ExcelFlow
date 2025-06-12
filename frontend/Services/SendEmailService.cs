using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json; // Used for JsonSerializer
using System.Net.Http.Json; // Used for PostAsJsonAsync

// Make sure ExcelFlow.Models is a shared project referenced by your WPF app
// or ensure these DTOs are defined within your WPF project if not shared.
using ExcelFlow.Models;

namespace ExcelFlow.Services // Correction du namespace
{
    public class SendEmailService
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;

        /// <summary>
        /// Initialise une nouvelle instance du service d'envoi d'emails.
        /// </summary>
        /// <param name="baseUrl">L'URL de base de votre API backend ASP.NET Core (ex: "https://localhost:7274").</param>
        public SendEmailService(string baseUrl)
        {
            _baseUrl = baseUrl;
            _httpClient = new HttpClient();
            // Vous pouvez configurer d'autres propriétés HttpClient ici si nécessaire,
            // comme BaseAddress pour simplifier les appels futurs.
            // _httpClient.BaseAddress = new Uri(baseUrl);
        }

        /// <summary>
        /// Démarre le processus d'envoi d'emails en appelant une API sur le backend.
        /// </summary>
        /// <param name="request">Les données de la requête d'envoi d'emails.</param>
        /// <param name="cancellationToken">Token d'annulation pour arrêter l'opération.</param>
        /// <returns>Un message de résultat de l'opération (succès ou erreur).</returns>
        public async Task<string> StartEmailSendingAsync(EmailSendRequest request, CancellationToken cancellationToken)
        {
            try
            {
                // Construit l'URL complète de l'API. Assurez-vous que votre backend a un endpoint correspondant.
                // Exemple d'endpoint : /api/email/start-sending
                var response = await _httpClient.PostAsJsonAsync($"{_baseUrl}/api/partneremailsender/send", request, cancellationToken);

                // Vérifie si la requête HTTP a réussi (statut 2xx)
                response.EnsureSuccessStatusCode();

                // Lit le contenu de la réponse du serveur
                var responseString = await response.Content.ReadAsStringAsync();
                return responseString; // Retourne le message de succès ou d'information du backend
            }
            catch (HttpRequestException ex)
            {
                // Gère les erreurs de requête HTTP (problèmes de réseau, serveur non joignable, erreur 4xx/5xx du serveur)
                return $"❌ Erreur réseau ou du serveur: {ex.Message}";
            }
            catch (OperationCanceledException)
            {
                // Gère l'annulation de l'opération
                return "Opération d'envoi d'emails annulée.";
            }
            catch (Exception ex)
            {
                // Gère toute autre exception inattendue
                return $"❌ Erreur inattendue lors du démarrage de l'envoi d'emails: {ex.Message}";
            }
        }
    }
}
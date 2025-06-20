using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelFlow.Models;
using System.Text.Json;
namespace ExcelFlow.Services
{
    public class SendEmailService
    {
        private readonly HttpClient _httpClient;

        public SendEmailService(string baseUrl)
        {
            // Initialise HttpClient avec BaseAddress pour éviter de répéter l'URL complète à chaque appel
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri(baseUrl)
            };
        }

        /// <summary>
        /// Prépare les emails à envoyer en appelant l'API /prepare.
        /// </summary>
        public async Task<List<EmailToSend>?> PrepareEmailsAsync(PrepareEmailRequest request, CancellationToken cancellationToken)
        {
            try
            {
                // Sérialisation en JSON pour inspection/debug
                var jsonRequest = JsonSerializer.Serialize(request, new JsonSerializerOptions { WriteIndented = true });
                Console.WriteLine("Request JSON envoyé :");
                Console.WriteLine(jsonRequest);

                var response = await _httpClient.PostAsJsonAsync("api/partneremailsender/prepare", request, cancellationToken);
                response.EnsureSuccessStatusCode();

                var emails = await response.Content.ReadFromJsonAsync<List<EmailToSend>>(cancellationToken: cancellationToken);
                return emails;
            }
            catch (HttpRequestException ex)
            {
                throw new ApplicationException($"❌ Erreur HTTP lors de la préparation : {ex.Message}");
            }
        }


        /// <summary>
        /// Envoie la liste d'emails préparés en appelant l'API /send.
        /// </summary>
        public async Task<string> SendPreparedEmailsAsync(List<EmailToSend> preparedEmails, CancellationToken cancellationToken)
        {
            try
            {
                var response = await _httpClient.PostAsJsonAsync("api/partneremailsender/send", preparedEmails, cancellationToken);
                response.EnsureSuccessStatusCode();

                return await response.Content.ReadAsStringAsync();
            }
            catch (HttpRequestException ex)
            {
                return $"❌ Erreur réseau ou serveur lors de l’envoi des emails préparés : {ex.Message}";
            }
            catch (OperationCanceledException)
            {
                return "⏹️ Envoi annulé.";
            }
            catch (Exception ex)
            {
                return $"❌ Erreur inattendue : {ex.Message}";
            }
        }

        /// <summary>
        /// Méthode regroupée qui prépare puis envoie les emails.
        /// </summary>
        public async Task<string> StartEmailSendingAsync(PrepareEmailRequest prepareRequest, CancellationToken cancellationToken)
        {
            try
            {
                var preparedEmails = await PrepareEmailsAsync(prepareRequest, cancellationToken);

                if (preparedEmails == null || preparedEmails.Count == 0)
                    return "❌ Aucun email préparé.";

                var sendResult = await SendPreparedEmailsAsync(preparedEmails, cancellationToken);
                return sendResult;
            }
            catch (Exception ex)
            {
                return $"❌ Erreur inattendue lors du processus d'envoi : {ex.Message}";
            }
        }
    }
}

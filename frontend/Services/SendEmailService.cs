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
        public async Task<List<EmailSendResult>> SendPreparedEmailsAsync(List<EmailToSend> preparedEmails, CancellationToken cancellationToken)
        {
            try
            {
                var response = await _httpClient.PostAsJsonAsync("api/partneremailsender/send", preparedEmails, cancellationToken);
                response.EnsureSuccessStatusCode();

                var results = await response.Content.ReadFromJsonAsync<List<EmailSendResult>>(cancellationToken: cancellationToken);
                return results ?? new List<EmailSendResult>();
            }
            catch (HttpRequestException ex)
            {
                // Ici tu peux choisir de retourner une liste vide ou une liste avec un seul résultat d’erreur, selon ta logique
                return new List<EmailSendResult>
        {
            new EmailSendResult { To = "", Success = false, ErrorMessage = $"❌ Erreur réseau ou serveur : {ex.Message}" }
        };
            }
            catch (OperationCanceledException)
            {
                return new List<EmailSendResult>
        {
            new EmailSendResult { To = "", Success = false, ErrorMessage = "⏹️ Envoi annulé." }
        };
            }
            catch (Exception ex)
            {
                return new List<EmailSendResult>
        {
            new EmailSendResult { To = "", Success = false, ErrorMessage = $"❌ Erreur inattendue : {ex.Message}" }
        };
            }
        }

      
    }
}

using ExcelFlow.Hubs;
using ExcelFlow.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting; // Souvent implicite avec .NET 6+, mais bon de le garder
using Microsoft.Extensions.Configuration; // Nécessaire si vous l'injectez directement, mais builder.Configuration est suffisant.

var builder = WebApplication.CreateBuilder(args);

// --- Configuration des Services (Add services to the container.) ---
// C'est ici que tous les services sont enregistrés dans le conteneur d'injection de dépendances.

// Services spécifiques à l'application
builder.Services.AddScoped<SendEmail>();
builder.Services.AddScoped<PartnerEmailSender>();
builder.Services.AddScoped<PartnerFileGenerator>(); // Votre service de génération de fichiers

// Services ASP.NET Core
builder.Services.AddControllers(); // Ajoute la prise en charge des contrôleurs MVC pour les APIs
builder.Services.AddEndpointsApiExplorer(); // Nécessaire pour Swagger/OpenAPI

// Services SignalR
builder.Services.AddSignalR(); // Ajoute les services nécessaires pour SignalR

// Configuration des URLs (principalement défini via launchSettings.json, mais peut être surchargé ici)
// Garder cette ligne si vous souhaitez forcer les URLs ou pour des environnements spécifiques.
builder.WebHost.UseUrls("http://localhost:5297", "https://localhost:7274");


var app = builder.Build();

// --- Configuration du Pipeline de Requêtes HTTP (Configure the HTTP request pipeline.) ---
// L'ordre des middlewares ici est CRUCIAL !


// Redirection HTTP vers HTTPS pour la sécurité
app.UseHttpsRedirection();

// Active le service de fichiers statiques (pour servir des fichiers comme HTML, CSS, JS depuis wwwroot)
app.UseStaticFiles();

// Active le routage pour faire correspondre les requêtes aux endpoints
app.UseRouting();

// Configure l'autorisation (doit venir après UseRouting et UseAuthentication si présent)
app.UseAuthorization();
// Si vous avez de l'authentification, décommentez et placez 'app.UseAuthentication();' ICI, AVANT UseAuthorization.

// --- Mapping des endpoints ---
// Associe les requêtes entrantes aux contrôleurs et aux hubs SignalR.

// Mappe les contrôleurs d'API
app.MapControllers();

// Mappe le hub SignalR à son chemin d'accès
app.MapHub<PartnerFileHub>("/partnerFileHub");

// --- Démarrage de l'application ---
app.Run();
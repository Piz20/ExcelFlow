using ExcelFlow.Hubs;
using ExcelFlow.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration; // Not strictly needed if using builder.Configuration directly
using ExcelFlow.Utilities; // Pour EmailDataExtractor
using ExcelFlow.Models; // Pour EmailData
using Microsoft.AspNetCore.SignalR; // Pour IHubContext<PartnerFileHub>
// Note: EmailData is now in ExcelFlow.Models, so we don't need a direct using for it here.
// The services that use it (EmailDataExtractor, EmailContentBuilder) will have their own using statements.

// No direct using for IHubContext<PartnerFileHub> needed here, as it's used within AddScoped lambda

var builder = WebApplication.CreateBuilder(args);

// --- Configuration des Services (Add services to the container.) ---
// C'est ici que tous les services sont enregistrés dans le conteneur d'injection de dépendances.

// Services spécifiques à l'application

// SendEmail Service
builder.Services.AddScoped<SendEmail>();

// PartnerExcelReader Service (Assumed to be simple instantiation, no complex dependencies shown)
builder.Services.AddScoped<PartnerExcelReader>();

// EmailContentBuilder Service (Assumed to be simple instantiation, no complex dependencies shown)
builder.Services.AddScoped<EmailContentBuilder>();

// EmailDataExtractor has been modified to take an IHubContext for logging purposes.
builder.Services.AddScoped<EmailDataExtractor>(provider =>
{
    // Get the IHubContext via the service provider
    var hubContext = provider.GetRequiredService<IHubContext<PartnerFileHub>>();
    // Create an instance of EmailDataExtractor, passing the hubContext
    return new EmailDataExtractor(hubContext);
});


builder.Services.AddScoped<IPartnerEmailSender, PartnerEmailSender>();


// PartnerEmailSender depends on SendEmail, EmailDataExtractor, PartnerExcelReader, and EmailContentBuilder.
// Ensure all its dependencies are registered *before* PartnerEmailSender itself.
builder.Services.AddScoped<PartnerEmailSender>();

// Your file generation service
builder.Services.AddScoped<PartnerFileGenerator>();

// Services ASP.NET Core
builder.Services.AddControllers(); // Adds MVC controllers support for APIs
builder.Services.AddEndpointsApiExplorer(); // Needed for Swagger/OpenAPI

// SignalR Services
builder.Services.AddSignalR(); // Adds the necessary services for SignalR

// URL Configuration (mainly defined via launchSettings.json, but can be overridden here)
builder.WebHost.UseUrls("http://localhost:5297", "https://localhost:7274");


var app = builder.Build();

// --- Configuration du Pipeline de Requêtes HTTP (Configure the HTTP request pipeline.) ---
// The order of middleware here is CRUCIAL!

// HTTP to HTTPS redirection for security
app.UseHttpsRedirection();

// Enables static files service (to serve files like HTML, CSS, JS from wwwroot)
app.UseStaticFiles();

// Enables routing to match requests to endpoints
app.UseRouting();

// Configures authorization (must come after UseRouting and UseAuthentication if present)
app.UseAuthorization();
// If you have authentication, uncomment and place 'app.UseAuthentication();' HERE, BEFORE UseAuthorization.

// --- Mapping endpoints ---
// Maps incoming requests to controllers and SignalR hubs.

// Maps API controllers
app.MapControllers();

// Maps the SignalR hub to its access path
app.MapHub<PartnerFileHub>("/partnerFileHub");

// --- Application startup ---
app.Run();
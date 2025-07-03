using ExcelFlow.Hubs;
using ExcelFlow.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using ExcelFlow.Utilities;
using ExcelFlow.Models;
using Microsoft.AspNetCore.SignalR;

var builder = WebApplication.CreateBuilder(args);

// --- Configuration des Services ---
// Application services
builder.Services.AddScoped<SendEmail>();
builder.Services.AddScoped<PartnerExcelReader>();
builder.Services.AddScoped<EmailContentBuilder>();

builder.Services.AddScoped<EmailDataExtractor>(provider =>
{
    var hubContext = provider.GetRequiredService<IHubContext<PartnerFileHub>>();
    return new EmailDataExtractor(hubContext);
});

builder.Services.AddScoped<IPartnerEmailSender, PartnerEmailSender>();
builder.Services.AddScoped<PartnerEmailSender>();
builder.Services.AddScoped<PartnerFileGenerator>();

// ASP.NET Core Services
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();

// SignalR
builder.Services.AddSignalR();

// ✅ Configuration de Kestrel pour n’utiliser que HTTP
builder.WebHost.ConfigureKestrel(options =>
{
    options.ListenAnyIP(5297); // Port HTTP seulement
});

// ❌ Supprimer les bindings HTTPS
// ❌ Supprimer les redirections HTTPS automatiques
// builder.WebHost.UseUrls("http://localhost:5297", "https://localhost:7274"); // supprimé

var app = builder.Build();

// --- Configuration du pipeline HTTP ---
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

// Endpoints
app.MapControllers();
app.MapHub<PartnerFileHub>("/partnerFileHub");

app.Run();

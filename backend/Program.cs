using ExcelFlow.Hubs;
using ExcelFlow.Services;
using Microsoft.AspNetCore.Builder; // Make sure this is present
using Microsoft.Extensions.DependencyInjection; // Make sure this is present
using Microsoft.Extensions.Hosting; // Make sure this is present
using Microsoft.Extensions.Configuration; // Make sure this is present

var builder = WebApplication.CreateBuilder(args);

// 🛠️ Ajout des services AVANT builder.Build()

// 1. Add IConfiguration (already loaded by builder.CreateApplicationBuilder)
// It's good practice to explicitly add it to the DI container if you're injecting it.
builder.Services.AddSingleton<IConfiguration>(builder.Configuration);

// 2. Register your SendEmail class for Dependency Injection
// Use AddTransient if SendEmail does not hold state between requests
// Use AddScoped if SendEmail needs to hold state within a single request
// Use AddSingleton if SendEmail should be a single instance for the entire app lifetime
// For an email sender, AddTransient or AddScoped are generally appropriate. AddScoped is a good default.
builder.Services.AddScoped<SendEmail>(); // <--- ADD/CONFIRM THIS LINE FOR SendEmail

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer(); // For Swagger/OpenAPI support, if you're using it
builder.Services.AddSignalR();
builder.Services.AddScoped<PartnerFileGenerator>(); // Your existing service registration

// Configure URLs (This is generally done via launchSettings.json,
// but setting it here also works for direct execution)
builder.WebHost.UseUrls("http://localhost:5297", "https://localhost:7274");

var app = builder.Build();

// Configure the HTTP request pipeline.

// Enable static files (like content in wwwroot)
app.UseStaticFiles();



app.UseRouting();

// Use authentication and authorization middleware (if you have any defined)
app.UseAuthorization();
// app.UseAuthentication(); // If you have authentication configured, add this before UseAuthorization


// Map endpoints for controllers and SignalR hubs
app.MapControllers();
app.MapHub<PartnerFileHub>("/partnerFileHub");

app.Run();
using ExcelFlow.Hubs;
using ExcelFlow.Services;


var builder = WebApplication.CreateBuilder(args);

// 🛠️ Ajout des services AVANT builder.Build()
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSignalR();
builder.Services.AddScoped<PartnerFileGenerator>();
builder.WebHost.UseUrls("http://localhost:5297", "https://localhost:7274");

var app = builder.Build();

app.UseStaticFiles(); // Permet d'exposer wwwroot


app.UseRouting();

app.UseAuthorization();

app.MapControllers();
app.MapHub<PartnerFileHub>("/partnerFileHub");

app.Run();
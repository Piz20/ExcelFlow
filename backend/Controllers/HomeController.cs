using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("/")] // ← rend disponible directement sur /
public class HomeController : ControllerBase
{
    [HttpGet]
    public string Get()
    {
        return "Hello from controller de Piz!";
    }
}

using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/info")]
public class InfoController : ControllerBase
{
    [HttpGet]
    public IActionResult Get()
    {
        var version = typeof(InfoController).Assembly.GetName().Version?.ToString(3)
                      ?? "0.0.0";
        return Ok(new
        {
            name    = "SaveAsPDF.Backend",
            version,
            author  = "Ofer Aharon",
            email   = "ofer@sw-eng.co.il"
        });
    }
}

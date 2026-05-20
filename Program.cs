
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

// Explicitly set the listening URL
builder.WebHost.UseUrls("http://0.0.0.0:3000");

// Enable CORS
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();

app.UseCors();

// Define AIA folder path relative to the project root
var aiaFolderPath = Path.Combine(app.Environment.ContentRootPath, "AIA");
if (!Directory.Exists(aiaFolderPath))
{
    Directory.CreateDirectory(aiaFolderPath);
    Console.WriteLine($"Created directory: {aiaFolderPath}");
}

// ✅ GET: Serve Excel template
app.MapGet("/api/get-assessment-template", async (HttpContext context) =>
{
    var filePath = Path.Combine(aiaFolderPath, "AI_Maturity_Assessment.xlsx");

    if (!File.Exists(filePath))
    {
        Console.WriteLine($"File not found: {filePath}");
        context.Response.StatusCode = 404;
        await context.Response.WriteAsync("AI_Maturity_Assessment.xlsx not found in the AIA folder.");
        return;
    }

    context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    await context.Response.SendFileAsync(filePath);
});

// ✅ POST: Save Excel results
app.MapPost("/api/save-assessment-results", async (HttpRequest request, HttpResponse response) =>
{
    var filename = request.Headers["x-filename"].FirstOrDefault() ?? "Results.xlsx";
    var filePath = Path.Combine(aiaFolderPath, filename);

    Console.WriteLine($"Received request to save: {filename}");

    using var memoryStream = new MemoryStream();
    await request.Body.CopyToAsync(memoryStream);

    if (memoryStream.Length == 0)
    {
        response.StatusCode = 400;
        await response.WriteAsync("No data received.");
        return;
    }

    try
    {
        await File.WriteAllBytesAsync(filePath, memoryStream.ToArray());

        Console.WriteLine($"Successfully saved updated workbook to {filePath}");

        response.ContentType = "application/json";
        await response.WriteAsJsonAsync(new
        {
            message = "File saved successfully",
            path = filePath
        });
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error saving file: {ex}");
        response.StatusCode = 500;
        await response.WriteAsync("Failed to save file.");
    }
});

app.Run();

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;

 // Initialize the web application builder to configure services and the web host.
 var builder = WebApplication.CreateBuilder(args);

// Force the application to listen on port 3000 across all network interfaces (0.0.0.0).
builder.WebHost.UseUrls("http://0.0.0.0:3000");

// Enable Cross-Origin Resource Sharing (CORS) with a permissive policy.
// This is useful for development but should be restricted to specific origins in production.
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

// Ensure the storage directory ('AIA') exists within the application's root directory.
// This folder will hold both the assessment templates and the uploaded results.
var aiaFolderPath = Path.Combine(app.Environment.ContentRootPath, "AIA");
if (!Directory.Exists(aiaFolderPath))
{
    Directory.CreateDirectory(aiaFolderPath);
    Console.WriteLine($"Created directory: {aiaFolderPath}");
}

// ✅ GET: Serve the AI Maturity Assessment Excel template.
// This endpoint locates the template file on disk and streams it back to the client
// with the appropriate Excel MIME type.
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

// ✅ POST: Receive and save assessment results.
// It reads a custom 'x-filename' header for the target name and copies the binary
// request body into a local file within the AIA folder.
app.MapPost("/api/save-assessment-results", async (HttpRequest request, HttpResponse response) =>
{
    // Default to Results.xlsx if no filename header is provided.
    var filename = request.Headers["x-filename"].FirstOrDefault() ?? "Results.xlsx";
    var filePath = Path.Combine(aiaFolderPath, filename);

    Console.WriteLine($"Received request to save: {filename}");

    using var memoryStream = new MemoryStream();
    await request.Body.CopyToAsync(memoryStream);

    if (memoryStream.Length == 0)
    {
        // Return a Bad Request if the body is empty.
        response.StatusCode = 400;
        await response.WriteAsync("No data received.");
        return;
    }

    try
    {
        // Save the stream contents to the specified path asynchronously.
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
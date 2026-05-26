
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;
using ClosedXML.Excel;

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

// ✅ POST: Receive and save assessment results as JSON data.
// This endpoint accepts a list of response objects, opens the local Excel master,
// appends the data to the 'AIA-Response-Capture' sheet, and saves the workbook.
app.MapPost("/api/save-assessment-results", async (List<AssessmentResponse> results, HttpResponse response) =>
{
    var filePath = Path.Combine(aiaFolderPath, "AI_Maturity_Assessment.xlsx");

    if (results == null || results.Count == 0)
    {
        response.StatusCode = 400;
        await response.WriteAsync("No data received.");
        return;
    }

    if (!File.Exists(filePath))
    {
        response.StatusCode = 404;
        await response.WriteAsync("Template file not found on server.");
        return;
    }

    try
    {
        // Open the existing workbook using ClosedXML
        using (var workbook = new XLWorkbook(filePath))
        {
            var sheet = workbook.Worksheet("AIA-Response-Capture") ?? workbook.AddWorksheet("AIA-Response-Capture");

            // Add headers if the sheet is new/empty
            if (sheet.LastRowUsed() == null)
            {
                sheet.Cell(1, 1).Value = "Respondent Role";
                sheet.Cell(1, 2).Value = "Question ID";
                sheet.Cell(1, 3).Value = "Chosen Response";
                sheet.Cell(1, 4).Value = "Response Score";
            }

            // Append each response to the next available row
            var nextRow = sheet.LastRowUsed()?.RowNumber() + 1 ?? 1;
            foreach (var res in results)
            {
                sheet.Cell(nextRow, 1).Value = res.Role;
                sheet.Cell(nextRow, 2).Value = res.QuestionId;
                sheet.Cell(nextRow, 3).Value = res.ResponseText;
                sheet.Cell(nextRow, 4).Value = res.ResponseScore;
                nextRow++;
            }

            workbook.Save();
        }

        Console.WriteLine($"Successfully appended {results.Count} results to {filePath}");

        response.ContentType = "application/json";
        await response.WriteAsJsonAsync(new { message = "Results saved successfully" });
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error saving file: {ex}");
        response.StatusCode = 500;
        await response.WriteAsync($"Failed to save file: {ex.Message}");
    }
});

app.Run();

// DTO for incoming assessment data
public record AssessmentResponse(string Role, string QuestionId, string ResponseText, int ResponseScore);
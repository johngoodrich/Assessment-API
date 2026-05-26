
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;
using System.IO;
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

// 1. Set up Paths and Global Configuration
var aiaFolderPath = Path.Combine(app.Environment.ContentRootPath, "AIA");
if (!Directory.Exists(aiaFolderPath))
{
    Directory.CreateDirectory(aiaFolderPath);
}

var excelPath = Path.Combine(aiaFolderPath, "AI_Maturity_Assessment.xlsx");
var frontendPath = Path.Combine(app.Environment.ContentRootPath, "..", "assessment-questionnaire", "dist", "assessment-questionnaire", "browser");

Console.WriteLine($"[INIT] Looking for Excel file at: {Path.GetFullPath(excelPath)}");

// 2. Middleware Configuration
app.UseCors();

// 3. API Routes (Grouped together for reliability)

// ✅ GET: Get list of roles from 'AIA-UI-Config'
app.MapGet("/api/config/roles", () =>
{
    if (!File.Exists(excelPath)) return Results.NotFound();

    try
    {
        using var workbook = new XLWorkbook(excelPath);
        var sheet = workbook.Worksheet("AIA-UI-Config");
        if (sheet == null) return Results.NotFound("UI Config sheet missing.");

        var roles = sheet.RowsUsed()
            .Skip(1) // Skip header
            .Select(r => r.Cell(1).GetValue<string>().Trim())
            .Where(r => !string.IsNullOrEmpty(r))
            .Distinct()
            .ToList();

        return Results.Ok(roles);
    }
    catch (Exception ex)
    {
        return Results.Problem($"Error reading roles: {ex.Message}");
    }
});

// ✅ GET: Get questions and mapped answers for a specific role
app.MapGet("/api/config/questions", (string role) =>
{
    if (!File.Exists(excelPath)) return Results.NotFound();
    if (string.IsNullOrEmpty(role)) return Results.BadRequest("Role is required.");

    try
    {
        using var workbook = new XLWorkbook(excelPath);
        
        // 1. Load Answer Map
        var answerSheet = workbook.Worksheet("AIA-Question-Response-Map");
        var answerMap = answerSheet.RowsUsed()
            .Skip(1)
            .Select(r => new { 
                Id = r.Cell(1).GetValue<string>(), 
                Text = r.Cell(2).GetValue<string>(), 
                Score = r.Cell(3).GetValue<int>() 
            })
            .GroupBy(a => a.Id)
            .ToDictionary(g => g.Key, g => g.Select(x => new ResponseOption(x.Text, x.Score)).ToList());

        // 2. Load Questions filtered by Role
        var questionSheet = workbook.Worksheet("AIA-Question-Config");
        if (questionSheet == null) return Results.NotFound("Question Config sheet missing.");

        var questions = questionSheet.RowsUsed()
            .Skip(1)
            .Where(r => r.Cell(2).GetValue<string>().Equals(role, StringComparison.OrdinalIgnoreCase))
            .Select(r => {
                var id = r.Cell(1).GetValue<string>();
                return new
                {
                    id = id,
                    role = r.Cell(2).GetValue<string>(),
                    dimension = r.Cell(3).GetValue<string>(),
                    pillar = r.Cell(4).GetValue<string>(),
                    question = r.Cell(6).GetValue<string>(),
                    responses = answerMap.TryGetValue(id, out var opts) ? opts : new List<ResponseOption>()
                };
            })
            .ToList();

        return Results.Ok(questions);
    }
    catch (Exception ex)
    {
        return Results.Problem($"Error reading questions: {ex.Message}");
    }
});

// ✅ GET: Get maturity band definitions from 'AIA-UI-Config'
app.MapGet("/api/config/bands", () =>
{
    if (!File.Exists(excelPath)) return Results.NotFound();

    try
    {
        using var workbook = new XLWorkbook(excelPath);
        var sheet = workbook.Worksheet("AIA-UI-Config");
        
        var bands = sheet.RowsUsed()
            .Skip(1)
            .Select(r => new {
                name = r.Cell(2).GetValue<string>(),
                score = r.Cell(3).GetValue<string>()
            })
            .Where(b => !string.IsNullOrEmpty(b.name))
            .ToList();

        return Results.Ok(bands);
    }
    catch (Exception ex)
    {
        return Results.Problem($"Error reading bands: {ex.Message}");
    }
});

// ✅ GET: Get calculated scores for the radar chart
app.MapGet("/api/chart-data", () =>
{
    if (!File.Exists(excelPath)) return Results.NotFound();

    try
    {
        // We use non-sharing mode to ensure we get a fresh read even if results were just saved
        using var workbook = new XLWorkbook(excelPath);
        
        // Force Excel to recalculate formulas so chart data is accurate 
        // based on the AIA-Response-Capture sheet data.
        workbook.RecalculateAllFormulas();

        var calcSheet = workbook.Worksheet("AIA-Calculations");
        var uiSheet = workbook.Worksheet("AIA-UI-Config");

        if (calcSheet == null || uiSheet == null) return Results.NotFound("Required sheets missing.");

        var datasets = calcSheet.Rows(2, 5).Select(r => new {
            label = r.Cell(1).GetValue<string>(),
            data = new[] { r.Cell(2).GetValue<double>(), r.Cell(3).GetValue<double>(), r.Cell(4).GetValue<double>() },
            backgroundColor = uiSheet.Row(r.RowNumber()).Cell(4).GetValue<string>(),
            borderColor = uiSheet.Row(r.RowNumber()).Cell(5).GetValue<string>(),
            borderWidth = uiSheet.Row(r.RowNumber()).Cell(6).GetValue<int>()
        }).ToList();

        return Results.Ok(datasets);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[ERROR] /api/chart-data: {ex.Message}");
        return Results.Problem("Error processing chart data. Ensure Excel formulas are evaluated and cells contain valid numbers.");
    }
});

// ✅ POST: Receive and save assessment results as JSON data.
app.MapPost("/api/save-assessment-results", async (List<AssessmentResponse> results, HttpResponse response) =>
{
    if (results == null || results.Count == 0)
    {
        response.StatusCode = 400;
        await response.WriteAsync("No data received.");
        return;
    }

    if (!File.Exists(excelPath))
    {
        response.StatusCode = 404;
        await response.WriteAsync("Template file not found on server.");
        return;
    }

    try
    {
        using (var workbook = new XLWorkbook(excelPath))
        {
            var sheet = workbook.Worksheet("AIA-Response-Capture") ?? workbook.AddWorksheet("AIA-Response-Capture");

            if (sheet.LastRowUsed() == null)
            {
                sheet.Cell(1, 1).Value = "Respondent Role";
                sheet.Cell(1, 2).Value = "Question ID";
                sheet.Cell(1, 3).Value = "Chosen Response";
                sheet.Cell(1, 4).Value = "Response Score";
            }

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

        Console.WriteLine($"[SUCCESS] Appended {results.Count} results to {excelPath}");
        response.ContentType = "application/json";
        await response.WriteAsJsonAsync(new { message = "Results saved successfully" });
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[ERROR] Failed to save results: {ex.Message}");
        response.StatusCode = 500;
        await response.WriteAsync($"Failed to save file: {ex.Message}");
    }
});

// ✅ GET: Serve the Excel template (Legacy endpoint for debugging)
app.MapGet("/api/get-assessment-template", () => 
    File.Exists(excelPath) 
        ? Results.File(excelPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") 
        : Results.NotFound());


// 4. Static Files and SPA Fallback (Must come LAST)
if (Directory.Exists(frontendPath))
{
    app.UseDefaultFiles(new DefaultFilesOptions
    {
        FileProvider = new PhysicalFileProvider(frontendPath)
    });
    app.UseStaticFiles(new StaticFileOptions
    {
        FileProvider = new PhysicalFileProvider(frontendPath)
    });
}
else
{
    Console.WriteLine($"Warning: Frontend build not found at {frontendPath}. Static files will not be served.");
}

if (Directory.Exists(frontendPath))
{
    app.MapFallbackToFile("index.html", new StaticFileOptions
    {
        FileProvider = new PhysicalFileProvider(frontendPath)
    });
}

app.Run();

// DTO for incoming assessment data
public record AssessmentResponse(string Role, string QuestionId, string ResponseText, int ResponseScore);

public record ResponseOption(string Response, int Score);
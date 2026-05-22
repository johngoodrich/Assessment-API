# AI Maturity Assessment - Backend API

A lightweight .NET Minimal API designed to manage Excel-based assessment templates and result persistence.

## Technologies Used
- .NET (C#)
- ASP.NET Core Minimal API

## Core Architecture
The service acts as a simple file bridge. It manages a local `AIA/` directory relative to the content root, serving as a repository for both the master template and user-submitted results.

## API Endpoints

### `GET /api/get-assessment-template`
Locates and streams the master Excel workbook.

**Response:**
- **Status:** `200 OK`
- **Content-Type:** `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- **Errors:** `404 Not Found` if the template is missing from the `AIA/` directory.

### `POST /api/save-assessment-results`
Performs a binary copy of the request body to a local file.

**Headers:**
- `x-filename`: (Optional) The target filename (e.g., `UserDept_Assessment.xlsx`). Defaults to `Results.xlsx`.

**Request Body:**
- Expected: Raw binary stream of the `.xlsx` file.

**Response:**
- **Status:** `200 OK`
- **Body:** 
  ```json
  {
    "message": "File saved successfully",
    "path": ".../AIA/filename.xlsx"
  }
  ```
- **Errors:** `400 Bad Request` if the body is empty; `500 Internal Server Error` on IO exceptions.

## Development Workflow

### Configuration
The application is hardcoded to listen on `http://0.0.0.0:3000`. This allows for easy access across containers or local network interfaces.

### Local Setup
```bash
# Restore and build
dotnet build

# Run in development mode
dotnet run
```

### Directory Structure
```text
assessment-backend/
├── AIA/                 # Auto-created storage directory
│   └── AI_Maturity_Assessment.xlsx  # Place master template here
├── Program.cs           # Main entry point & Route definitions
└── assessment-backend.csproj
```

## Infrastructure Notes

### CORS
Currently uses a global permissive policy:
```csharp
policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod();
```
*Warning: This should be restricted to the specific frontend origin before production deployment.*

### File Handling
The service uses `MemoryStream` to buffer the request body before writing to disk. For very large workbooks, consider refactoring to stream directly to a `FileStream` to reduce memory overhead.

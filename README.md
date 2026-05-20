# Assessment Backend

This is the backend service for the AI Maturity Assessment application. It provides API endpoints to serve an Excel assessment template and to save the results of a completed assessment.

## Technologies Used

- .NET (C#)
- ASP.NET Core Minimal APIs

## Prerequisites

- .NET SDK (version 6.0 or higher)

## Setup and Running

1.  **Clone the repository:**

    ```bash
    git clone <repository-url>
    cd assessment-backend
    ```

2.  **Restore dependencies:**

    ```bash
    dotnet restore
    ```

3.  **Run the application:**

    ```bash
    dotnet run
    ```

    The application will start and listen on `http://localhost:3000`.

## API Endpoints

### `GET /api/get-assessment-template`

Serves the `AI_Maturity_Assessment.xlsx` Excel template file.

### `POST /api/save-assessment-results`

Receives an Excel file (e.g., `Results.xlsx`) and saves it to the `AIA` folder.

**Headers:**
- `x-filename`: (Optional) Specifies the desired filename for the saved results. Defaults to `Results.xlsx`.

**Request Body:**
- The raw binary content of the Excel file.

## Project Structure

- `Program.cs`: Entry point of the application, configures services, CORS, and defines API endpoints.
- `AIA/`: Directory where the `AI_Maturity_Assessment.xlsx` template is stored and where assessment results are saved. This folder is created if it doesn't exist.

## CORS Policy

The application is configured with a permissive CORS policy (`AllowAnyOrigin`, `AllowAnyHeader`, `AllowAnyMethod`) for development purposes. For production deployments, it is recommended to restrict this policy to specific origins.

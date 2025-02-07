# Genie For Excel

## Description

Genie For Excel is a desktop application designed to enhance productivity and automate repetitive tasks in Microsoft Excel. This tool integrates seamlessly with Excel's COM objects and offers additional AI-driven functionality for analyzing workbooks, generating formula insights, auditing spreadsheets, creating macros, and more.

## Requirements

-   **Operating System**: Windows 10 or above (Windows 11 recommended)
-   **.NET SDK**: .NET 8.0 (or higher)
-   **Excel**: Microsoft Office Excel 2016 or later
-   **Internet Connection**: Needed for certain cloud/AI features
-   **Palantir Foundry**: (Optional) If leveraging Palantir integration for advanced analytics

## Installation

1.  **Clone or Download** the repository:

    ```bash
    git clone [https://github.com/YourOrg/GenieDevConSubmission.git](https://github.com/YourOrg/GenieDevConSubmission.git)
    ```

    Or download the .zip file and extract it to a folder of your choice.

2.  Open the project in Visual Studio 2022 or a compatible IDE. Alternatively, you can build from the command line using:

    ```bash
    cd GenieDevConSubmission
    dotnet build
    ```

3.  Restore Packages: The project references several NuGet packages for Excel interop, iText, DocumentFormat.OpenXml, NPOI, etc. They will be restored automatically on build.

4.  Build the Application:
    -   Inside Visual Studio, select Build -> Build Solution.
    -   Or from the command line, run:

        ```bash
        dotnet build
        ```

5.  Run the Executable:
    -   You can launch the project from Visual Studio (F5).
    -   Or navigate to the output folder (e.g., `bin\Debug\net8.0-windows`) and run `ExcelGenie.exe`.

## Configuration

1.  Palantir Integration (Optional):
    -   Within the `PalantirService.cs` file, add your bearer token, base URL, and ontology ID. Make sure these values match your environment.
    -   Example placeholders:

        ```csharp
        private const string PalantirToken = "YOUR_PALANTIR_BEARER_TOKEN";
        private const string baseUrl = "[https://your-palantir-instance.com](https://your-palantir-instance.com)";
        private const string ontologyId = "YOUR_ONTOLOGY_ID";
        ```

    -   Rebuild the project after updating these values.

2.  Custom Instructions:
    -   Custom instructions can be saved through the application’s Custom button or found in `CustomInstructionsWindow.xaml.cs` under user settings. They are stored in the local user configuration so they persist across sessions.

3.  Logging and Folders:
    -   Log files are stored in the `\Documents\GenieForExcel\Logs` folder. The application automatically creates and rotates log files based on the session start time.
    -   You can change the log file path in `BlackboxLogger.cs` by modifying `SetLogFilePath`.

4.  Settings:
    -   The application settings file `Settings.settings` and `Settings.Designer.cs` store user-specific data, including any custom instructions or default user email. Adjust these as needed.

## Usage

1.  Running the Application:
    -   Launch `ExcelGenie.exe`. You will see a main window with toolbar and a place to open or create Excel files.

2.  Opening or Creating Workbooks:
    -   In the Start screen, either Open File or Create a New File. You can also select from recently used files.

3.  Main Features:
    -   **Workbook Explorer**: Displays workbooks, worksheets, and charts in a tree view.
    -   **Context/Drag & Drop**: Drag additional context files (Docx, PDF, etc.) onto the “Add Context” button to embed relevant data for AI-driven tasks.
    -   **Select Range**: Choose a cell range directly from Excel for more advanced tasks, like formula generation or data processing.
    -   **Quick Actions**: One-click shortcuts such as “Audit Formulas”, “Suggest Improvements”, and “Make Print-Ready.”
    -   **Chat Panel**: Interact with the AI by typing instructions. AI can generate formulas, macros, or suggestions. The conversation is saved locally to maintain context.

4.  AI-Driven Steps:
    -   After sending a prompt (click submit), the system may retrieve an “Agent Plan” or “Plan Steps” that can be directly executed in Excel to perform advanced tasks.

5.  Palantir Integration:
    -   If you have a Palantir environment configured, Genie For Excel can push your workbook structure, cells, and charts into Palantir Foundry for advanced analytics. The response from Palantir can further drive actionable “Agent Plans” inside Excel.

6.  Undo & Backup:
    -   Use the built-in backup to revert changes if a generated macro modifies your workbook undesirably. The system creates a backup (SaveCopyAs) and can revert to that backup file if needed.

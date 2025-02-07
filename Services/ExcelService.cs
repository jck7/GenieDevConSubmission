using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using System.IO;
using System.Text;
using System.Reflection;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace ExcelGenie.Services
{
    public class ExcelService : IDisposable
    {
        private Excel.Application? _excelApp;
        private Excel.Workbook? _workbook;
        private readonly BlackboxLogger _logger;
        private readonly Dispatcher _uiDispatcher;
        private bool _disposed;

        // Events
        public event EventHandler<string>? WorkbookLoaded;
        public event EventHandler<string>? WorksheetActivated;
        public event EventHandler<string>? ErrorOccurred;
        public event EventHandler? WorkbookClosed;
        public event EventHandler<(string oldName, string newName)>? WorksheetRenamed;
        public event EventHandler<string>? WorksheetDeleted;
        public event EventHandler? ConnectionLost;
        public event EventHandler? ConnectionEstablished;

        // Properties
        public bool IsWorkbookOpen => _workbook != null;
        public bool IsConnected => _excelApp != null && _workbook != null;
        public string? CurrentWorkbookPath { get; private set; }
        public Excel.Workbook? ActiveWorkbook => _workbook;
        public Excel.Worksheet? ActiveSheet => _excelApp?.ActiveSheet as Excel.Worksheet;
        public Excel.Window? ActiveWindow => _excelApp?.ActiveWindow;
        public Excel.Range? Selection => _excelApp?.Selection as Excel.Range;
        public int Hwnd => _excelApp?.Hwnd ?? 0;

        // Excel Application Properties
        public bool DisplayAlerts
        {
            get => _excelApp?.DisplayAlerts ?? false;
            set
            {
                if (_excelApp != null)
                {
                    _excelApp.DisplayAlerts = value;
                }
            }
        }

        public bool Visible
        {
            get => _excelApp?.Visible ?? false;
            set
            {
                if (_excelApp != null)
                {
                    _excelApp.Visible = value;
                }
            }
        }

        public bool Interactive
        {
            get => _excelApp?.Interactive ?? false;
            set
            {
                if (_excelApp != null)
                {
                    _excelApp.Interactive = value;
                }
            }
        }

        public bool ScreenUpdating
        {
            get => _excelApp?.ScreenUpdating ?? false;
            set
            {
                if (_excelApp != null)
                {
                    _excelApp.ScreenUpdating = value;
                }
            }
        }

        public bool EnableEvents
        {
            get => _excelApp?.EnableEvents ?? false;
            set
            {
                if (_excelApp != null)
                {
                    _excelApp.EnableEvents = value;
                }
            }
        }

        public Excel.XlWindowState WindowState
        {
            get => _excelApp?.WindowState ?? Excel.XlWindowState.xlNormal;
            set
            {
                if (_excelApp != null)
                {
                    _excelApp.WindowState = value;
                }
            }
        }

        public Excel.Application? Application => _excelApp;

        public ExcelService(BlackboxLogger logger, Dispatcher uiDispatcher)
        {
            _logger = logger;
            _uiDispatcher = uiDispatcher;
        }

        public async Task<Excel.Application?> GetExcelInstance()
        {
            try
            {
                // Create Excel instance on a background thread
                var newApp = await Task.Run(() =>
                {
                    try
                    {
                        var app = new Excel.Application();
                        if (app == null)
                        {
                            throw new Exception("Failed to create Excel instance");
                        }

                        _logger.LogSystemMessage($"Excel Application Version: {app.Version}");
                        _logger.LogSystemMessage($"Excel Application Ready: {app.Ready}");
                        _logger.LogSystemMessage($"Excel Application Visible: {app.Visible}");
                        _logger.LogSystemMessage($"Excel Application Interactive: {app.Interactive}");
                        return app;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error creating Excel instance on background thread: {ex.Message}", ex);
                        return null;
                    }
                });

                return newApp;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetExcelInstance: {ex.Message}", ex);
                return null;
            }
        }

        public async Task<bool> InitializeExcel()
        {
            try
            {
                _excelApp = await GetExcelInstance();
                if (_excelApp == null)
                {
                    throw new Exception("Could not create Excel instance.");
                }

                // Configure initial Excel settings
                _excelApp.Visible = false;
                _excelApp.DisplayAlerts = false;
                _excelApp.ScreenUpdating = false;
                _excelApp.Interactive = false;
                _excelApp.EnableEvents = false;
                _excelApp.WindowState = Excel.XlWindowState.xlNormal;

                ConnectionEstablished?.Invoke(this, EventArgs.Empty);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error initializing Excel: {ex.Message}");
                ConnectionLost?.Invoke(this, EventArgs.Empty);
                return false;
            }
        }

        public async Task<bool> LoadWorkbookAsync(string? filePath, string genieForExcelFolderPath, bool createNew = false)
        {
            if (!createNew && string.IsNullOrEmpty(filePath))
            {
                _logger.LogError("No file path provided for workbook load");
                return false;
            }

            bool success = false;
            try
            {
                if (!createNew && filePath != null)
                {
                    if (!File.Exists(filePath))
                    {
                        string error = $"File not found: {filePath}";
                        _logger.LogError(error);
                        ErrorOccurred?.Invoke(this, error);
                        return false;
                    }
                }

                if (_excelApp == null)
                {
                    if (!await InitializeExcel())
                    {
                        return false;
                    }
                }

                if (createNew)
                {
                    _workbook = _excelApp?.Workbooks.Add();
                    if (_workbook == null)
                    {
                        throw new Exception("Could not create new workbook.");
                    }
                    CurrentWorkbookPath = null;
                    success = true;
                }
                else if (filePath != null)
                {
                    _workbook = _excelApp?.Workbooks.Open(filePath);
                    if (_workbook == null)
                    {
                        throw new Exception("Could not open workbook.");
                    }
                    CurrentWorkbookPath = filePath;
                    success = true;
                }

                if (success)
                {
                    _excelApp!.Visible = true;
                    _excelApp.DisplayAlerts = true;
                    _excelApp.ScreenUpdating = true;
                    _excelApp.Interactive = true;
                    _excelApp.EnableEvents = true;

                    WorkbookLoaded?.Invoke(this, CurrentWorkbookPath ?? "New Workbook");
                }

                return success;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error loading workbook: {ex.Message}");
                ErrorOccurred?.Invoke(this, ex.Message);
                return false;
            }
        }

        public async Task<string> ExtractWorksheetDataAsync(Excel.Worksheet worksheet)
        {
            if (worksheet == null) return string.Empty;

            try
            {
                var sb = new StringBuilder();
                Excel.Range usedRange = worksheet.UsedRange;
                object[,] values = usedRange.Value2;

                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    for (int col = 1; col <= values.GetLength(1); col++)
                    {
                        if (values[row, col] != null)
                        {
                            sb.Append(values[row, col].ToString()).Append("\t");
                        }
                    }
                    sb.AppendLine();
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error extracting worksheet data: {ex.Message}");
                return string.Empty;
            }
        }

        public async Task ExecuteVBAInExcel(string vbaCode, bool removeModuleAfterRun = true)
        {
            try
            {
                if (_workbook?.VBProject == null)
                {
                    throw new Exception("VBA Project is not accessible. Make sure the workbook is open and " +
                                      "'Trust access to the VBA project object model' is enabled.");
                }

                string moduleName = $"GenieModule_{DateTime.Now:yyyyMMddHHmmssfff}";
                var vbComponent = await AddVBAModule(moduleName, vbaCode);

                if (vbComponent != null)
                {
                    try
                    {
                        _excelApp?.Run($"{moduleName}.Main");
                    }
                    finally
                    {
                        if (removeModuleAfterRun)
                        {
                            await RemoveVBAModule(moduleName, vbComponent);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error executing VBA code: {ex.Message}", ex);
                throw;
            }
        }

        public async Task ExecutePlanStepsAsync(IEnumerable<(int index, string title, string vbaCode)> planSteps)
        {
            foreach (var step in planSteps)
            {
                _logger.LogSystemMessage($"=== Running plan step #{step.index}: {step.title} ===");
                await ExecuteVBAInExcel(step.vbaCode);
            }
        }

        public async Task ExecutePlanStepsAsync(IEnumerable<string> planSteps)
        {
            int stepIndex = 1;
            foreach (var stepCode in planSteps)
            {
                _logger.LogSystemMessage($"=== Running plan step #{stepIndex} ===");
                await ExecuteVBAInExcel(stepCode);
                stepIndex++;
            }
        }

        private async Task<VBIDE.VBComponent?> AddVBAModule(string moduleName, string vbaCode)
        {
            if (_workbook == null)
                throw new InvalidOperationException("No active workbook");
            if (_workbook.VBProject == null)
                throw new InvalidOperationException("VBA Project not accessible. " +
                                                  "Make sure 'Trust access to the VBA project object model' is enabled in Excel.");

            return await Task.Run(() =>
            {
                try
                {
                    VBIDE.VBComponent vbComponent = _workbook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
                    vbComponent.Name = moduleName;
                    vbComponent.CodeModule.AddFromString(vbaCode);
                    return vbComponent;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error adding VBA module: {ex.Message}", ex);
                    throw;
                }
            });
        }

        private async Task RemoveVBAModule(string moduleName, VBIDE.VBComponent vbComponent)
        {
            if (_workbook?.VBProject == null) return;

            await Task.Run(() =>
            {
                try
                {
                    _workbook.VBProject.VBComponents.Remove(vbComponent);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error removing VBA module {moduleName}: {ex.Message}", ex);
                    throw;
                }
            });
        }

        public async Task CleanupExcelObjects()
        {
            try
            {
                if (_workbook != null)
                {
                    _workbook.Close(false);
                    Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }

                if (_excelApp != null)
                {
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                    _excelApp = null;
                }

                CurrentWorkbookPath = null;
                WorkbookClosed?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error cleaning up Excel objects: {ex.Message}");
            }
        }

        public void ActivateWorksheet(string worksheetName)
        {
            try
            {
                if (_workbook == null)
                {
                    throw new InvalidOperationException("No active workbook");
                }

                var worksheet = _workbook.Worksheets[worksheetName] as Excel.Worksheet;
                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Worksheet '{worksheetName}' not found");
                }

                worksheet.Activate();
                WorksheetActivated?.Invoke(this, worksheetName);
                Marshal.ReleaseComObject(worksheet);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error activating worksheet: {ex.Message}");
                throw;
            }
        }

        public object? InputBox(string Prompt, string Title, object? Default = null, int XPos = -1, int YPos = -1, object? HelpFile = null, object? HelpContextID = null, int Type = 0)
        {
            return _excelApp?.InputBox(Prompt, Title, Default ?? Missing.Value, XPos, YPos, HelpFile ?? Missing.Value, HelpContextID ?? Missing.Value, Type);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    try
                    {
                        if (_workbook != null)
                        {
                            _workbook.Close(false);
                            Marshal.ReleaseComObject(_workbook);
                            _workbook = null;
                        }

                        if (_excelApp != null)
                        {
                            _excelApp.Quit();
                            Marshal.ReleaseComObject(_excelApp);
                            _excelApp = null;
                        }

                        ConnectionLost?.Invoke(this, EventArgs.Empty);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error during disposal: {ex.Message}");
                    }
                }
                _disposed = true;
            }
        }

        public async Task<string> GetCurrentWorksheetData()
        {
            if (ActiveSheet == null)
                return string.Empty;

            return await Task.Run(() => ExtractWorksheetDataAsync(ActiveSheet));
        }

        public async Task<string> GetActiveWorksheetName()
        {
            if (ActiveSheet == null)
                return string.Empty;

            return await Task.Run(() => ActiveSheet.Name);
        }

        public List<string> GetWorksheetNames()
        {
            if (_workbook == null) return new List<string>();
            return _workbook.Worksheets.Cast<Worksheet>().Select(ws => ws.Name).ToList();
        }

        public async Task<Dictionary<string, string>> GetExcelApplicationInfo()
        {
            if (Application == null)
            {
                throw new InvalidOperationException("Excel application is not initialized");
            }

            return new Dictionary<string, string>
            {
                { "version_timestamp", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") },
                { "last_mod_time", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") },
                { "excel_version", Application.Version?.ToString() ?? "Unknown" },
                { "active_workbook_guid", Guid.NewGuid().ToString() },
                { "user_name", Application.UserName?.ToString() ?? "Unknown" },
                { "command_bars", (Convert.IsDBNull(Application.CommandBars) ? 0 : Application.CommandBars.Count).ToString() },
                { "protected_view_windows", (Convert.IsDBNull(Application.ProtectedViewWindows) ? 0 : Application.ProtectedViewWindows.Count).ToString() },
                { "auto_correct", (Convert.IsDBNull(Application.AutoCorrect) ? false : Application.AutoCorrect.DisplayAutoCorrectOptions).ToString() },
                { "file_converters", (Convert.IsDBNull(Application.FileConverters) ? 0 : Application.FileConverters.Count).ToString() },
                { "language_settings", (Convert.IsDBNull(Application.LanguageSettings) ? 0 : Application.LanguageSettings.LanguageID[Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI]).ToString() },
                { "error_check_options", (Convert.IsDBNull(Application.ErrorCheckingOptions) ? false : Application.ErrorCheckingOptions.BackgroundChecking).ToString() }
            };
        }

        public async Task<Dictionary<string, string>> GetWorkbookInfo()
        {
            var workbook = ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("No active workbook");
            }

            var workbookGuid = Guid.NewGuid().ToString();

            try
            {
                object? activeSheet = workbook.ActiveSheet;
                string sheetName = "Sheet1";

                if (activeSheet is Excel.Worksheet sheet)
                {
                    sheetName = sheet.Name ?? "Sheet1";
                }

                return new Dictionary<string, string>
                {
                    { "active_sheet", sheetName },
                    { "author", workbook.Author ?? "Unknown" },
                    { "connection_list", "[]" },  // This will be handled separately by CreateExcelConnections
                    { "workbook_events", "[]" },  // This can be enhanced if needed
                    { "workbook_guid", workbookGuid },
                    { "version_timestamp", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") },
                    { "last_mod_time", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") }
                };
            }
            catch (Exception ex)
            {
                _logger.LogError("Error getting workbook info", ex);
                throw;
            }
        }

        public Excel.Workbook? GetActiveWorkbook()
        {
            return ActiveWorkbook;
        }

        public async Task<List<Excel.WorkbookConnection>> GetWorkbookConnections()
        {
            var workbook = GetActiveWorkbook();
            if (workbook == null)
            {
                return new List<Excel.WorkbookConnection>();
            }

            var connections = new List<Excel.WorkbookConnection>();
            foreach (Excel.WorkbookConnection connection in workbook.Connections)
            {
                connections.Add(connection);
            }
            return connections;
        }

        public async Task<List<Excel.Worksheet>> GetWorksheets()
        {
            var workbook = GetActiveWorkbook();
            if (workbook == null)
            {
                return new List<Excel.Worksheet>();
            }

            var worksheets = new List<Excel.Worksheet>();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                worksheets.Add(worksheet);
            }
            return worksheets;
        }

        public List<string> GetChartsInWorksheet(string worksheetName)
        {
            var charts = new List<string>();

            try
            {
                if (_workbook == null)
                {
                    _logger.LogError("No workbook is open");
                    return charts;
                }

                var worksheet = _workbook.Worksheets[worksheetName] as Excel.Worksheet;
                if (worksheet == null)
                {
                    _logger.LogError($"Worksheet {worksheetName} not found");
                    return charts;
                }

                var chartObjects = worksheet.ChartObjects() as Excel.ChartObjects;
                if (chartObjects != null)
                {
                    foreach (Excel.ChartObject chartObject in chartObjects)
                    {
                        if (chartObject.Name != null)
                        {
                            charts.Add(chartObject.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting charts in worksheet {worksheetName}: {ex.Message}");
            }

            return charts;
        }

        public async Task ExecutePlanStepsFromJson(dynamic planStepsResponse)
        {
            try
            {
                var planStepsList = new List<(int index, string title, string vbaCode)>();
                
                foreach (var step in planStepsResponse.value)
                {
                    string vbaCode = step.validVBACode?.ToString() ?? "";
                    // Consider any VBA code with a Sub declaration as executable
                    if (!string.IsNullOrWhiteSpace(vbaCode) && vbaCode.Contains("Sub "))
                    {
                        int stepIndex = Convert.ToInt32(step.stepIndex);
                        string stepTitle = step.stepTitle?.ToString() ?? $"Step {stepIndex}";

                        _logger.LogSystemMessage($"Parsed step {stepIndex}: {stepTitle}");
                        _logger.LogSystemMessage($"VBA code length: {vbaCode.Length} characters");
                        planStepsList.Add((stepIndex, stepTitle, vbaCode));
                    }
                    else
                    {
                        _logger.LogSystemMessage($"Skipping step {step.stepIndex}: No executable VBA code found");
                    }
                }

                if (!planStepsList.Any())
                {
                    _logger.LogSystemMessage("No executable VBA steps found");
                    return;
                }

                _logger.LogSystemMessage($"Found {planStepsList.Count} executable VBA steps");
                _logger.LogSystemMessage("Beginning execution of VBA steps...");

                // Sort steps by index to ensure correct order
                var orderedSteps = planStepsList.OrderBy(s => s.index).ToList();
                
                foreach (var step in orderedSteps)
                {
                    _logger.LogSystemMessage($"Executing step {step.index}: {step.title}");
                    try
                    {
                        await ExecuteVBAInExcel(step.vbaCode);
                        _logger.LogSystemMessage($"Successfully completed step {step.index}");
                    }
                    catch (Exception stepEx)
                    {
                        _logger.LogError($"Failed to execute step {step.index}: {stepEx.Message}", stepEx);
                        throw;
                    }
                }

                _logger.LogSystemMessage("VBA steps execution completed successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error executing VBA steps: {ex.Message}", ex);
                throw;
            }
        }

        public async Task ExecutePlanStepsWithBackup(IEnumerable<(int index, string title, string vbaCode)> planSteps)
        {
            if (_workbook == null)
            {
                throw new InvalidOperationException("No active workbook");
            }

            string? backupPath = null;
            try
            {
                // Create backup before starting
                backupPath = await CreateWorkbookBackup();
                _logger.LogSystemMessage($"Created backup at: {backupPath}");

                await ExecutePlanStepsAsync(planSteps);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error during plan execution: {ex.Message}", ex);
                
                if (backupPath != null && File.Exists(backupPath))
                {
                    _logger.LogSystemMessage("Attempting to restore from backup...");
                    try
                    {
                        await CleanupExcelObjects();
                        await LoadWorkbookAsync(backupPath, Path.GetDirectoryName(backupPath) ?? "");
                        _logger.LogSystemMessage("Backup restored successfully");
                    }
                    catch (Exception restoreEx)
                    {
                        _logger.LogError($"Failed to restore backup: {restoreEx.Message}", restoreEx);
                    }
                }
                throw;
            }
            finally
            {
                // Clean up backup file if it exists
                if (backupPath != null && File.Exists(backupPath))
                {
                    try
                    {
                        File.Delete(backupPath);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Failed to delete backup file: {ex.Message}", ex);
                    }
                }
            }
        }

        private async Task<string> CreateWorkbookBackup()
        {
            if (_workbook == null || string.IsNullOrEmpty(CurrentWorkbookPath))
            {
                throw new InvalidOperationException("No active workbook to backup");
            }

            string backupPath = Path.Combine(
                Path.GetDirectoryName(CurrentWorkbookPath) ?? "",
                $"Backup_{Path.GetFileNameWithoutExtension(CurrentWorkbookPath)}_{DateTime.Now:yyyyMMddHHmmssfff}{Path.GetExtension(CurrentWorkbookPath)}"
            );

            await Task.Run(() =>
            {
                _workbook.SaveCopyAs(backupPath);
            });

            return backupPath;
        }

        public async Task ExecutePlanStepsWithBackupFromJson(dynamic planStepsResponse)
        {
            try
            {
                var planStepsList = new List<(int index, string title, string vbaCode)>();
                
                foreach (var step in planStepsResponse.value)
                {
                    if (step.validVBACode == null || string.IsNullOrEmpty(step.validVBACode.ToString()))
                    {
                        continue;
                    }

                    int stepIndex = Convert.ToInt32(step.stepIndex);
                    string stepTitle = step.stepTitle?.ToString() ?? $"Step {stepIndex}";
                    string vbaCode = step.validVBACode.ToString();

                    planStepsList.Add((stepIndex, stepTitle, vbaCode));
                }

                if (planStepsList.Any())
                {
                    await ExecutePlanStepsWithBackup(planStepsList.OrderBy(s => s.index));
                }
                else
                {
                    throw new InvalidOperationException("No valid plan steps found to execute");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error executing plan steps with backup: {ex.Message}", ex);
                throw;
            }
        }
    }
} 


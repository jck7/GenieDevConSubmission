using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Net.Http.Json;
using System.Text.Json;

public class PalantirService : IDisposable
{
    public event EventHandler<bool>? ConnectionStatusChanged;

    public class EntityReference
    {
        public string LocalGuid { get; set; }
        public string PalantirPrimaryKey { get; set; }

        public EntityReference(string localGuid, string palantirPrimaryKey)
        {
            LocalGuid = localGuid;
            PalantirPrimaryKey = palantirPrimaryKey;
        }
    }

    private readonly ILogger<PalantirService> _logger;
    private readonly HttpClient _httpClient;
    private readonly string _bearerToken;
    private string _latestPromptPalantirKey;
    private readonly string _ontologyId;
    private readonly string _baseUrl;
    private bool _cellValidationDone = false;
    private bool _chartValidationDone = false;
    private const int BATCH_SIZE = 20;

    public string BaseUrl => _baseUrl;
    public string OntologyId => _ontologyId;
    public string BearerToken => _bearerToken;
    public string LatestPromptKey => _latestPromptPalantirKey;
    public HttpClient HttpClient => _httpClient;

    public PalantirService(ILogger<PalantirService> logger, HttpClient httpClient, string bearerToken, string latestPromptPalantirKey, string ontologyId, string baseUrl)
    {
        _logger = logger;
        _httpClient = httpClient;
        _bearerToken = bearerToken;
        _latestPromptPalantirKey = latestPromptPalantirKey;
        _ontologyId = ontologyId;
        _baseUrl = baseUrl;
    }

    public void UpdateLatestPromptKey(string key)
    {
        _latestPromptPalantirKey = key;
        _logger.LogInformation($"Updated latest prompt key to: {key}");
    }

    private void LogSystemMessage(string message)
    {
        _logger.LogInformation(message);
    }

    private void LogError(string message)
    {
        _logger.LogError(message);
    }

    private async Task<string> PollWorkbookForAgentPlan(string workbookId, int timeoutMinutes = 5, int pollIntervalSeconds = 5)
    {
        var startTime = DateTime.UtcNow;
        var timeoutTime = startTime.AddMinutes(timeoutMinutes);

        while (DateTime.UtcNow < timeoutTime)
        {
            try
            {
                var url = $"{_baseUrl}/v2/ontologies/{_ontologyId}/objects/ExcelWorkbook/{workbookId}";
                LogSystemMessage($"\n=== Polling Workbook for Agent Plan ===");
                LogSystemMessage($"URL: {url}");

                using var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);

                using var response = await _httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var workbookObj = JsonConvert.DeserializeObject<dynamic>(content);
                    if (workbookObj?.agentPlan != null)
                    {
                        var agentPlan = workbookObj.agentPlan.ToString();
                        if (!string.IsNullOrWhiteSpace(agentPlan))
                        {
                            LogSystemMessage($"\n=== Agent Plan Found ===\n{agentPlan}");
                            return agentPlan;
                        }
                    }
                }

                LogSystemMessage($"No agent plan found yet, waiting {pollIntervalSeconds} seconds before next attempt...");
                await Task.Delay(TimeSpan.FromSeconds(pollIntervalSeconds));
            }
            catch (Exception ex)
            {
                LogError($"\n=== Error polling for agent plan ===");
                LogError($"Error Type: {ex.GetType().Name}");
                LogError($"Error Message: {ex.Message}");
                // Continue polling despite errors
            }
        }

        throw new TimeoutException($"Timed out waiting for agent plan after {timeoutMinutes} minutes");
    }

    public async Task<string> GetAgentPlan(string workbookId)
    {
        LogSystemMessage("\n=== Starting GetAgentPlan Request ===");
        LogSystemMessage($"Latest Prompt Palantir Key: {_latestPromptPalantirKey}");
        LogSystemMessage($"Excel Workbook Primary Key: {workbookId}");

        var url = $"{_baseUrl}/v2/ontologies/{_ontologyId}/actions/get-agent-plan/apply";
        LogSystemMessage("\n=== Agent Plan Request Details ===");
        LogSystemMessage($"URL: {url}");

        var requestBody = new
        {
            parameters = new
            {
                userPrompt = _latestPromptPalantirKey,
                excelWorkbook = workbookId
            },
            options = new
            {
                mode = "VALIDATE_AND_EXECUTE",
                returnEdits = "ALL"
            }
        };

        var jsonBody = JsonConvert.SerializeObject(requestBody, Formatting.Indented);
        LogSystemMessage($"Request Body: {jsonBody}");

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using var response = await _httpClient.SendAsync(request);
            var content = await response.Content.ReadAsStringAsync();
            LogSystemMessage($"Response: {content}");

            if (!response.IsSuccessStatusCode)
            {
                LogError($"Failed to get agent plan. Status: {response.StatusCode}. Content: {content}");
                throw new HttpRequestException($"Failed to get agent plan. Status: {response.StatusCode}. Content: {content}");
            }

            // Get the agent plan by polling
            var agentPlan = await PollWorkbookForAgentPlan(workbookId);
            LogSystemMessage($"Agent plan received: {agentPlan.Length} chars");

            // Now get the plan steps
            LogSystemMessage("\n=== Getting Plan Steps ===");
            var planStepsUrl = $"{_baseUrl}/v2/ontologies/{_ontologyId}/queries/genieForExcelPlanSteps/execute";
            LogSystemMessage($"Plan Steps URL: {planStepsUrl}");

            var planStepsBody = new
            {
                parameters = new
                {
                    excelWorkbook = workbookId,
                    userPrompt = _latestPromptPalantirKey
                }
            };

            var planStepsJson = JsonConvert.SerializeObject(planStepsBody, Formatting.Indented);
            LogSystemMessage($"Plan Steps Request Body: {planStepsJson}");

            using var planStepsRequest = new HttpRequestMessage(HttpMethod.Post, planStepsUrl);
            planStepsRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
            planStepsRequest.Content = new StringContent(planStepsJson, Encoding.UTF8, "application/json");

            using var planStepsResponse = await _httpClient.SendAsync(planStepsRequest);
            var planStepsContent = await planStepsResponse.Content.ReadAsStringAsync();
            LogSystemMessage($"Plan Steps Response Status: {planStepsResponse.StatusCode}");
            LogSystemMessage($"Plan Steps Response Content: {planStepsContent}");

            if (!planStepsResponse.IsSuccessStatusCode)
            {
                LogError($"Failed to get plan steps. Status: {planStepsResponse.StatusCode}");
                LogError($"Error Content: {planStepsContent}");
            }
            else
            {
                LogSystemMessage("=== Plan Steps Retrieved Successfully ===");
            }

            return agentPlan;
        }
        catch (Exception ex)
        {
            LogError($"\n=== Error in GetAgentPlan ===");
            LogError($"Error Type: {ex.GetType().Name}");
            LogError($"Error Message: {ex.Message}");
            LogError($"Stack Trace: {ex.StackTrace}");
            throw;
        }
    }

    public async Task CreateUserPrompt(string userInput, string? customInstructions, string? selectedRange, string? selectedWorksheet)
    {
        LogSystemMessage("\n=== Starting CreateUserPrompt Request ===");
        LogSystemMessage($"User Input: {userInput}");
        LogSystemMessage($"Custom Instructions: {customInstructions ?? "None"}");
        LogSystemMessage($"Selected Range: {selectedRange ?? "None"}");
        LogSystemMessage($"Selected Worksheet: {selectedWorksheet ?? "None"}");

        var url = $"{_baseUrl}/v2/ontologies/{_ontologyId}/actions/create-user-prompt/apply";
        LogSystemMessage("\n=== User Prompt Request Details ===");
        LogSystemMessage($"URL: {url}");

        var promptGuid = Guid.NewGuid().ToString();
        var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

        var requestBody = new
        {
            parameters = new
            {
                prompt_guid = promptGuid,
                conversation_session_guidref = promptGuid,
                prompt_text = userInput,
                user_custom_instructions = customInstructions ?? "",
                user_selected_range = selectedRange ?? "",
                user_selected_worksheet = selectedWorksheet ?? "",
                version_timestamp = timestamp,
                last_mod_time = timestamp,
                user_provided_file_attachment_context = "",
                user_selected_ontology_objects = ""
            },
            options = new
            {
                mode = "VALIDATE_AND_EXECUTE",
                returnEdits = "ALL"
            }
        };

        var jsonBody = JsonConvert.SerializeObject(requestBody, Formatting.Indented);
        LogSystemMessage($"Request Body: {jsonBody}");

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using var response = await _httpClient.SendAsync(request);
            var statusCode = response.StatusCode;
            var statusDescription = response.ReasonPhrase;
            var headers = string.Join(", ", response.Headers.Select(h => $"{h.Key}: {string.Join(", ", h.Value)}"));
            var content = await response.Content.ReadAsStringAsync();

            LogSystemMessage("\n=== User Prompt Response ===");
            LogSystemMessage($"Status Code: {statusCode}");
            LogSystemMessage($"Status Description: {statusDescription}");
            LogSystemMessage($"Response Headers: {headers}");
            LogSystemMessage($"Response Content: {content}");

            if (!response.IsSuccessStatusCode)
            {
                LogError("\n=== User Prompt Request Failed ===\n");
                LogError($"Status: {statusCode}\n");
                LogError($"Reason: {statusDescription}\n");
                LogError($"Response Content: {content}\n");
                throw new HttpRequestException($"Failed to create user prompt. Status: {statusCode}. Content: {content}");
            }

            // Parse the response to get the prompt ID
            var responseObj = JsonConvert.DeserializeObject<dynamic>(content);
            LogSystemMessage($"\nParsed Response Structure:");
            if (responseObj != null)
            {
                foreach (var prop in ((Newtonsoft.Json.Linq.JObject)responseObj).Properties())
                {
                    LogSystemMessage($"Property: {prop.Name}, Type: {prop.Value?.Type}");
                }
            }

            // Try different paths to find the primary key
            string? primaryKey = null;

            if (responseObj?.edits?.edits != null)
            {
                var edits = responseObj.edits.edits;
                LogSystemMessage($"Found edits array with {edits.Count} items");
                if (edits.Count > 0)
                {
                    var firstEdit = edits[0];
                    LogSystemMessage($"First edit type: {firstEdit.type}");
                    if (firstEdit.primaryKey != null)
                    {
                        primaryKey = firstEdit.primaryKey.ToString();
                    }
                    else if (firstEdit.objectRid != null)
                    {
                        primaryKey = firstEdit.objectRid.ToString();
                    }
                }
            }
            else if (responseObj?.validation?.result != null)
            {
                LogSystemMessage($"Validation result: {responseObj.validation.result}");
                // If we have a successful validation but no edits, we might need to handle this differently
                if (responseObj.validation.result.ToString() == "VALID")
                {
                    // Use the promptGuid as a fallback
                    primaryKey = promptGuid;
                    LogSystemMessage("Using promptGuid as fallback primary key");
                }
            }

            if (primaryKey != null)
            {
                LogSystemMessage($"Extracted primary key: {primaryKey}");
                UpdateLatestPromptKey(primaryKey);
                return;
            }

            LogError("\n=== Failed to extract primary key ===");
            LogError($"Full response content for debugging:\n{content}");
            throw new InvalidOperationException("Could not extract prompt ID from response. See logs for details.");
        }
        catch (Exception ex)
        {
            LogError("\n=== Error in CreateUserPrompt ===\n");
            LogError($"Error Type: {ex.GetType().Name}\n");
            LogError($"Error Message: {ex.Message}\n");
            LogError($"Stack Trace: {ex.StackTrace}\n");
            throw;
        }
    }

    public async Task<EntityReference> CreateExcelWorkbook(Dictionary<string, string> workbookInfo)
    {
        var localGuid = Guid.NewGuid().ToString();
        var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

        try
        {
            var url = $"{_baseUrl}/v2/ontologies/{_ontologyId}/actions/create-excel-workbook/apply";
            LogSystemMessage("\n=== Create Excel Workbook Request ===");
            LogSystemMessage($"URL: {url}");

            var filePath = workbookInfo.GetValueOrDefault("workbook_path", "");
            var fileInfo = new FileInfo(filePath);

            var requestBody = new
            {
                parameters = new
                {
                    workbook_guid = localGuid,
                    active_sheet = "",  // Will be populated after workbook is loaded
                    app_guidref = localGuid,  // Using same GUID for app reference
                    author = workbookInfo.GetValueOrDefault("author", Environment.UserName),
                    category = "Excel",
                    connection_list = "[]",  // Empty JSON array as string
                    file_path = filePath,
                    file_size = fileInfo.Exists ? fileInfo.Length : 0,
                    has_password = false,
                    macro_enabled = filePath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase),
                    main_genie_excel_file_guidref = "",
                    protect_structure = false,
                    protect_windows = false,
                    read_only = false,
                    template_excel_file_guidref = "",
                    workbook_events = "[]",  // Empty JSON array as string
                    workbook_name = Path.GetFileName(filePath),
                    last_mod_time = timestamp,
                    version_timestamp = timestamp,
                    is_template = false,
                    workbook_summary = "",
                    agent_plan = ""  // Initialize with empty string
                },
                options = new
                {
                    mode = "VALIDATE_AND_EXECUTE",
                    returnEdits = "ALL"
                }
            };

            var jsonBody = JsonConvert.SerializeObject(requestBody, Formatting.Indented);
            LogSystemMessage($"Request Body: {jsonBody}");

            using var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using var response = await _httpClient.SendAsync(request);
            var content = await response.Content.ReadAsStringAsync();
            LogSystemMessage($"Response: {content}");

            if (!response.IsSuccessStatusCode)
            {
                LogError($"Failed to create workbook. Status: {response.StatusCode}. Content: {content}");
                throw new HttpRequestException($"Failed to create workbook. Status: {response.StatusCode}. Content: {content}");
            }

            var responseObj = JsonConvert.DeserializeObject<dynamic>(content);
            string? primaryKey = null;

            if (responseObj?.edits?.edits != null)
            {
                var edits = responseObj.edits.edits;
                LogSystemMessage($"Found {edits.Count} edits in response");
                if (edits.Count > 0)
                {
                    var firstEdit = edits[0];
                    LogSystemMessage($"First edit type: {firstEdit.type}");
                    if (firstEdit.primaryKey != null)
                    {
                        primaryKey = firstEdit.primaryKey.ToString();
                        LogSystemMessage($"Extracted primary key: {primaryKey}");
                    }
                }
            }

            if (primaryKey == null)
            {
                LogError("\n=== Failed to extract primary key from workbook creation response ===");
                LogError($"Full response content for debugging:\n{content}");
                throw new InvalidOperationException("Could not extract workbook primary key from response. See logs for details.");
            }

            return new EntityReference(localGuid, primaryKey);
        }
        catch (Exception ex)
        {
            LogError($"\n=== Error in CreateExcelWorkbook ===");
            LogError($"Error Type: {ex.GetType().Name}");
            LogError($"Error Message: {ex.Message}");
            LogError($"Stack Trace: {ex.StackTrace}");
            throw;
        }
    }

    public async Task<EntityReference> CreateExcelWorksheet(Excel.Worksheet worksheet, string workbookPrimaryKey)
    {
        var localGuid = Guid.NewGuid().ToString();
        var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

        try
        {
            var url = $"{_baseUrl}/v2/ontologies/{_ontologyId}/actions/create-excel-worksheet/apply";
            LogSystemMessage("\n=== Create Excel Worksheet Request ===");
            LogSystemMessage($"URL: {url}");
            LogSystemMessage($"Worksheet Name: {worksheet.Name}");
            LogSystemMessage($"Workbook Primary Key: {workbookPrimaryKey}");

            // Get worksheet properties
            var usedRange = worksheet.UsedRange;
            Excel.Window window = worksheet.Parent.Windows[1];
            var splitRow = 0;
            var splitColumn = 0;
            
            // Check if any panes are frozen
            if (window.FreezePanes || window.SplitHorizontal > 0 || window.SplitVertical > 0)
            {
                splitRow = (int)window.SplitHorizontal;
                splitColumn = (int)window.SplitVertical;
            }

            var pageSetup = worksheet.PageSetup;
            var fitToPagesTall = 1;
            var fitToPagesWide = 1;

            try
            {
                // Get page setup properties safely
                if (pageSetup != null)
                {
                    try { fitToPagesTall = pageSetup.FitToPagesTall; } catch { }
                    try { fitToPagesWide = pageSetup.FitToPagesWide; } catch { }
                }

                // Ensure valid values
                fitToPagesTall = fitToPagesTall <= 0 ? 1 : fitToPagesTall;
                fitToPagesWide = fitToPagesWide <= 0 ? 1 : fitToPagesWide;
            }
            catch (Exception ex)
            {
                // If we can't access these properties, use defaults
                LogSystemMessage($"Could not access page setup properties, using defaults. Error: {ex.Message}");
            }

            var requestBody = new
            {
                parameters = new
                {
                    worksheet_guid = localGuid,
                    display_gridlines = true,  // Default to true since we can't access this directly
                    fit_pages_tall = fitToPagesTall,
                    fit_pages_wide = fitToPagesWide,
                    freeze_left_column = splitColumn,
                    freeze_pane_split_column = splitColumn,
                    freeze_pane_split_row = splitRow,
                    freeze_top_row = splitRow,
                    h_page_breaks = "[]",  // Empty JSON array
                    outline_data = "[]",  // Empty JSON array
                    page_orientation = pageSetup?.Orientation.ToString() ?? "Portrait",
                    print_area = pageSetup?.PrintArea ?? "",
                    print_title_columns = pageSetup?.PrintTitleColumns ?? "",
                    print_title_rows = pageSetup?.PrintTitleRows ?? "",
                    @protected = worksheet.ProtectContents,
                    sheet_index = worksheet.Index,
                    sheet_name = worksheet.Name,
                    sheet_type = "Worksheet",  // Default type
                    used_range_addr = usedRange != null,
                    visible = worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible,
                    v_page_breaks = "[]",  // Empty JSON array
                    workbook_guidref = workbookPrimaryKey,
                    zoom = (int)window.Zoom,  // Cast zoom to integer
                    last_mod_time = timestamp,
                    version_timestamp = timestamp
                },
                options = new
                {
                    mode = "VALIDATE_AND_EXECUTE",
                    returnEdits = "ALL"
                }
            };

            var jsonBody = JsonConvert.SerializeObject(requestBody, Formatting.Indented);
            LogSystemMessage($"Request Body: {jsonBody}");

            using var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using var response = await _httpClient.SendAsync(request);
            var content = await response.Content.ReadAsStringAsync();
            LogSystemMessage($"Response: {content}");

            if (!response.IsSuccessStatusCode)
            {
                LogError($"Failed to create worksheet. Status: {response.StatusCode}. Content: {content}");
                throw new HttpRequestException($"Failed to create worksheet. Status: {response.StatusCode}. Content: {content}");
            }

            var responseObj = JsonConvert.DeserializeObject<dynamic>(content);
            string? primaryKey = null;

            if (responseObj?.edits?.edits != null)
            {
                var edits = responseObj.edits.edits;
                LogSystemMessage($"Found {edits.Count} edits in response");
                if (edits.Count > 0)
                {
                    var firstEdit = edits[0];
                    LogSystemMessage($"First edit type: {firstEdit.type}");
                    if (firstEdit.primaryKey != null)
                    {
                        primaryKey = firstEdit.primaryKey.ToString();
                        LogSystemMessage($"Extracted primary key: {primaryKey}");
                    }
                }
            }

            if (primaryKey == null)
            {
                LogError("\n=== Failed to extract primary key from worksheet creation response ===");
                LogError($"Full response content for debugging:\n{content}");
                throw new InvalidOperationException("Could not extract worksheet primary key from response. See logs for details.");
            }

            return new EntityReference(localGuid, primaryKey);
        }
        catch (Exception ex)
        {
            LogError($"\n=== Error in CreateExcelWorksheet ===");
            LogError($"Error Type: {ex.GetType().Name}");
            LogError($"Error Message: {ex.Message}");
            LogError($"Stack Trace: {ex.StackTrace}");
            throw;
        }
    }

    public string GetWorkbookId(string name)
    {
        // TODO: Implement actual lookup of workbook ID
        return "dummy-workbook-id";
    }

    public string GetWorksheetId(string name)
    {
        // TODO: Implement actual lookup of worksheet ID
        return "dummy-worksheet-id";
    }

    public string GetChartId(string name)
    {
        // TODO: Implement actual lookup of chart ID
        return "dummy-chart-id";
    }

    public async Task CreateExcelConnections(Excel.Workbook workbook, string workbookGuid)
    {
        // TODO: Implement actual creation of Excel connections
    }

    public async Task ProcessWorksheetCells(Excel.Worksheet worksheet, string parentWorksheetPalantirKey, string? rangeGuid = null)
    {
        _logger.LogInformation($"\n=== Processing Cells for Worksheet: {worksheet.Name} ===");
        _logger.LogInformation($"Using Parent Worksheet Palantir Key: {parentWorksheetPalantirKey}");
        if (rangeGuid != null)
        {
            _logger.LogInformation($"Using Range GUID: {rangeGuid}");
        }

        var cellBatch = new List<Dictionary<string, object>>();
        Excel.Range? nonEmptyCells = null;
        Excel.Range? usedRange = null;

        try
        {
            usedRange = worksheet.UsedRange;
            if (usedRange == null)
            {
                _logger.LogInformation("No used range found in worksheet");
                return;
            }

            try
            {
                nonEmptyCells = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants | Excel.XlCellType.xlCellTypeFormulas);
                _logger.LogInformation("Successfully retrieved non-empty cells using SpecialCells");
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"SpecialCells failed, falling back to UsedRange: {ex.Message}");
                nonEmptyCells = usedRange;
            }

            if (nonEmptyCells == null)
            {
                _logger.LogInformation("No cells to process");
                return;
            }

            int processedCells = 0;
            int skippedCells = 0;
            int batchNumber = 0;

            foreach (Excel.Range area in nonEmptyCells.Areas)
            {
                int rowStart = area.Row;
                int rowEnd = rowStart + area.Rows.Count - 1;
                int colStart = area.Column;
                int colEnd = colStart + area.Columns.Count - 1;

                for (int row = rowStart; row <= rowEnd; row++)
                {
                    for (int col = colStart; col <= colEnd; col++)
                    {
                        Excel.Range cell = area.Cells[row - rowStart + 1, col - colStart + 1];

                        try
                        {
                            if (IsCellBlankOrDefault(cell))
                            {
                                skippedCells++;
                                continue;
                            }

                            var cellData = CreateCellData(cell, parentWorksheetPalantirKey, rangeGuid);

                            if (!_cellValidationDone && cellBatch.Count == 0)
                            {
                                await ValidateFirstCell(cellData);
                                _cellValidationDone = true;
                            }

                            cellBatch.Add(cellData);
                            processedCells++;

                            // Send batch if we hit the size limit
                            if (cellBatch.Count >= BATCH_SIZE)
                            {
                                batchNumber++;
                                _logger.LogInformation($"Sending batch #{batchNumber} ({cellBatch.Count} cells)");

                                try
                                {
                                    await SendCellBatch(cellBatch);
                                    _logger.LogInformation($"Successfully sent batch #{batchNumber}");
                                }
                                catch (Exception batchEx)
                                {
                                    _logger.LogError($"Failed to send batch #{batchNumber}");
                                    _logger.LogError($"Batch error details: {batchEx.Message}");
                                    throw;
                                }
                                finally
                                {
                                    cellBatch.Clear();
                                }
                            }
                        }
                        catch (Exception cellEx)
                        {
                            _logger.LogError($"Error processing cell at Row {row}, Column {col}: {cellEx.Message}");
                            // Don't throw here - continue processing other cells
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(cell);
                        }
                    }

                    if ((row - rowStart) % 100 == 0 || row == rowEnd)
                    {
                        _logger.LogInformation($"Progress update - Processed {row - rowStart + 1}/{rowEnd - rowStart + 1} rows in current area");
                        _logger.LogInformation($"Stats: Processed cells: {processedCells}, Skipped cells: {skippedCells}, Batches sent: {batchNumber}");
                    }
                }
            }

            // Send any remaining cells in the final batch
            if (cellBatch.Count > 0)
            {
                batchNumber++;
                _logger.LogInformation($"Sending final batch #{batchNumber} ({cellBatch.Count} cells)");
                try
                {
                    await SendCellBatch(cellBatch);
                    _logger.LogInformation($"Successfully sent final batch #{batchNumber}");
                }
                catch (Exception batchEx)
                {
                    _logger.LogError($"Failed to send final batch #{batchNumber}");
                    _logger.LogError($"Batch error details: {batchEx.Message}");
                    throw;
                }
            }

            _logger.LogInformation("\n=== Worksheet Processing Summary ===");
            _logger.LogInformation($"Total cells processed: {processedCells}");
            _logger.LogInformation($"Total cells skipped: {skippedCells}");
            _logger.LogInformation($"Total batches sent: {batchNumber}");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Critical error processing worksheet {worksheet.Name}");
            _logger.LogError($"Error message: {ex.Message}");
            _logger.LogError($"Stack trace: {ex.StackTrace}");

            // Try to send any remaining cells before throwing
            if (cellBatch.Count > 0)
            {
                _logger.LogInformation($"Attempting to send {cellBatch.Count} remaining cells before exiting...");
                try
                {
                    await SendCellBatch(cellBatch);
                    _logger.LogInformation("Successfully sent remaining cells");
                }
                catch (Exception finalBatchEx)
                {
                    _logger.LogError($"Failed to send remaining cells: {finalBatchEx.Message}");
                }
            }

            throw;
        }
        finally
        {
            if (usedRange != null && usedRange != nonEmptyCells)
            {
                Marshal.ReleaseComObject(usedRange);
            }
            if (nonEmptyCells != null)
            {
                Marshal.ReleaseComObject(nonEmptyCells);
            }
        }
    }

    public async Task ProcessWorksheetCharts(Excel.Worksheet worksheet, string parentWorksheetPalantirKey)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        _logger.LogInformation($"\n=== Processing Charts for Worksheet: {worksheet.Name} ===");
        _logger.LogInformation($"Parent Worksheet Palantir Key: {parentWorksheetPalantirKey}");

        try
        {
            Excel.ChartObjects? chartObjects = null;
            try
            {
                chartObjects = worksheet.ChartObjects() as Excel.ChartObjects;
                if (chartObjects == null || chartObjects.Count == 0)
                {
                    _logger.LogInformation("No charts found in worksheet");
                    return;
                }

                _logger.LogInformation($"Found {chartObjects.Count} charts to process");

                foreach (Excel.ChartObject chartObject in chartObjects)
                {
                    try
                    {
                        await ProcessChart(chartObject, parentWorksheetPalantirKey);
                    }
                    catch (Exception chartEx)
                    {
                        _logger.LogError($"Error processing chart {chartObject.Name}: {chartEx.Message}");
                        continue;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(chartObject);
                    }
                }
            }
            finally
            {
                if (chartObjects != null)
                {
                    Marshal.ReleaseComObject(chartObjects);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Critical error processing charts in worksheet {worksheet.Name}", ex);
            throw;
        }
    }

    private async Task ProcessChart(Excel.ChartObject chartObject, string parentWorksheetPalantirKey)
    {
        if (chartObject == null)
        {
            throw new ArgumentNullException(nameof(chartObject));
        }

        Excel.Chart? chart = null;
        try
        {
            chart = chartObject.Chart;
            if (chart == null)
            {
                throw new InvalidOperationException("Chart object is null");
            }

            var chartGuid = Guid.NewGuid().ToString();
            var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

            _logger.LogInformation($"\nProcessing chart: {chartObject.Name}");

            var parameters = new Dictionary<string, object>
            {
                ["chart_guid"] = chartGuid,
                ["worksheet_guidref"] = parentWorksheetPalantirKey,
                ["chart_name"] = chartObject.Name ?? "Unnamed Chart",
                ["chart_type"] = chart.ChartType.ToString(),
                ["version_timestamp"] = timestamp,
                ["last_mod_time"] = timestamp,
                ["left"] = chartObject.Left,
                ["top"] = chartObject.Top,
                ["width"] = chartObject.Width,
                ["height"] = chartObject.Height
            };

            try
            {
                if (chart.ChartArea != null)
                {
                    parameters["chart_area_border_color"] = GetColorString(chart.ChartArea.Border?.Color) ?? "#000000";
                    parameters["chart_area_border_style"] = ConvertBorderStyle((Excel.XlLineStyle)(chart.ChartArea.Border?.LineStyle ?? Excel.XlLineStyle.xlContinuous));
                    parameters["chart_area_border_weight"] = chart.ChartArea.Border?.Weight ?? 1.0;
                    parameters["chart_area_fill_color"] = GetColorString(chart.ChartArea.Interior?.Color) ?? "#FFFFFF";
                    parameters["chart_area_fill_pattern"] = ConvertInteriorPattern((Excel.XlPattern)(chart.ChartArea.Interior?.Pattern ?? Excel.XlPattern.xlPatternSolid));
                    parameters["chart_area_fill_pattern_color"] = GetColorString(chart.ChartArea.Interior?.PatternColor) ?? "#000000";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting chart area properties: {ex.Message}");
                parameters["chart_area_border_color"] = "#000000";
                parameters["chart_area_border_style"] = "Continuous";
                parameters["chart_area_border_weight"] = 1.0;
                parameters["chart_area_fill_color"] = "#FFFFFF";
                parameters["chart_area_fill_pattern"] = "Solid";
                parameters["chart_area_fill_pattern_color"] = "#000000";
            }

            try
            {
                if (chart.PlotArea != null)
                {
                    parameters["plot_area_border_color"] = GetColorString(chart.PlotArea.Border?.Color) ?? "#000000";
                    parameters["plot_area_border_style"] = ConvertBorderStyle((Excel.XlLineStyle)(chart.PlotArea.Border?.LineStyle ?? Excel.XlLineStyle.xlContinuous));
                    parameters["plot_area_border_weight"] = chart.PlotArea.Border?.Weight ?? 1.0;
                    parameters["plot_area_fill_color"] = GetColorString(chart.PlotArea.Interior?.Color) ?? "#FFFFFF";
                    parameters["plot_area_fill_pattern"] = ConvertInteriorPattern((Excel.XlPattern)(chart.PlotArea.Interior?.Pattern ?? Excel.XlPattern.xlPatternSolid));
                    parameters["plot_area_fill_pattern_color"] = GetColorString(chart.PlotArea.Interior?.PatternColor) ?? "#000000";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting plot area properties: {ex.Message}");
                parameters["plot_area_border_color"] = "#000000";
                parameters["plot_area_border_style"] = "Continuous";
                parameters["plot_area_border_weight"] = 1.0;
                parameters["plot_area_fill_color"] = "#FFFFFF";
                parameters["plot_area_fill_pattern"] = "Solid";
                parameters["plot_area_fill_pattern_color"] = "#000000";
            }

            try
            {
                parameters["chart_is3d"] = false;
                parameters["chart_rotation"] = chart.Rotation;
                parameters["chart_elevation"] = chart.Elevation;
                parameters["chart_depth_percent"] = chart.DepthPercent;
                parameters["chart_height_percent"] = chart.HeightPercent;
                parameters["chart_perspective"] = "0";
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting 3D properties: {ex.Message}");
                parameters["chart_is3d"] = false;
                parameters["chart_rotation"] = 0;
                parameters["chart_elevation"] = 0;
                parameters["chart_depth_percent"] = 100;
                parameters["chart_height_percent"] = 100;
                parameters["chart_perspective"] = "0";
            }

            try
            {
                parameters["chart_gap_width"] = 150;
                parameters["chart_overlap"] = 0;
                parameters["chart_style"] = chart.ChartStyle?.ToString() ?? "Default";
                parameters["plot_by"] = chart.PlotBy.ToString();
                parameters["chart_location_type"] = "Embedded";
                parameters["data_source_linked"] = chart.HasDataTable ? "Yes" : "No";
                parameters["pivot_chart"] = "No";
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting style properties: {ex.Message}");
                parameters["chart_gap_width"] = 150;
                parameters["chart_overlap"] = 0;
                parameters["chart_style"] = "Default";
                parameters["plot_by"] = "Columns";
                parameters["chart_location_type"] = "Embedded";
                parameters["data_source_linked"] = "No";
                parameters["pivot_chart"] = "No";
            }

            try
            {
                if (chart.HasTitle)
                {
                    var title = chart.ChartTitle;
                    parameters["chart_title_text"] = title.Text;
                    parameters["chart_title_visible"] = true;
                    parameters["chart_title_font_name"] = title.Font?.Name ?? "Calibri";
                    parameters["chart_title_font_size"] = title.Font?.Size ?? 11.0;
                    parameters["chart_title_font_color"] = GetColorString(title.Font?.Color) ?? "#000000";
                    parameters["chart_title_bold"] = title.Font?.Bold ?? false;
                    parameters["chart_title_italic"] = title.Font?.Italic ?? false;
                    parameters["chart_title_underline"] = (title.Font?.Underline ?? 0) != 0;
                }
                else
                {
                    parameters["chart_title_text"] = "";
                    parameters["chart_title_visible"] = false;
                    parameters["chart_title_font_name"] = "Calibri";
                    parameters["chart_title_font_size"] = 11.0;
                    parameters["chart_title_font_color"] = "#000000";
                    parameters["chart_title_bold"] = false;
                    parameters["chart_title_italic"] = false;
                    parameters["chart_title_underline"] = false;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting title properties: {ex.Message}");
                parameters["chart_title_text"] = "";
                parameters["chart_title_visible"] = false;
                parameters["chart_title_font_name"] = "Calibri";
                parameters["chart_title_font_size"] = 11.0;
                parameters["chart_title_font_color"] = "#000000";
                parameters["chart_title_bold"] = false;
                parameters["chart_title_italic"] = false;
                parameters["chart_title_underline"] = false;
            }

            try
            {
                if (chart.HasLegend)
                {
                    var legend = chart.Legend;
                    if (legend != null)
                    {
                        parameters["has_legend"] = true;
                        parameters["legend_position"] = legend.Position.ToString();
                        parameters["legend_visible"] = true;
                        parameters["legend_font_name"] = legend.Font?.Name ?? "Calibri";
                        parameters["legend_font_size"] = legend.Font?.Size ?? 11.0;
                        parameters["legend_font_color"] = GetColorString(legend.Font?.Color) ?? "#000000";
                        parameters["legend_bold"] = legend.Font?.Bold ?? false;
                        parameters["legend_italic"] = legend.Font?.Italic ?? false;
                        parameters["legend_underline"] = (legend.Font?.Underline ?? 0) != 0;
                        parameters["legend_background_color"] = GetColorString(legend.Interior?.Color) ?? "#FFFFFF";
                        parameters["legend_background_pattern"] = ConvertInteriorPattern((Excel.XlPattern)(legend.Interior?.Pattern ?? Excel.XlPattern.xlPatternSolid));
                        parameters["legend_background_pattern_color"] = GetColorString(legend.Interior?.PatternColor) ?? "#000000";
                        parameters["legend_border_color"] = GetColorString(legend.Border?.Color) ?? "#000000";
                        parameters["legend_border_style"] = ConvertBorderStyle((Excel.XlLineStyle)(legend.Border?.LineStyle ?? Excel.XlLineStyle.xlContinuous));
                        parameters["legend_border_weight"] = legend.Border?.Weight ?? 1.0;
                    }
                }
                else
                {
                    parameters["has_legend"] = false;
                    parameters["legend_position"] = "None";
                    parameters["legend_visible"] = false;
                    parameters["legend_font_name"] = "Calibri";
                    parameters["legend_font_size"] = 11.0;
                    parameters["legend_font_color"] = "#000000";
                    parameters["legend_bold"] = false;
                    parameters["legend_italic"] = false;
                    parameters["legend_underline"] = false;
                    parameters["legend_background_color"] = "#FFFFFF";
                    parameters["legend_background_pattern"] = "None";
                    parameters["legend_background_pattern_color"] = "#000000";
                    parameters["legend_border_color"] = "#000000";
                    parameters["legend_border_style"] = "None";
                    parameters["legend_border_weight"] = 1.0;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting legend properties: {ex.Message}");
                parameters["has_legend"] = false;
                parameters["legend_position"] = "None";
                parameters["legend_visible"] = false;
                parameters["legend_font_name"] = "Calibri";
                parameters["legend_font_size"] = 11.0;
                parameters["legend_font_color"] = "#000000";
                parameters["legend_bold"] = false;
                parameters["legend_italic"] = false;
                parameters["legend_underline"] = false;
                parameters["legend_background_color"] = "#FFFFFF";
                parameters["legend_background_pattern"] = "None";
                parameters["legend_background_pattern_color"] = "#000000";
                parameters["legend_border_color"] = "#000000";
                parameters["legend_border_style"] = "None";
                parameters["legend_border_weight"] = 1.0;
            }

            Excel.Axis? xAxis = null;
            try
            {
                if (chart.HasAxis(Excel.XlAxisType.xlCategory))
                {
                    xAxis = chart.Axes(Excel.XlAxisType.xlCategory);
                    if (xAxis != null)
                    {
                        foreach (var prop in GetAxisProperties(xAxis, "x_axis_"))
                        {
                            parameters[prop.Key] = prop.Value;
                        }
                    }
                    else
                    {
                        AddDefaultAxisProperties(parameters, "x_axis_");
                    }
                }
                else
                {
                    AddDefaultAxisProperties(parameters, "x_axis_");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting X-axis properties: {ex.Message}");
                AddDefaultAxisProperties(parameters, "x_axis_");
            }
            finally
            {
                if (xAxis != null)
                {
                    Marshal.ReleaseComObject(xAxis);
                }
            }

            Excel.Axis? yAxis = null;
            try
            {
                if (chart.HasAxis(Excel.XlAxisType.xlValue))
                {
                    yAxis = chart.Axes(Excel.XlAxisType.xlValue);
                    if (yAxis != null)
                    {
                        foreach (var prop in GetAxisProperties(yAxis, "y_axis_"))
                        {
                            parameters[prop.Key] = prop.Value;
                        }
                    }
                    else
                    {
                        AddDefaultAxisProperties(parameters, "y_axis_");
                    }
                }
                else
                {
                    AddDefaultAxisProperties(parameters, "y_axis_");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error getting Y-axis properties: {ex.Message}");
                AddDefaultAxisProperties(parameters, "y_axis_");
            }
            finally
            {
                if (yAxis != null)
                {
                    Marshal.ReleaseComObject(yAxis);
                }
            }

            if (!_chartValidationDone)
            {
                await ValidateChartParameters(parameters);
                _chartValidationDone = true;
            }

            var requestBody = new { parameters, options = new { returnEdits = "ALL" } };
            var url = $"{_baseUrl}/ontologies/{_ontologyId}/actions/create-excel-chart/apply";
            var json = System.Text.Json.JsonSerializer.Serialize(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            _logger.LogInformation("Sending chart creation request to Palantir...");
            var response = await _httpClient.PostAsync(url, content);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"Failed to create chart. Status: {response.StatusCode}, Response: {responseContent}");
            }

            _logger.LogInformation($"Successfully created chart {chartObject.Name} in Palantir");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error processing chart {chartObject.Name}", ex);
            throw;
        }
        finally
        {
            if (chart != null)
            {
                Marshal.ReleaseComObject(chart);
            }
        }
    }

    private async Task ValidateFirstCell(Dictionary<string, object> cellData)
    {
        var validateRequestBody = new {
            parameters = cellData,
            options = new {
                mode = "VALIDATE_ONLY",
                returnEdits = "NONE"
            }
        };

        var validateUrl = $"{_baseUrl}/v2/ontologies/{_ontologyId}/actions/create-excel-cell/apply";
        _logger.LogInformation("Performing one-time cell validation...");
        _logger.LogInformation($"Cell details - Row: {cellData["row_index"]}, Column: {cellData["col_index"]}, Address: {cellData["cell_address"]}");

        var jsonBody = System.Text.Json.JsonSerializer.Serialize(validateRequestBody);
        using var request = new HttpRequestMessage(HttpMethod.Post, validateUrl);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
        request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
        
        var validateResponse = await _httpClient.SendAsync(request);
        var validateContent = await validateResponse.Content.ReadAsStringAsync();

        _logger.LogInformation($"Validation Response Status: {validateResponse.StatusCode}");
        _logger.LogInformation($"Validation Response Content: {validateContent}");

        if (!validateResponse.IsSuccessStatusCode)
        {
            _logger.LogError("Cell validation failed");
            _logger.LogError($"Request URL: {validateUrl}");
            _logger.LogError($"Request Body: {jsonBody}");
            _logger.LogError($"Cell data: {System.Text.Json.JsonSerializer.Serialize(cellData, new System.Text.Json.JsonSerializerOptions { WriteIndented = true })}");
            throw new HttpRequestException($"Cell validation failed. Status: {validateResponse.StatusCode}, Response: {validateContent}");
        }

        _logger.LogInformation("One-time cell validation successful");
    }

    private async Task SendCellBatch(List<Dictionary<string, object>> cellBatch)
    {
        var requestBody = new {
            requests = cellBatch.Select(parameters => new { parameters }).ToList(),
            options = new { 
                returnEdits = "NONE" 
            }
        };

        var requestUrl = $"{_baseUrl}/v2/ontologies/{_ontologyId}/actions/create-excel-cell/applyBatch";
        try
        {
            _logger.LogInformation($"Sending batch of {cellBatch.Count} cells to {requestUrl}");
            _logger.LogInformation($"First cell in batch: Row {cellBatch[0]["row_index"]}, Column {cellBatch[0]["col_index"]}, Address {cellBatch[0]["cell_address"]}");

            var jsonBody = System.Text.Json.JsonSerializer.Serialize(requestBody);
            using var request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            
            var response = await _httpClient.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            _logger.LogInformation($"Response Status: {response.StatusCode}");
            _logger.LogInformation($"Response Content: {responseContent}");

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError($"Failed to create cell batch. Status: {response.StatusCode}");
                _logger.LogError($"Request URL: {requestUrl}");
                _logger.LogError($"Request Body: {jsonBody}");
                _logger.LogError($"Response Content: {responseContent}");
                throw new HttpRequestException($"Failed to create Excel cell batch. Status: {response.StatusCode}, Response: {responseContent}");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error sending cell batch: {ex.Message}");
            _logger.LogError($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    private Dictionary<string, object> CreateCellData(Excel.Range cell, string parentWorksheetPalantirKey, string? rangeGuid = null)
    {
        var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        string localCellGuid = Guid.NewGuid().ToString();

        string horizontalAlignment = ConvertHorizontalAlignment((Excel.XlHAlign)cell.HorizontalAlignment);
        string verticalAlignment = ConvertVerticalAlignment((Excel.XlVAlign)cell.VerticalAlignment);

        var topBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeTop];
        var bottomBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom];
        var leftBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeLeft];
        var rightBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeRight];

        bool isUnderlined = false;
        try
        {
            var underlineStyle = cell.Font.Underline;
            isUnderlined = (underlineStyle != null && (int)underlineStyle != (int)Excel.XlUnderlineStyle.xlUnderlineStyleNone);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error checking underline style for cell {cell.Address}: {ex.Message}");
        }

        var data = new Dictionary<string, object>
        {
            ["cell_guid"] = localCellGuid,
            ["worksheet_guidref"] = parentWorksheetPalantirKey,
            ["range_guidref"] = rangeGuid ?? "no-range",
            ["version_timestamp"] = timestamp,
            ["last_mod_time"] = timestamp,
            ["row_index"] = cell.Row,
            ["col_index"] = cell.Column,
            ["cell_address"] = cell.Address ?? "",
            ["cell_value"] = cell.Text ?? "",
            ["formula"] = cell.Formula ?? "",
            ["error_value"] = "",
            ["font_name"] = cell.Font.Name ?? "Calibri",
            ["font_size"] = ConvertToDouble(cell.Font.Size) > 0 ? ConvertToDouble(cell.Font.Size) : 11.0,
            ["font_color"] = GetColorString(cell.Font.Color) ?? "#000000",
            ["bold"] = ConvertToBoolean(cell.Font.Bold),
            ["italic"] = ConvertToBoolean(cell.Font.Italic),
            ["underline"] = isUnderlined,
            ["strikethrough"] = ConvertToBoolean(cell.Font.Strikethrough),
            ["subscript"] = ConvertToBoolean(cell.Font.Subscript),
            ["superscript"] = ConvertToBoolean(cell.Font.Superscript),
            ["border_top_color"] = GetColorString(topBorder.Color) ?? "#000000",
            ["border_top_style"] = ConvertBorderStyle((Excel.XlLineStyle)topBorder.LineStyle),
            ["border_top_weight"] = ConvertBorderWeight((Excel.XlBorderWeight)topBorder.Weight),
            ["border_bottom_color"] = GetColorString(bottomBorder.Color) ?? "#000000",
            ["border_bottom_style"] = ConvertBorderStyle((Excel.XlLineStyle)bottomBorder.LineStyle),
            ["border_bottom_weight"] = ConvertBorderWeight((Excel.XlBorderWeight)bottomBorder.Weight),
            ["border_left_color"] = GetColorString(leftBorder.Color) ?? "#000000",
            ["border_left_style"] = ConvertBorderStyle((Excel.XlLineStyle)leftBorder.LineStyle),
            ["border_left_weight"] = ConvertBorderWeight((Excel.XlBorderWeight)leftBorder.Weight),
            ["border_right_color"] = GetColorString(rightBorder.Color) ?? "#000000",
            ["border_right_style"] = ConvertBorderStyle((Excel.XlLineStyle)rightBorder.LineStyle),
            ["border_right_weight"] = ConvertBorderWeight((Excel.XlBorderWeight)rightBorder.Weight),
            ["bg_color"] = GetColorString(cell.Interior.Color) ?? "#FFFFFF",
            ["interior_pattern"] = ConvertInteriorPattern((Excel.XlPattern)cell.Interior.Pattern),
            ["interior_pattern_color"] = GetColorString(cell.Interior.PatternColor) ?? "#000000",
            ["horizontal_alignment"] = horizontalAlignment,
            ["vertical_alignment"] = verticalAlignment,
            ["indent_level"] = ConvertToInt(cell.IndentLevel),
            ["shrink_to_fit"] = ConvertToBoolean(cell.ShrinkToFit),
            ["wrap_text"] = ConvertToBoolean(cell.WrapText),
            ["locked"] = ConvertToBoolean(cell.Locked),
            ["number_format"] = cell.NumberFormat ?? "General",
            ["orientation"] = ConvertToInt(cell.Orientation)
        };

        return data;
    }

    private string ConvertHorizontalAlignment(Excel.XlHAlign align)
    {
        return align switch
        {
            Excel.XlHAlign.xlHAlignLeft => "Left",
            Excel.XlHAlign.xlHAlignCenter => "Center",
            Excel.XlHAlign.xlHAlignRight => "Right",
            Excel.XlHAlign.xlHAlignJustify => "Justify",
            Excel.XlHAlign.xlHAlignDistributed => "Distributed",
            Excel.XlHAlign.xlHAlignFill => "Fill",
            Excel.XlHAlign.xlHAlignCenterAcrossSelection => "CenterAcrossSelection",
            _ => "General"
        };
    }

    private string ConvertVerticalAlignment(Excel.XlVAlign align)
    {
        return align switch
        {
            Excel.XlVAlign.xlVAlignTop => "Top",
            Excel.XlVAlign.xlVAlignCenter => "Center",
            Excel.XlVAlign.xlVAlignBottom => "Bottom",
            Excel.XlVAlign.xlVAlignJustify => "Justify",
            Excel.XlVAlign.xlVAlignDistributed => "Distributed",
            _ => "Bottom"
        };
    }

    private string ConvertBorderStyle(Excel.XlLineStyle style)
    {
        return style switch
        {
            Excel.XlLineStyle.xlContinuous => "Continuous",
            Excel.XlLineStyle.xlDash => "Dash",
            Excel.XlLineStyle.xlDashDot => "DashDot",
            Excel.XlLineStyle.xlDashDotDot => "DashDotDot",
            Excel.XlLineStyle.xlDot => "Dot",
            Excel.XlLineStyle.xlDouble => "Double",
            Excel.XlLineStyle.xlSlantDashDot => "SlantDashDot",
            Excel.XlLineStyle.xlLineStyleNone => "None",
            _ => "None"
        };
    }

    private string ConvertBorderWeight(Excel.XlBorderWeight weight)
    {
        return weight switch
        {
            Excel.XlBorderWeight.xlHairline => "Hairline",
            Excel.XlBorderWeight.xlThin => "Thin",
            Excel.XlBorderWeight.xlMedium => "Medium",
            Excel.XlBorderWeight.xlThick => "Thick",
            _ => "Thin"
        };
    }

    private string ConvertInteriorPattern(Excel.XlPattern pattern)
    {
        return pattern switch
        {
            Excel.XlPattern.xlPatternAutomatic => "Automatic",
            Excel.XlPattern.xlPatternChecker => "Checker",
            Excel.XlPattern.xlPatternCrissCross => "CrissCross",
            Excel.XlPattern.xlPatternDown => "Down",
            Excel.XlPattern.xlPatternGray16 => "Gray16",
            Excel.XlPattern.xlPatternGray25 => "Gray25",
            Excel.XlPattern.xlPatternGray50 => "Gray50",
            Excel.XlPattern.xlPatternGray75 => "Gray75",
            Excel.XlPattern.xlPatternGray8 => "Gray8",
            Excel.XlPattern.xlPatternGrid => "Grid",
            Excel.XlPattern.xlPatternHorizontal => "Horizontal",
            Excel.XlPattern.xlPatternLightDown => "LightDown",
            Excel.XlPattern.xlPatternLightHorizontal => "LightHorizontal",
            Excel.XlPattern.xlPatternLightUp => "LightUp",
            Excel.XlPattern.xlPatternLightVertical => "LightVertical",
            Excel.XlPattern.xlPatternNone => "None",
            Excel.XlPattern.xlPatternSemiGray75 => "SemiGray75",
            Excel.XlPattern.xlPatternSolid => "Solid",
            Excel.XlPattern.xlPatternUp => "Up",
            Excel.XlPattern.xlPatternVertical => "Vertical",
            _ => "None"
        };
    }

    private string GetColorString(dynamic color)
    {
        try
        {
            if (color == null || color == DBNull.Value) return "#000000";

            if (color < 0)
            {
                return "#000000";
            }

            string hex = color.ToString("X6");
            if (hex.Length > 6) hex = hex.Substring(hex.Length - 6);
            if (hex.Length < 6) hex = hex.PadLeft(6, '0');

            return $"#{hex}";
        }
        catch
        {
            return "#000000";
        }
    }

    private bool ConvertToBoolean(object? value)
    {
        if (value == null || value == DBNull.Value)
        {
            return false;
        }
        return Convert.ToBoolean(value);
    }

    private string ConvertToString(object? value)
    {
        if (value == null || value == DBNull.Value)
            return "";
        return value.ToString() ?? "";
    }

    private int ConvertToInt(object? value)
    {
        if (value == null || value == DBNull.Value)
            return 0;
        return Convert.ToInt32(value);
    }

    private double ConvertToDouble(object? value)
    {
        if (value == null || value == DBNull.Value)
            return 0.0;
        return Convert.ToDouble(value);
    }

    private bool IsCellBlankOrDefault(Excel.Range cell)
    {
        if (cell == null)
        {
            return true;
        }

        try
        {
            string cellText = ConvertToString(cell.Text);
            string cellFormula = ConvertToString(cell.Formula);
            bool hasNoValue = string.IsNullOrEmpty(cellText?.Trim());
            bool hasNoFormula = string.IsNullOrEmpty(cellFormula?.Trim());

            var font = cell.Font;
            bool isBold = font != null && ConvertToBoolean(font.Bold);
            bool isItalic = font != null && ConvertToBoolean(font.Italic);
            bool hasStrikethrough = font != null && ConvertToBoolean(font.Strikethrough);

            bool isUnderlined = false;
            try
            {
                if (font != null)
                {
                    var underline = font.Underline;
                    if (underline != null && underline != DBNull.Value)
                    {
                        Excel.XlUnderlineStyle style = (Excel.XlUnderlineStyle)underline;
                        isUnderlined = (style != Excel.XlUnderlineStyle.xlUnderlineStyleNone);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Warning: Could not get underline style: {ex.Message}");
            }

            bool hasFill = false;
            try
            {
                var interior = cell.Interior;
                if (interior != null)
                {
                    var interiorColor = interior.Color;
                    hasFill = interiorColor != null && interiorColor != DBNull.Value && Convert.ToDouble(interiorColor) != 16777215;
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Warning: Could not get interior color: {ex.Message}");
            }

            if (hasNoValue && hasNoFormula && !isBold && !isItalic && !isUnderlined && !hasStrikethrough && !hasFill)
            {
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error checking if cell is blank: {ex.Message}");
            return false;
        }
    }

    private Dictionary<string, object> GetAxisProperties(Excel.Axis axis, string prefix)
    {
        var props = new Dictionary<string, object>
        {
            [$"{prefix}crosses"] = axis.Crosses.ToString(),
            [$"{prefix}crosses_at"] = axis.CrossesAt,
            [$"{prefix}line_color"] = GetColorString(axis.Border.Color) ?? "#000000",
            [$"{prefix}line_style"] = ConvertBorderStyle((Excel.XlLineStyle)axis.Border.LineStyle),
            [$"{prefix}line_weight"] = axis.Border.Weight,
            [$"{prefix}major_unit"] = axis.MajorUnit,
            [$"{prefix}max"] = axis.MaximumScale,
            [$"{prefix}min"] = axis.MinimumScale,
            [$"{prefix}minor_unit"] = axis.MinorUnit,
            [$"{prefix}scale_type"] = axis.ScaleType.ToString(),
            [$"{prefix}tick_label_font_name"] = axis.TickLabels.Font.Name,
            [$"{prefix}tick_label_font_size"] = axis.TickLabels.Font.Size,
            [$"{prefix}tick_label_orientation"] = (int)axis.TickLabelPosition,
            [$"{prefix}tick_label_position"] = axis.TickLabelPosition.ToString(),
            [$"{prefix}title"] = "",
            [$"{prefix}title_bold"] = false,
            [$"{prefix}title_font_color"] = "#000000",
            [$"{prefix}title_font_name"] = "Calibri",
            [$"{prefix}title_font_size"] = 11.0,
            [$"{prefix}title_italic"] = false,
            [$"{prefix}title_underline"] = false
        };

        return props;
    }

    private void AddDefaultAxisProperties(Dictionary<string, object> parameters, string prefix)
    {
        parameters.Add($"{prefix}crosses", "AutoZero");
        parameters.Add($"{prefix}crosses_at", 0.0);
        parameters.Add($"{prefix}line_color", "#000000");
        parameters.Add($"{prefix}line_style", "Continuous");
        parameters.Add($"{prefix}line_weight", 1.0);
        parameters.Add($"{prefix}major_unit", 0.0);
        parameters.Add($"{prefix}max", 0.0);
        parameters.Add($"{prefix}min", 0.0);
        parameters.Add($"{prefix}minor_unit", 0.0);
        parameters.Add($"{prefix}scale_type", "Linear");
        parameters.Add($"{prefix}tick_label_font_name", "Calibri");
        parameters.Add($"{prefix}tick_label_font_size", 11.0);
        parameters.Add($"{prefix}tick_label_orientation", 0);
        parameters.Add($"{prefix}tick_label_position", "NextToAxis");
        parameters.Add($"{prefix}title", "");
        parameters.Add($"{prefix}title_bold", false);
        parameters.Add($"{prefix}title_font_color", "#000000");
        parameters.Add($"{prefix}title_font_name", "Calibri");
        parameters.Add($"{prefix}title_font_size", 11.0);
        parameters.Add($"{prefix}title_italic", false);
        parameters.Add($"{prefix}title_underline", false);
    }

    private async Task ValidateChartParameters(Dictionary<string, object> parameters)
    {
        var validateRequestBody = new {
            parameters,
            options = new {
                mode = "VALIDATE_ONLY",
                returnEdits = "NONE"
            }
        };

        var url = $"{_baseUrl}/ontologies/{_ontologyId}/actions/create-excel-chart/apply";
        var validateJson = System.Text.Json.JsonSerializer.Serialize(validateRequestBody);
        var validateContent = new StringContent(validateJson, Encoding.UTF8, "application/json");

        _logger.LogInformation("Validating chart parameters...");
        var validateResponse = await _httpClient.PostAsync(url, validateContent);
        var validateResponseContent = await validateResponse.Content.ReadAsStringAsync();

        if (!validateResponse.IsSuccessStatusCode)
        {
            _logger.LogError("Chart parameter validation failed");
            _logger.LogError($"Parameters: {validateJson}");
            _logger.LogError($"Response: {validateResponseContent}");
            throw new HttpRequestException($"Chart validation failed. Status: {validateResponse.StatusCode}");
        }

        _logger.LogInformation("Chart parameter validation successful");
    }

    public async Task GenerateWorkbookSummary(string workbookPrimaryKey)
    {
        _logger.LogInformation("\n=== Starting Workbook Summary Generation ===");
        _logger.LogInformation($"Workbook Primary Key: {workbookPrimaryKey}");
        _logger.LogInformation($"Base URL: {_baseUrl}");

        try
        {
            if (string.IsNullOrEmpty(workbookPrimaryKey))
            {
                throw new ArgumentException("Workbook primary key cannot be null or empty");
            }

            // Make URL consistent with other endpoints by adding /v2
            var generateUrl = $"{_baseUrl}/v2/ontologies/{_ontologyId}/actions/generate-workbook-summary/apply";
            _logger.LogInformation($"Request URL: {generateUrl}");

            var requestBody = new
            {
                parameters = new
                {
                    excelWorkbook = workbookPrimaryKey
                },
                options = new
                {
                    mode = "VALIDATE_AND_EXECUTE",
                    returnEdits = "ALL"
                }
            };

            var jsonBody = JsonConvert.SerializeObject(requestBody);
            _logger.LogInformation($"Request Body: {jsonBody}");

            using var request = new HttpRequestMessage(HttpMethod.Post, generateUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _bearerToken);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            _logger.LogInformation("Sending summary generation request...");
            using var response = await _httpClient.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            _logger.LogInformation("\n=== Summary Generation Response ===");
            _logger.LogInformation($"Status: {response.StatusCode}");
            _logger.LogInformation($"Response Content: {responseContent}");

            if (!response.IsSuccessStatusCode)
            {
                if (responseContent.Contains("Logic function not executable"))
                {
                    _logger.LogError("The workbook summary logic function is not executable. Please check your Palantir Foundry configuration and ensure that workbook summary generation is enabled for this ontology.");
                }
                _logger.LogError($"Failed to generate workbook summary");
                _logger.LogError($"Status Code: {response.StatusCode}");
                _logger.LogError($"Response Content: {responseContent}");
                _logger.LogError($"Request URL: {generateUrl}");
                _logger.LogError($"Request Body: {jsonBody}");
                throw new HttpRequestException($"Failed to generate workbook summary. Status: {response.StatusCode}. Response: {responseContent}");
            }

            // Parse response to verify success
            var responseObj = JsonConvert.DeserializeObject<dynamic>(responseContent);
            if (responseObj?.validation?.result?.ToString() != "VALID")
            {
                _logger.LogError("Summary generation validation failed");
                _logger.LogError($"Validation Result: {responseObj?.validation?.result}");
                throw new InvalidOperationException($"Workbook summary validation failed: {responseContent}");
            }

            _logger.LogInformation("Successfully completed workbook summary generation");
        }
        catch (Exception ex)
        {
            _logger.LogError($"\n=== Error in workbook summary generation ===");
            _logger.LogError($"Error Type: {ex.GetType().Name}");
            _logger.LogError($"Error Message: {ex.Message}");
            _logger.LogError($"Stack Trace: {ex.StackTrace}");
            throw;
        }
    }

    public async Task CreateExcelWorksheets(Excel.Workbook workbook, EntityReference workbookRef)
    {
        _logger.LogInformation("=== Starting Excel Workbook Worksheets Processing ===");
        _logger.LogInformation($"Using Workbook Primary Key: {workbookRef.PalantirPrimaryKey}");
        _logger.LogInformation($"Using Local GUID: {workbookRef.LocalGuid}");
        _logger.LogInformation($"Base URL: {_baseUrl}");

        try
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    _logger.LogInformation($"\n=== Processing worksheet: {worksheet.Name} ===");
                    var worksheetRef = await CreateExcelWorksheet(worksheet, workbookRef.PalantirPrimaryKey);
                    
                    _logger.LogInformation($"Processing cells for worksheet: {worksheet.Name}");
                    await ProcessWorksheetCells(worksheet, worksheetRef.PalantirPrimaryKey);
                    
                    _logger.LogInformation($"Processing charts for worksheet: {worksheet.Name}");
                    await ProcessWorksheetCharts(worksheet, worksheetRef.PalantirPrimaryKey);
                    
                    _logger.LogInformation($"Successfully processed worksheet: {worksheet.Name}");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error processing worksheet {worksheet.Name}: {ex.Message}");
                    _logger.LogError($"Stack trace: {ex.StackTrace}");
                    continue;
                }
            }

            // Generate workbook summary after all worksheets are processed
            _logger.LogInformation("\n=== Initiating Final Workbook Summary ===");
            _logger.LogInformation($"Using Workbook Primary Key for Summary: {workbookRef.PalantirPrimaryKey}");
            await GenerateWorkbookSummary(workbookRef.PalantirPrimaryKey);
            _logger.LogInformation("=== Workbook Summary Generation Complete ===");
        }
        catch (Exception ex)
        {
            _logger.LogError("Critical error processing workbook worksheets");
            _logger.LogError($"Error message: {ex.Message}");
            _logger.LogError($"Stack trace: {ex.StackTrace}");
            throw;
        }
        finally
        {
            _logger.LogInformation("=== Completed Excel Workbook Worksheets Processing ===\n");
        }
    }

    public void Dispose()
    {
        // Cleanup code if needed
        GC.SuppressFinalize(this);
    }
} 
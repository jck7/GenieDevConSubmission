using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using ExcelGenie.Services;
using ExcelGenie.ViewModels;
using ExcelGenie.Controls;
using ExcelGenie.Models;
using ExcelGenie.Converters;
using MediaColor = System.Windows.Media.Color;
using MediaColors = System.Windows.Media.Colors;
using WinFontFamily = System.Windows.Media.FontFamily;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using Microsoft.Extensions.Configuration;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Reflection;
using Border = System.Windows.Controls.Border;
using System.Windows.Shell;
using NLog;
using Run = System.Windows.Documents.Run;
using Paragraph = System.Windows.Documents.Paragraph;
using TextElement = System.Windows.Documents.TextElement;
using WinPath = System.Windows.Shapes.Path;
using IOPath = System.IO.Path;
using MediaGradientStop = System.Windows.Media.GradientStop;
using Color = System.Windows.Media.Color;
using System.Net.Http;
using Microsoft.Win32;
using Microsoft.Extensions.Logging;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace ExcelGenie
{
    public partial class MainWindow : Window, INotifyPropertyChanged, IDisposable
    {
        private string? selectedFilePath;
        private ExcelService? excelService;
        private enum OfficeApp { Excel }
        private OfficeApp selectedOfficeApp = OfficeApp.Excel;
        private readonly Dictionary<OfficeApp, MediaColor> appColors = new Dictionary<OfficeApp, MediaColor>
        {
            { OfficeApp.Excel, (MediaColor)ColorConverter.ConvertFromString("#1D6F42") }
        };

        [DllImport("user32.dll")]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, RedrawWindowFlags flags);

        [Flags]
        enum RedrawWindowFlags : uint
        {
            Invalidate = 0x1,
            InternalPaint = 0x2,
            Erase = 0x4,
            Validate = 0x8,
            NoInternalPaint = 0x10,
            NoErase = 0x20,
            NoChildren = 0x40,
            AllChildren = 0x80,
            UpdateNow = 0x100,
            EraseNow = 0x200,
            Frame = 0x400,
            NoFrame = 0x800
        }

        private string? supportingFile = null;
        private readonly string[] allowedExtensions = { ".docx", ".xlsx", ".pdf" };
        private string supportingFileContent = string.Empty;
        private readonly BlackboxLogger _logger;
        private ExcelService? _excelService;
        private readonly FileExtractionService _fileExtractionService;
        private readonly LoadingIndicatorManager _loadingManager;

        private readonly string SettingsFilePath = IOPath.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "GenieForExcel",
            "settings.txt"
        );

        private readonly string GenieForExcelFolderPath = IOPath.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "GenieForExcel"
        );

        private bool _isManualSelected;
        public bool IsManualSelected
        {
            get => _isManualSelected;
            set
            {
                if (_isManualSelected != value)
                {
                    _isManualSelected = value;
                    OnPropertyChanged(nameof(IsManualSelected));
                }
            }
        }

        private bool _isSupportingFileSelectionEnabled;
        public bool IsSupportingFileSelectionEnabled
        {
            get => _isSupportingFileSelectionEnabled;
            set
            {
                if (_isSupportingFileSelectionEnabled != value)
                {
                    _isSupportingFileSelectionEnabled = value;
                    OnPropertyChanged(nameof(IsSupportingFileSelectionEnabled));
                }
            }
        }

        // Modify the PropertyChanged event to be nullable
        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private bool _isFileSelected;
        public bool IsFileSelected
        {
            get => _isFileSelected;
            set
            {
                if (_isFileSelected != value)
                {
                    _isFileSelected = value;
                    OnPropertyChanged(nameof(IsFileSelected));
                }
            }
        }

        // Add this property to store the email
        private string? UserEmail { get; set; }
        private string? UserSubscriptionStatus { get; set; }
        private const string ApiUrl = "";
        private const string BearerToken = "";

        // Add new fields for worksheet and range tracking
        private ObservableCollection<string> WorksheetNames { get; } = new ObservableCollection<string>();
        private string? selectedRange;
        private string? selectedRangeContent;

        // Add these fields at the top of the MainWindow class
        private readonly RecentFilesService recentFilesService;
        private readonly ChatManager chatManager;
        private bool _autoScroll = true;

        private bool _disposed;
        private readonly Dispatcher _uiDispatcher;
        private readonly DispatcherTimer _internetCheckTimer;
        private bool _hasInternetConnection = true;
        private bool _isDarkTheme = true;
        public bool isDarkTheme
        {
            get => _isDarkTheme;
            set
            {
                if (_isDarkTheme != value)
                {
                    _isDarkTheme = value;
                    OnPropertyChanged(nameof(isDarkTheme));
                }
            }
        }

        public bool HasInternetConnection
        {
            get => _hasInternetConnection;
            private set
            {
                if (_hasInternetConnection != value)
                {
                    _hasInternetConnection = value;
                    _uiDispatcher.Invoke(() =>
                    {
                        NoInternetOverlay.Visibility = value ? Visibility.Collapsed : Visibility.Visible;
                    });
                    OnPropertyChanged(nameof(HasInternetConnection));
                }
            }
        }

        // Add these P/Invoke declarations at the top of the class
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        static readonly IntPtr HWND_TOP = new IntPtr(0);
        const uint SWP_SHOWWINDOW = 0x0040;

        // Add at class level
        private readonly Dictionary<string, (SolidColorBrush dark, SolidColorBrush light)> ThemeColors = new()
        {
            { "Background", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#181818")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#FFFFFF"))
            )},
            { "SecondaryBackground", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#212121")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#F5F5F5"))
            )},
            { "Border", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#3A3A3A")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#E0E0E0"))
            )},
            { "Text", (
                dark: new SolidColorBrush(MediaColors.White),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#202020"))
            )},
            { "ButtonBackground", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#2D2D2D")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#F8F8F8"))
            )},
            { "ButtonHover", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#3D3D3D")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#F0F0F0"))
            )},
            { "ButtonPressed", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#323232")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#E8E8E8"))
            )},
            { "PlaceholderText", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#505050")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#A0A0A0"))
            )},
            { "ScrollBarBackground", (
                dark: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#808080")),
                light: new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#C0C0C0"))
            )}
        };

        private bool isProMode = false;

        private bool _hasText;
        public bool HasText
        {
            get => _hasText;
            private set
            {
                if (_hasText != value)
                {
                    _hasText = value;
                    OnPropertyChanged(nameof(HasText));
                }
            }
        }

        private bool isHighlighting = false;

        private readonly PalantirService _palantirService;
        private const string PalantirToken = "";
        private CancellationTokenSource? _generateCancellationTokenSource;

        private bool _isExcelConnected;
        public bool IsExcelConnected
        {
            get => _isExcelConnected;
            set
            {
                if (_isExcelConnected != value)
                {
                    _isExcelConnected = value;
                    OnPropertyChanged(nameof(IsExcelConnected));
                    UpdateConnectionButton();
                }
            }
        }

        private RichTextBox _inputTextBox;
        private ListBox _recentFilesList;
        private Grid _mainGrid;

        private bool _isConnected;
        public bool IsConnected
        {
            get => _isConnected;
            set
            {
                if (_isConnected != value)
                {
                    _isConnected = value;
                    OnPropertyChanged(nameof(IsConnected));
                    UpdateConnectionButton();
                }
            }
        }

        private void UpdateConnectionButton()
        {
            if (ConnectionStatusButton != null)
            {
                if (IsExcelConnected)
                {
                    ConnectionStatusButton.Content = "Connected";
                    ConnectionStatusButton.Background = new SolidColorBrush(System.Windows.Media.Colors.Green);
                }
                else
                {
                    ConnectionStatusButton.Content = "Connect";  // Changed from "Not Connected" to "Connect"
                    ConnectionStatusButton.Background = new SolidColorBrush(System.Windows.Media.Colors.Red);
                }
                ConnectionStatusButton.IsEnabled = true;
            }
        }

        private void OnExcelConnectionLost(object? sender, EventArgs e)
        {
            IsConnected = false;
            IsExcelConnected = false;  // This will update the button
        }

        private void OnExcelConnectionEstablished(object? sender, EventArgs e)
        {
            IsConnected = true;
            IsExcelConnected = true;  // This will update the button
        }

        public class GenerationResult
        {
            public string Code { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
        }

        private bool _isPalantirConnected;
        public bool IsPalantirConnected
        {
            get => _isPalantirConnected;
            set
            {
                if (_isPalantirConnected != value)
                {
                    _isPalantirConnected = value;
                    OnPropertyChanged(nameof(IsPalantirConnected));
                    Dispatcher.Invoke(() =>
                    {
                        ConnectionStatusMessage.Visibility = value ? Visibility.Collapsed : Visibility.Visible;
                    });
                }
            }
        }

        public MainWindow()
        {
            try
            {
                InitializeComponent();
                DataContext = this;
                
                _loadingManager = new LoadingIndicatorManager(this);
                PalantirProcessingOverlay.Visibility = Visibility.Collapsed; // Hide initially

                _logger = new BlackboxLogger();
                _fileExtractionService = new FileExtractionService(_logger);
                _excelService = new ExcelService(_logger, Dispatcher);
                
                // Initialize PalantirService with all required parameters
                var httpClient = new HttpClient();
                const string ontologyId = "";
                const string baseUrl = "";
                var loggerFactory = LoggerFactory.Create(builder => builder.AddConsole());
                var palantirLogger = loggerFactory.CreateLogger<PalantirService>();
                _palantirService = new PalantirService(palantirLogger, httpClient, PalantirToken, "", ontologyId, baseUrl);
                
                _uiDispatcher = Dispatcher.CurrentDispatcher;
                _mainGrid = MainGrid;
                _inputTextBox = InputTextBox;
                _recentFilesList = new ListBox();
                chatManager = new ChatManager();
                recentFilesService = new RecentFilesService(
                    IOPath.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GenieForExcel")
                );
                _internetCheckTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
                _internetCheckTimer.Tick += async (s, e) => await CheckInternetConnection();

                if (_mainGrid == null || _inputTextBox == null)
                {
                    throw new InvalidOperationException("Failed to initialize UI elements");
                }

                InitializeServices();
                InitializeEventHandlers();
                LoadRecentFiles();
                SetupDragAndDrop();
                SetupWindowChrome();
                SetupTheme();
                SetupChatManager();
                SetupPalantirConnection();

                // Start checking internet connection
                _ = CheckInternetConnection();
                _internetCheckTimer.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing application: {ex.Message}", "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        private void InitializeServices()
        {
            string logFilePath = IOPath.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "GenieForExcel",
                "Logs",
                $"blackbox_{DateTime.Now:yyyyMMdd_HHmmss}.log"
            );

            _logger.SetLogFilePath(logFilePath);
            _logger.LogSystemMessage($"=== Starting New Session ===");
            _logger.LogSystemMessage($"Log File: {logFilePath}");
            _logger.LogSystemMessage($"Application Version: 1.0.0");
            _logger.LogSystemMessage($"OS Version: {Environment.OSVersion}");
            _logger.LogSystemMessage($"Machine Name: {Environment.MachineName}");
            _logger.LogSystemMessage($"User Name: {Environment.UserName}");
            _logger.LogSystemMessage("===============================\n");

            // Rest of service initialization...
        }

        private void InitializeEventHandlers()
        {
            // Subscribe to Excel connection events
            _excelService.ConnectionLost += OnExcelConnectionLost;
            _excelService.ConnectionEstablished += OnExcelConnectionEstablished;

            // Initialize connection status
            IsConnected = _excelService.IsConnected;
            IsExcelConnected = _excelService.IsConnected;

            // Subscribe to events
            if (excelService != null)
            {
                excelService.WorksheetRenamed += ExcelService_WorksheetRenamed;
                excelService.WorksheetDeleted += ExcelService_WorksheetDeleted;
            }

            AddEventHandlersToAllControls();
        }

        private void SetupChatManager()
        {
            if (chatManager != null)
            {
                chatManager.MessageAdded += ChatManager_MessageAdded;
                chatManager.LoadingMessageAdded += ChatManager_LoadingMessageAdded;
                chatManager.LoadingMessageRemoved += ChatManager_LoadingMessageRemoved;
                chatManager.SystemResponseAdded += ChatManager_SystemResponseAdded;
                chatManager.ConversationReverted += ChatManager_ConversationReverted;
            }
        }

        private void SetupPalantirConnection()
        {
            if (_palantirService != null)
            {
                _palantirService.ConnectionStatusChanged += OnPalantirConnectionStatusChanged;
            }
        }

        private void OnPalantirConnectionStatusChanged(object? sender, bool isConnected)
        {
            Dispatcher.Invoke(() =>
            {
                IsPalantirConnected = isConnected;
            });
        }

        private void InitializeControls()
        {
            // No need for additional initialization since it's handled by XAML
        }

        private void AddEventHandlersToAllControls()
        {
            foreach (UIElement element in _mainGrid.Children)
            {
                if (element is FrameworkElement fe && fe.Name == "BottomButtons")
                    continue;
                    
                element.AddHandler(UIElement.MouseDownEvent, new MouseButtonEventHandler(Element_MouseDown), true);
                element.AddHandler(UIElement.TouchDownEvent, new EventHandler<TouchEventArgs>(Element_TouchDown), true);
            }
        }

        // This is the first instance of UpdateUIBasedOnSubscription - keep it
        private void UpdateUIBasedOnSubscription()
        {
            // Always enable all controls
                UpdateMainWindowControls(true);
        }

        private void UpdateMainWindowControls(bool enable)
        {
            if (InputTextBox != null) InputTextBox.IsEnabled = enable;
            if (GenerateButton != null) GenerateButton.IsEnabled = enable;
            if (DragDropArea != null) DragDropArea.IsEnabled = enable;
            if (SelectRangeButton != null) SelectRangeButton.IsEnabled = enable;
            if (CustomInstructionsButton != null) CustomInstructionsButton.IsEnabled = enable;
            
            // These controls are always enabled
            if (SettingsButton != null) SettingsButton.IsEnabled = true;
            if (InfoButton != null) InfoButton.IsEnabled = true;
            if (ThemeToggleButton != null) ThemeToggleButton.IsEnabled = true;
        }

        private string GetInputText()
        {
            if (InputTextBox == null) return string.Empty;

            TextRange textRange = new TextRange(
                InputTextBox.Document.ContentStart,
                InputTextBox.Document.ContentEnd
            );
            return textRange.Text.Trim();
        }

        private void SetInputText(string text)
        {
            if (InputTextBox != null)
        {
            InputTextBox.Document.Blocks.Clear();
            InputTextBox.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
        }

        private void SetCaretPosition(int position)
        {
            if (InputTextBox?.Document == null) return;
            var caretPosition = InputTextBox.Document.ContentStart.GetPositionAtOffset(position);
            if (caretPosition != null)
            {
                InputTextBox.CaretPosition = caretPosition;
            }
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
                    _logger?.Dispose();
                    _excelService?.Dispose();
                    _fileExtractionService?.Dispose();
                    _internetCheckTimer?.Stop();
                    _palantirService?.Dispose();
                }
                _disposed = true;
            }
        }

        ~MainWindow()
        {
            Dispose(false);
        }

        private void ChatManager_MessageAdded(object? sender, (string message, bool isUser) e)
        {
            var (message, isUser) = e;
            AddChatMessageToUI(message, isUser);
        }

        private void ChatManager_LoadingMessageAdded(object? sender, LoadingMessageInfo e)
        {
            var loadingPanel = CreateLoadingMessageUI();
            e.Element = loadingPanel;
            ChatPanel.Children.Add(loadingPanel);
            ChatPanel.UpdateLayout();
            ScrollChatToBottom();
        }

        private void ChatManager_LoadingMessageRemoved(object? sender, LoadingMessageInfo e)
        {
            if (e.Element == null) return;
            
            Dispatcher.Invoke(() =>
            {
                ChatPanel.Children.Remove(e.Element);
                ScrollChatToBottom();
            });
        }

        private void ChatManager_SystemResponseAdded(object? sender, (string description, string code) e)
        {
            var (description, _) = e;
            AddSystemResponseToUI(description, "");
        }

        private void ChatManager_ConversationReverted(object? sender, int chatPanelChildrenCount)
        {
            Dispatcher.Invoke(() =>
            {
                while (ChatPanel.Children.Count > chatPanelChildrenCount)
                {
                    ChatPanel.Children.RemoveAt(ChatPanel.Children.Count - 1);
                }
            });
        }

        private async Task<string> CreateApiMessage(string userInput, string customInstructions, string? selectedRange, string? selectedWorksheet)
        {
            try
            {
                _logger.LogSystemMessage("\n=== Starting CreateApiMessage ===");
                _logger.LogSystemMessage($"User Input: {userInput}");
                _logger.LogSystemMessage($"Selected Range: {selectedRange ?? "None"}");
                _logger.LogSystemMessage($"Selected Worksheet: {selectedWorksheet ?? "None"}");

                // Get workbook ID from WorkbookExplorer
                var workbookId = WorkbookExplorer.Workbooks.FirstOrDefault()?.PalantirId;
                _logger.LogSystemMessage($"Workbook ID: {workbookId ?? "None"}");

                _logger.LogSystemMessage("Creating user prompt...");
                await _palantirService.CreateUserPrompt(
                    userInput,
                    customInstructions,
                    selectedRange,
                    selectedWorksheet
                );

                _logger.LogSystemMessage("Getting agent plan...");
                if (workbookId == null)
                {
                    throw new InvalidOperationException("No workbook ID available");
                }

                // Get the agent plan first
                var agentPlan = await _palantirService.GetAgentPlan(workbookId);
                _logger.LogSystemMessage($"Agent plan received: {(string.IsNullOrEmpty(agentPlan) ? "Empty" : agentPlan.Length + " chars")}");

                // Display the agent plan immediately
                chatManager?.AddSystemResponse("Agent Plan\n" + agentPlan, "");

                // Start plan steps retrieval in the background
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await GetPlanSteps(workbookId, _palantirService.LatestPromptKey);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error retrieving plan steps: {ex.Message}");
                        // Optionally notify the user of the error through the UI
                        await Dispatcher.InvokeAsync(() =>
                        {
                            chatManager?.AddSystemResponse("**Error retrieving execution steps**\nPlease try again or contact support if the issue persists.", "error");
                        });
                    }
                });
                
                return agentPlan;
            }
            catch (Exception ex)
            {
                _logger.LogError("Error in CreateApiMessage", ex);
                throw;
            }
        }

        private async Task GetPlanSteps(string workbookId, string promptId)
        {
            const int MaxAttempts = 60; // 5 minutes = 60 * 5 seconds
            const int PollIntervalSeconds = 5;

            try
            {
                _logger.LogSystemMessage($"\n=== Starting GetPlanSteps Function ===");
                _logger.LogSystemMessage($"Input - Workbook ID: {workbookId}");
                _logger.LogSystemMessage($"Input - Prompt ID: {promptId}");

                var url = $"{_palantirService.BaseUrl}/v2/ontologies/{_palantirService.OntologyId}/queries/genieForExcelPlanSteps/execute";
                _logger.LogSystemMessage($"Request URL: {url}");
                
                var requestBody = new
                {
                    parameters = new
                    {
                        excelWorkbook = workbookId,
                        userPrompt = promptId
                    }
                };

                var jsonBody = JsonConvert.SerializeObject(requestBody, Formatting.Indented);
                _logger.LogSystemMessage($"Request Body:\n{jsonBody}");

                for (int attempt = 1; attempt <= MaxAttempts; attempt++)
                {
                    try
                    {
                        using var planStepsRequest = new HttpRequestMessage(HttpMethod.Post, url);
                        planStepsRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _palantirService.BearerToken);
                        planStepsRequest.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                        using var planStepsResponse = await _palantirService.HttpClient.SendAsync(planStepsRequest);
                        var planStepsContent = await planStepsResponse.Content.ReadAsStringAsync();
                        
                        _logger.LogSystemMessage($"Plan Steps Response Status: {planStepsResponse.StatusCode}");
                        _logger.LogSystemMessage($"Plan Steps Response Content: {planStepsContent}");

                        if (planStepsResponse.IsSuccessStatusCode)
                        {
                            _logger.LogSystemMessage("=== Plan Steps Retrieved Successfully ===");
                            await Dispatcher.InvokeAsync(() => DisplayPlanSteps(planStepsContent));
                            return;
                        }
                        else if (planStepsContent.Contains("TIMEOUT") || planStepsContent.Contains("QueryTimeExceededLimit"))
                        {
                            _logger.LogSystemMessage($"Attempt {attempt}/{MaxAttempts}: Request timed out, waiting {PollIntervalSeconds} seconds before retry...");
                            await Task.Delay(PollIntervalSeconds * 1000);
                            continue;
                        }
                        else
                        {
                            _logger.LogError($"Failed to get plan steps. Status: {planStepsResponse.StatusCode}");
                            _logger.LogError($"Error Content: {planStepsContent}");
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error on attempt {attempt}: {ex.Message}");
                        if (attempt == MaxAttempts)
                        {
                            await Dispatcher.InvokeAsync(() =>
                            {
                                chatManager?.AddSystemResponse(
                                    "I apologize, but I was unable to process your request after several attempts. " +
                                    "Please try again in a few minutes.", 
                                    "error"
                                );
                            });
                            throw;
                        }
                        await Task.Delay(PollIntervalSeconds * 1000);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"\n=== Error in GetPlanSteps ===");
                _logger.LogError($"Error Type: {ex.GetType().Name}");
                _logger.LogError($"Error Message: {ex.Message}");
                _logger.LogError($"Stack Trace: {ex.StackTrace}");
                throw;
            }
        }

        private async void DisplayPlanSteps(string planStepsContent)
        {
            try
            {
                var planStepsResponse = JsonConvert.DeserializeObject<dynamic>(planStepsContent);
                if (planStepsResponse?.value != null)
                {
                    var steps = planStepsResponse.value;
                    int totalSteps = steps.Count;

                    // First, check if any steps contain actual VBA code
                    bool hasExecutableSteps = false;
                    foreach (var step in steps)
                    {
                        string vbaCode = step.validVBACode?.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(vbaCode) && vbaCode.Contains("Sub "))
                        {
                            hasExecutableSteps = true;
                            break;
                        }
                    }

                    // Display steps and execute only if there are executable steps
                    foreach (var step in steps)
                    {
                        string stepTitle = step.stepTitle.ToString();
                        string stepDescription = step.stepPastTenseDescription.ToString();
                        int stepIndex = (int)step.stepIndex;
                        string vbaCode = step.validVBACode?.ToString() ?? "";

                        // Add each step as a system message
                        chatManager?.AddSystemResponse($"Step {stepIndex}) {stepTitle}\n{stepDescription}", "");
                    }

                    // Only attempt execution if there are actual VBA steps
                    if (hasExecutableSteps)
                    {
                        try
                        {
                            _logger.LogSystemMessage("Starting execution of VBA steps...");
                            await _excelService?.ExecutePlanStepsFromJson(planStepsResponse);
                            _logger.LogSystemMessage("VBA steps execution completed");
                        }
                        catch (Exception execEx)
                        {
                            _logger.LogError($"Error executing VBA steps: {execEx.Message}");
                            chatManager?.AddSystemResponse("**Error**", $"Failed to execute steps: {execEx.Message}");
                        }
                    }
                    else
                    {
                        _logger.LogSystemMessage("No executable VBA steps found - skipping execution phase");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error parsing plan steps: {ex.Message}");
            }
        }

        private async Task<(string vbaCode, string description)> GenerateVBACode(string apiMessage, CancellationToken cancellationToken)
        {
            // VBA functionality is disabled
            return ("", "");
        }

        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            if (!IsExcelConnected)
            {
                MessageBox.Show("Please connect to Excel first.", "Not Connected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(GetInputText()))
            {
                return;
            }

            try
            {
                _logger.LogSystemMessage("=== Starting Generate Button Click ===");
                _generateCancellationTokenSource = new CancellationTokenSource();
                
                if (GenerateButton != null)
                {
                    GenerateButton.IsEnabled = false;
                }

                string userInput = GetInputText();
                SetInputText("");

                _logger.LogSystemMessage($"User input received: {userInput}");
                chatManager?.AddUserMessage(userInput);

                _logger.LogSystemMessage("Creating API message...");
                string apiMessage = await CreateApiMessage(userInput, "", selectedRange, WorksheetNames.Count > 0 ? WorksheetNames[0] : null);
                _logger.LogSystemMessage("API message created successfully");

                // Remove the duplicate display of the agent plan
                // chatManager?.AddSystemResponse(apiMessage, "");
            }
            catch (Exception ex)
            {
                _logger.LogError("Error processing request", ex);
                MessageBox.Show($"Error processing request: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (GenerateButton != null)
                {
                    GenerateButton.IsEnabled = true;
                }
                _generateCancellationTokenSource = null;
                _logger.LogSystemMessage("=== Generate Button Click Completed ===");
            }
        }

        private async void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            if (excelService == null || !excelService.IsWorkbookOpen)
            {
                chatManager.AddSystemMessage("Error: No active workbook to revert.");
                return;
            }

            var lastAction = chatManager.PopUndoAction();
            if (lastAction == null) return;

            try
            {
                Debug.WriteLine("Starting revert operation...");

                // Check if backup file still exists
                if (!File.Exists(lastAction.BackupFilePath))
                {
                    Debug.WriteLine($"Backup file not found: {lastAction.BackupFilePath}");
                    // Push the action back onto the stack since we couldn't complete the revert
                    chatManager.PushUndoAction(lastAction);
                    chatManager.AddSystemMessage("Error: Backup file not found. Unable to revert changes.");
                    return;
                }

                Debug.WriteLine($"Opening backup file: {lastAction.BackupFilePath}");
                
                // Remove chat messages back to the state before the action
                // We need to remove both the system response and the user's message
                // So we remove one more message than the stored count indicates
                Dispatcher.Invoke(() =>
                {
                    while (ChatPanel.Children.Count > Math.Max(0, lastAction.ChatPanelChildrenCount - 1))
                    {
                        ChatPanel.Children.RemoveAt(ChatPanel.Children.Count - 1);
                    }
                });

                // Reopen from the backup file
                bool success = await excelService.LoadWorkbookAsync(lastAction.BackupFilePath, GenieForExcelFolderPath);
                if (!success)
                {
                    // Push the action back onto the stack since we couldn't complete the revert
                    chatManager.PushUndoAction(lastAction);
                    chatManager.AddSystemMessage("Error: Could not load backup file. Unable to revert changes.");
                    return;
                }

                Debug.WriteLine("Revert operation completed successfully.");
                chatManager.AddSystemMessage("Changes reverted successfully.");
            }
            catch (Exception ex)
            {
                // Log the detailed error for debugging
                Debug.WriteLine($"Error during revert operation: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                
                // Push the action back onto the stack since we couldn't complete the revert
                chatManager.PushUndoAction(lastAction);
                chatManager.AddSystemMessage($"Error: Unable to revert changes. {ex.Message}");
            }
        }

        private void AddChatMessageToUI(string message, bool isUser)
        {
            // Ensure ChatPanel is visible and NoFileView is collapsed
            ChatPanel.Visibility = Visibility.Visible;
            NoFileView.Visibility = Visibility.Collapsed;

            var messagePanel = new StackPanel
            {
                Margin = new Thickness(10, 2, 10, 2),
                HorizontalAlignment = isUser ? System.Windows.HorizontalAlignment.Right : System.Windows.HorizontalAlignment.Left
            };

            var messageBorder = new Border
            {
                Background = new SolidColorBrush(isUser ? 
                    (MediaColor)ColorConverter.ConvertFromString("#1D6F42") : 
                    (MediaColor)ColorConverter.ConvertFromString("#212121")),
                BorderBrush = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#3A3A3A")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(8),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left  // Always left align the border
            };

            var messageRichTextBox = new RichTextBox
            {
                IsReadOnly = true,
                Background = null,
                BorderThickness = new Thickness(0),
                Foreground = new SolidColorBrush(MediaColors.White),
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 13,
                Margin = new Thickness(0),
                Style = (Style)FindResource("TransparentRichTextBox"),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                Width = double.NaN,  // Auto width
                MaxWidth = 600  // Maximum width constraint
            };

            // Create a simple paragraph with the message text
            var paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run(message));
            messageRichTextBox.Document = new FlowDocument(paragraph);

            messageBorder.Child = messageRichTextBox;
            messagePanel.Children.Add(messageBorder);

            Dispatcher.Invoke(() =>
            {
                ChatPanel.Children.Add(messagePanel);
                ScrollChatToBottom();
            });
        }

        private UIElement CreateLoadingMessageUI()
        {
            var messagePanel = new StackPanel
            {
                Margin = new Thickness(10, 2, 10, 2),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            };

            var messageBorder = new Border
            {
                Background = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#212121")),
                BorderBrush = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#3A3A3A")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(8, 0, 8, 0),
                Height = 23,  // Match the height of the revert button
                MinWidth = 100  // Ensure enough space for text and animation
            };

            var loadingContent = new Grid  // Changed to Grid for better alignment
            {
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                Margin = new Thickness(0)
            };

            var loadingText = new TextBlock
            {
                Text = "Generating",  // Start with Generating
                Foreground = new SolidColorBrush(MediaColors.White),
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                FontSize = 13,
                FontFamily = new WinFontFamily("Segoe UI"),
                Margin = new Thickness(0, 1, 4, 0)  // Small top margin for vertical centering
            };

            var dotsPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                Margin = new Thickness(loadingText.Text.Length * 7, 1, 0, 0)  // Position based on text length
            };

            for (int i = 0; i < 3; i++)
            {
                var dot = new TextBlock
                {
                    Text = ".",
                    Foreground = new SolidColorBrush(MediaColors.White),
                    FontSize = 13,
                    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                    Opacity = 0,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                    Margin = new Thickness(1, -2, 1, 0)  // Negative top margin to align dots
                };

                var animation = new DoubleAnimation
                {
                    From = 0,
                    To = 1,
                    Duration = TimeSpan.FromSeconds(0.5),
                    AutoReverse = true,
                    RepeatBehavior = RepeatBehavior.Forever,
                    BeginTime = TimeSpan.FromSeconds(i * 0.2)
                };

                dot.BeginAnimation(UIElement.OpacityProperty, animation);
                dotsPanel.Children.Add(dot);
            }

            loadingContent.Children.Add(loadingText);
            loadingContent.Children.Add(dotsPanel);
            messageBorder.Child = loadingContent;
            messagePanel.Children.Add(messageBorder);

            return messagePanel;
        }

        private void AddSystemResponseToUI(string description, string code)
        {
            // Ensure ChatPanel is visible and NoFileView is collapsed
            ChatPanel.Visibility = Visibility.Visible;
            NoFileView.Visibility = Visibility.Collapsed;

            var messagePanel = new Grid();
            var messageBorder = new Border
            {
                Background = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#212121")),
                BorderBrush = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#3A3A3A")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(8)
            };

            var messageStackPanel = new StackPanel();

            // Description
            if (!string.IsNullOrWhiteSpace(description))
            {
                var expander = new Expander
                {
                    IsExpanded = false,
                    ExpandDirection = ExpandDirection.Down,
                    Margin = new Thickness(0, 4, 0, 4),
                    Foreground = new SolidColorBrush(MediaColors.White)
                };

                // Header content
                var headerTextBlock = new TextBlock
                {
                    Text = new FirstLineConverter().Convert(description, typeof(string), null, null) as string,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(MediaColors.White)
                };
                expander.Header = headerTextBlock;

                // Full content (excluding the first line)
                var contentRichTextBox = new RichTextBox
                {
                    IsReadOnly = true,
                    Background = null,
                    BorderThickness = new Thickness(0),
                    Foreground = new SolidColorBrush(MediaColors.White),
                    FontFamily = new WinFontFamily("Segoe UI"),
                    FontSize = 13,
                    Margin = new Thickness(10, 0, 0, 0),
                    Style = (Style)FindResource("TransparentRichTextBox"),
                    Width = double.NaN,  // Auto width
                    MaxWidth = 600  // Maximum width constraint
                };

                // Get the remaining lines (skip the first line)
                var lines = description.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                var remainingText = string.Join(Environment.NewLine, lines.Skip(1));

                // Create a simple paragraph with the remaining text
                var paragraph = new Paragraph();
                paragraph.Inlines.Add(new Run(remainingText));
                contentRichTextBox.Document = new FlowDocument(paragraph);

                expander.Content = contentRichTextBox;
                messageStackPanel.Children.Add(expander);
            }

            messageBorder.Child = messageStackPanel;
            messagePanel.Children.Add(messageBorder);

            Dispatcher.Invoke(() =>
            {
                ChatPanel.Children.Add(messagePanel);
                ScrollChatToBottom();
            });
        }

        private void ResetUIState()
        {
            SelectFileButton.Content = "Select Excel File";
            InputTextBox.IsEnabled = false;
            GenerateButton.IsEnabled = false;
            selectedFilePath = null;
            excelService = null;
            UpdateButtonState();
            IsSupportingFileSelectionEnabled = false;
            IsFileSelected = false;
            ChatPanel.Children.Clear();
            chatManager.ClearConversation();
            WorkbookExplorer.Workbooks.Clear();
        }

        private void Element_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // Remove subscription check - allow all clicks
            if (sender is UIElement element)
            {
                element.Focus();
            }
        }

        private void Element_TouchDown(object? sender, TouchEventArgs e)
        {
            if (sender is UIElement element)
            {
                element.Focus();
            }
        }

        private async void LoadSettingsAndCheckSubscription()
        {
            LoadSettings();
            // Always enable UI without subscription check
                UpdateUIBasedOnSubscription();
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    UserEmail = File.ReadAllText(SettingsFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveSettings()
        {
            try
            {
                File.WriteAllText(SettingsFilePath, UserEmail);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task LoadWorkbook()
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                return;
            }

            try
            {
                _logger.LogSystemMessage($"Loading workbook: {selectedFilePath}");
                UpdateCurrentOperation("Opening workbook...");

                if (_excelService == null)
                {
                    _excelService = new ExcelService(_logger, Dispatcher.CurrentDispatcher);
                    AttachExcelEventHandlers();
                }

                var success = await _excelService.LoadWorkbookAsync(selectedFilePath, GenieForExcelFolderPath);
                if (!success)
                {
                    throw new Exception("Failed to open workbook");
                }

                var workbook = _excelService.ActiveWorkbook;
                if (workbook == null)
                {
                    throw new Exception("Failed to get active workbook");
                }

                // Create workbook in Palantir
                var workbookInfo = new Dictionary<string, string>
                {
                    ["workbook_path"] = selectedFilePath,
                    ["author"] = Environment.UserName
                };

                var workbookRef = await _palantirService.CreateExcelWorkbook(workbookInfo);
                
                // Process the workbook with the new UI
                await ProcessWorkbook(workbook, workbookRef);

                // Update UI state
                IsFileSelected = true;
                UpdateWorksheetList();
                UpdateButtonState();
                
                // Populate WorkbookExplorer
                var workbookViewModel = new WorkbookViewModel
                {
                    Name = IOPath.GetFileName(selectedFilePath),
                    PalantirId = workbookRef.PalantirPrimaryKey
                };

                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    var worksheetViewModel = new WorksheetViewModel
                    {
                        Name = worksheet.Name,
                        PalantirId = _palantirService.GetWorksheetId(worksheet.Name)
                    };

                    // Get charts in the worksheet
                    var charts = _excelService.GetChartsInWorksheet(worksheet.Name);
                    foreach (var chartName in charts)
                    {
                        worksheetViewModel.Charts.Add(new ChartViewModel
                        {
                            Name = chartName,
                            PalantirId = _palantirService.GetChartId(chartName)
                        });
                    }

                    workbookViewModel.Worksheets.Add(worksheetViewModel);
                }

                WorkbookExplorer.Workbooks.Clear();
                WorkbookExplorer.Workbooks.Add(workbookViewModel);

                // Add a delay to ensure Excel is fully loaded
                await Task.Delay(1000);

                // Try positioning windows multiple times with delays
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        await Task.Delay(500 * (i + 1));
                        await PositionWindowsAfterInit();
                        
                        // Verify the positioning worked
                        if (_excelService?.ActiveWorkbook != null)
                        {
                            IntPtr excelHwnd = new IntPtr(_excelService.Hwnd);
                            RECT rect;
                            if (GetWindowRect(excelHwnd, out rect))
                            {
                                int width = rect.Right - rect.Left;
                                var screen = System.Windows.Forms.Screen.PrimaryScreen;
                                if (screen != null && width <= (int)(screen.WorkingArea.Width * 0.62)) // Allow for small margin of error
                                {
                                    break; // Positioning successful
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Attempt {i + 1} to position windows failed", ex);
                        if (i == 2) throw; // Throw on last attempt
                    }
                }
                
                _logger.LogSystemMessage("Workbook loaded successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error loading workbook: {ex.Message}");
                MessageBox.Show($"Error loading workbook: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task CleanupExcelProcesses()
        {
            try
            {
                if (_excelService != null)
                {
                    _excelService.ConnectionLost -= OnExcelConnectionLost;
                    _excelService.ConnectionEstablished -= OnExcelConnectionEstablished;
                    await Task.Run(() => _excelService.Dispose());
                    _excelService = null;
                }
                IsConnected = false;
                IsExcelConnected = false;
                IsFileSelected = false;
                WorkbookExplorer.Workbooks.Clear();
                _logger.LogSystemMessage("Excel processes cleaned up successfully");
            }
            catch (Exception ex)
            {
                IsConnected = false;
                IsExcelConnected = false;
                IsFileSelected = false;
                WorkbookExplorer.Workbooks.Clear();
                _logger.LogError($"Error cleaning up Excel processes: {ex.Message}");
            }
        }

        private async void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xls|All Files|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    _logger.LogSystemMessage($"File selected: {dialog.FileName}");

                    // Create a new log file for this session
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string logFileName = $"blackbox_log_{timestamp}.txt";
                    string logFolderPath = IOPath.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        "GenieForExcel",
                        "Logs"
                    );
                    string logFilePath = IOPath.Combine(logFolderPath, logFileName);
                    _logger.LogSystemMessage($"Creating new log file at: {logFilePath}");

                    // Create the log file
                    try
                    {
                        if (!Directory.Exists(logFolderPath))
                        {
                            Directory.CreateDirectory(logFolderPath);
                        }
                        File.WriteAllText(logFilePath, ""); // Create empty file
                        _logger.SetLogFilePath(logFilePath);
                        _logger.LogSystemMessage($"New session started with file: {dialog.FileName}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Error creating log file", ex);
                        MessageBox.Show($"Warning: Could not create log file: {ex.Message}", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                    selectedFilePath = dialog.FileName;
                    AddToRecentFiles(dialog.FileName);

                    // Load the workbook
                    LoadWorkbook();
                }
                catch (Exception ex)
                {
                    _logger.LogError("Error in SelectFile_Click", ex);
                    MessageBox.Show($"Error selecting file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void WorksheetButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Content is string worksheetName)
            {
                excelService?.ActivateWorksheet(worksheetName);
            }
        }

        private void AttachExcelEventHandlers()
        {
            if (_excelService == null) return;

            // First detach any existing handlers to prevent duplicates
            DetachExcelEventHandlers();

            // Attach new handlers
            _excelService.WorkbookLoaded += ExcelService_WorkbookLoaded;
            _excelService.WorksheetActivated += ExcelService_WorksheetActivated;
            _excelService.ErrorOccurred += ExcelService_ErrorOccurred;
            _excelService.WorkbookClosed += ExcelService_WorkbookClosed;
            _excelService.WorksheetRenamed += ExcelService_WorksheetRenamed;
            _excelService.WorksheetDeleted += ExcelService_WorksheetDeleted;

            // Add timer for periodic updates
            var updateTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            updateTimer.Tick += (s, e) => UpdateWorksheetList();
            updateTimer.Start();
        }

        private void DetachExcelEventHandlers()
        {
            if (_excelService == null) return;

            _excelService.WorkbookLoaded -= ExcelService_WorkbookLoaded;
            _excelService.WorksheetActivated -= ExcelService_WorksheetActivated;
            _excelService.ErrorOccurred -= ExcelService_ErrorOccurred;
            _excelService.WorkbookClosed -= ExcelService_WorkbookClosed;
            _excelService.WorksheetRenamed -= ExcelService_WorksheetRenamed;
            _excelService.WorksheetDeleted -= ExcelService_WorksheetDeleted;
        }

        private async Task PositionWindowsWithRetry()
        {
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    await Task.Delay(500 * (i + 1));
                    var result = await Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        if (_excelService?.ActiveWorkbook != null)
                        {
                            await PositionWindowsAfterInit();
                            return true;
                        }
                        return false;
                    }).Task;

                    if (await result) break;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Attempt {i + 1} to position windows failed: {ex.Message}");
                    _logger.LogError($"Attempt {i + 1} to position windows failed", ex);
                }
            }
        }

        private async Task<string> CreateWorkbookBackup()
        {
            _logger.LogSystemMessage("=== Starting Workbook Backup Process ===");
            try
            {
                // Verify Excel service state
                if (_excelService == null)
                {
                    _logger.LogError("Backup: Excel service is not available");
                    throw new InvalidOperationException("Excel service is not available. Please try reopening the file.");
                }

                if (_excelService.ActiveWorkbook == null)
                {
                    _logger.LogError("Backup: No active workbook");
                    throw new InvalidOperationException("No active workbook is loaded. Please try reopening the file.");
                }

                // Create a unique backup file name
                string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                string backupFileName = $"backup_{timestamp}.xlsm";
                string backupPath = IOPath.Combine(GenieForExcelFolderPath, backupFileName);

                _logger.LogSystemMessage($"Backup: Creating backup at path - {backupPath}");

                // Ensure the directory exists
                Directory.CreateDirectory(GenieForExcelFolderPath);

                // Verify Excel state before backup
                _logger.LogSystemMessage("Backup: Verifying Excel state");
                bool isConnected = await VerifyExcelConnection();
                if (!isConnected)
                {
                    _logger.LogError("Backup: Excel connection verification failed");
                    throw new InvalidOperationException("Excel connection is not stable. Please try reopening the file.");
                }

                try
                {
                    // Configure Excel for backup
                    _logger.LogSystemMessage("Backup: Configuring Excel settings");
                    _excelService.DisplayAlerts = false;
                    _excelService.ScreenUpdating = false;
                    _excelService.EnableEvents = false;

                    // Save current workbook state
                    _logger.LogSystemMessage("Backup: Saving current workbook state");
                    _excelService.ActiveWorkbook.Save();

                    // Create backup using SaveCopyAs
                    _logger.LogSystemMessage("Backup: Creating backup copy");
                    _excelService.ActiveWorkbook.SaveCopyAs(backupPath);

                    // Verify backup file was created
                    if (!File.Exists(backupPath))
                    {
                        throw new InvalidOperationException("Backup file was not created successfully.");
                    }

                    _logger.LogSystemMessage($"Backup: Successfully created at {backupPath}");
                    return backupPath;
                }
                catch (Exception ex)
                {
                    _logger.LogError("Backup: Error during backup creation", ex);
                    throw new InvalidOperationException($"Failed to create backup: {ex.Message}. Please ensure Excel is running and the workbook is accessible.");
                }
                finally
                {
                    // Restore Excel settings
                    try
                    {
                        _logger.LogSystemMessage("Backup: Restoring Excel settings");
                        _excelService.DisplayAlerts = true;
                        _excelService.ScreenUpdating = true;
                        _excelService.EnableEvents = true;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Backup: Error restoring Excel settings", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Backup: Critical error", ex);
                throw;
            }
            finally
            {
                _logger.LogSystemMessage("=== Workbook Backup Process Completed ===");
            }
        }

        private void ExcelService_WorkbookLoaded(object? sender, string e)
        {
            selectedFilePath = e;
            UpdateSelectFileButtonUI(TruncateFileName(IOPath.GetFileName(e)));
            InputTextBox.IsEnabled = true;
            GenerateButton.IsEnabled = true;
            IsFileSelected = true;
            UpdateWorksheetList();
            
            // Update the WorkbookExplorer with the loaded workbook
            var workbookVm = new ViewModels.WorkbookViewModel
            {
                Name = e,
                PalantirId = _palantirService.GetWorkbookId(e)
            };

            // Add worksheets
            foreach (var worksheetName in WorksheetNames)
            {
                var worksheetVm = new ViewModels.WorksheetViewModel
                {
                    Name = worksheetName,
                    PalantirId = _palantirService.GetWorksheetId(worksheetName)
                };

                // Add charts if any
                if (_excelService != null)
                {
                    var charts = _excelService.GetChartsInWorksheet(worksheetName);
                    foreach (var chart in charts)
                    {
                        worksheetVm.Charts.Add(new ViewModels.ChartViewModel
                        {
                            Name = chart,
                            PalantirId = _palantirService.GetChartId(chart)
                        });
                    }
                }

                workbookVm.Worksheets.Add(worksheetVm);
            }

            // Update the explorer
            WorkbookExplorer.Workbooks.Clear();
            WorkbookExplorer.Workbooks.Add(workbookVm);
        }

        private void ExcelService_WorksheetActivated(object? sender, string e)
        {
            UpdateWorksheetList();
        }

        private void ExcelService_ErrorOccurred(object? sender, string e)
        {
            _logger.LogError(e);
            AddSystemDebugMessage(e);
        }

        private void ExcelService_WorkbookClosed(object? sender, EventArgs e)
        {
            selectedFilePath = null;
            IsFileSelected = false;
            UpdateSelectFileButtonUI("Select File");
            InputTextBox.IsEnabled = false;
            GenerateButton.IsEnabled = false;
        }

        private void AddSystemDebugMessage(string message)
        {
            _logger.LogSystemMessage($"[Debug] {message}");

            var messagePanel = new StackPanel
            {
                Margin = new Thickness(10, 2, 10, 2),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            };

            var messageBorder = new Border
            {
                Background = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#1E1E1E")),
                BorderBrush = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#3A3A3A")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(10)
            };

            var messageText = new TextBlock
            {
                Text = "I encountered an issue while processing your request. I'll try to handle it differently. Please continue with your task.",
                TextWrapping = TextWrapping.Wrap,
                FontSize = 13,
                Foreground = new SolidColorBrush(MediaColors.White),
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI")
            };

            messageBorder.Child = messageText;
            messagePanel.Children.Add(messageBorder);

            // Log the actual error message to Debug output
            Debug.WriteLine($"System Debug Message: {message}");

            Dispatcher.Invoke(() =>
            {
                ChatPanel.Children.Add(messagePanel);
                ScrollChatToBottom();
            });
        }

        protected override void OnClosed(EventArgs e)
        {
            _internetCheckTimer.Stop();
            base.OnClosed(e);

            if (excelService != null)
            {
                excelService.Dispose();
                excelService = null;
            }
        }

        private void UpdateButtonState()
        {
            if (GenerateButton != null)
            {
                GenerateButton.IsEnabled = !string.IsNullOrEmpty(selectedFilePath) && !string.IsNullOrWhiteSpace(GetInputText());
            }
        }

        private void UpdateUIForSelectedOfficeApp()
        {
            // Update UI elements based on the selected Office app
            switch (selectedOfficeApp)
            {
                case OfficeApp.Excel:
                    SelectFileButton.Content = "Select File";
                    break;
            }

            // Update the global UI color
            UpdateGlobalColor();
        }

        private void UpdateGlobalColor()
        {
            MediaColor newColor = appColors[selectedOfficeApp];
            SolidColorBrush newBrush = new SolidColorBrush(newColor);

            // Update the PrimaryColor resource
            this.Resources["PrimaryColor"] = newBrush;

            // Update the window title
            this.Title = $"Genie for {selectedOfficeApp}";
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new SettingsWindow(UserEmail ?? string.Empty, "Active");
            settingsWindow.Owner = this;
            if (settingsWindow.ShowDialog() == true)
            {
                UserEmail = settingsWindow.Email;
                SaveSettings();
                // No need to check subscription status
                UpdateUIBasedOnSubscription();
            }
        }

        private void CheckSubscriptionAfterSave()
        {
            // Always enable UI without subscription check
                UpdateUIBasedOnSubscription();
        }

        private void PositionWindows()
        {
            _ = PositionWindowsAfterInit();
        }

        // Add this at the top of your file
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private async void DragDropArea_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string filePath = files[0];
                    if (allowedExtensions.Contains(IOPath.GetExtension(filePath).ToLower()))
                    {
                        AddSupportingFile(filePath);
                    }
                }
            }
            DragDropArea.Background = (SolidColorBrush)FindResource("DragDropAreaBackground");
        }

        private void DragDropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string extension = IOPath.GetExtension(files[0]).ToLower();
                    if (extension == ".xlsx" || extension == ".xls" || extension == ".xlsm")
                    {
                        e.Effects = DragDropEffects.Copy;
                        return;
                    }
                }
            }
            e.Effects = DragDropEffects.None;
        }

        private void DragDropArea_DragLeave(object sender, DragEventArgs e)
        {
            // No need to change the background color as it's always white now
        }

        private async void DragDropArea_MouseLeftButtonDown(object sender, RoutedEventArgs e)
        {
            if (supportingFile != null)
            {
                ClearSupportingFile();
            }
            else
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Supported Files|*.docx;*.xlsx;*.pdf"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    await AddSupportingFile(openFileDialog.FileName);
                }
            }
        }

        private void UpdateSupportingFilesList(string? customText = null)
        {
            var button = new Button
            {
                Content = customText ?? "No supporting file",
                Style = (Style)FindResource("RoundedButton"),
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(16, 6, 16, 6),
                Background = new SolidColorBrush(MediaColors.Transparent),
                Foreground = new SolidColorBrush(MediaColors.White),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3A3A3A")),
                BorderThickness = new Thickness(1)
            };

            button.Click += (s, e) => ClearSupportingFile();

            if (DragDropArea is Button dragDropButton)
            {
                dragDropButton.Content = button;
            }
        }

        // Add this helper method to handle file addition
        private async Task AddSupportingFile(string filePath)
        {
            try
            {
                supportingFile = filePath;
                supportingFileContent = await _fileExtractionService.ExtractTextFromFile(filePath);
                UpdateSupportingFilesList(IOPath.GetFileName(filePath));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearSupportingFile();
            }
        }

        private void ClearSupportingFile()
        {
            supportingFile = null;
            supportingFileContent = string.Empty;
            UpdateSupportingFilesList();
        }

        private string GetOriginalFileName(string filePath)
        {
            string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            return fileName != null ? System.Text.RegularExpressions.Regex.Replace(fileName, @"_\d{2}_\d{2}_\d{4}_\d{2}$", "") : string.Empty;
        }

        private void InfoButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://www.genieforapps.com/infoforexcel";
            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open the URL: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void WorksheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                // Add your logic here
            }
        }

        private void UseCaseComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ComboBoxItem selectedItem && selectedItem.Content != null)
            {
                string selectedUseCase = selectedItem.Content.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(selectedUseCase) && selectedUseCase != "Use Case")
                {
                    // You can use the selectedUseCase variable to adjust your application's behavior
                }
            }
        }

        // Add this method to your MainWindow class
        private string TruncateFileName(string fileName, int maxLength = 40)
        {
            if (string.IsNullOrEmpty(fileName) || fileName.Length <= maxLength)
                return fileName;

            return fileName.Substring(0, maxLength - 3) + "...";
        }

        private string GetUseCaseContext(string useCase)
        {
            switch (useCase)
            {
                case "Accounting":
                    return "Follow best practices for accounting in Excel VBA, including proper data validation, audit trails, and adherence to accounting principles. Ensure calculations are accurate and clearly documented.";
                case "Finance":
                    return "Adhere to financial modeling best practices, including clear labeling of inputs and outputs, separation of assumptions and calculations, and proper error handling for financial functions.";
                case "Management Consulting":
                    return "Focus on creating flexible, scalable solutions that can handle various data inputs. Implement modular code structure and provide clear documentation for ease of modification by consultants.";
                case "Data Analysis":
                    return "Emphasize efficient data processing techniques, implement error handling for large datasets, and include data validation checks. Consider performance optimization for handling large volumes of data.";
                case "Project Management":
                    return "Incorporate features for tracking timelines, resources, and milestones. Ensure the code is adaptable to different project structures and includes reporting functionalities.";
                case "Personal Use":
                    return "Prioritize user-friendly interfaces and clear, commented code. Focus on automating repetitive tasks and creating intuitive workflows for personal productivity.";
                default:
                    return "Follow general Excel VBA best practices, including proper error handling, clear code structure, and efficient algorithm implementation.";
            }
        }

        private void OpenSettingsWindow()
        {
            var settingsWindow = new SettingsWindow(UserEmail ?? string.Empty, "Active");
            settingsWindow.Owner = this;
            if (settingsWindow.ShowDialog() == true)
            {
                UserEmail = settingsWindow.Email;
                SaveSettings();
                // No need to check subscription status
                UpdateUIBasedOnSubscription();
            }
        }

        private void UpdateSelectFileButtonUI(string content)
        {
            Dispatcher.Invoke(() =>
            {
                SelectFileButton.Content = content ?? "Select File";
                SelectFileButton.IsEnabled = content is not "Loading..." and not "Opening...";
            });
        }

        private void ScrollChatToBottom()
        {
            if (!_autoScroll) return;

            if (ChatScrollViewer?.Content is FrameworkElement content)
            {
                ChatScrollViewer.ScrollToVerticalOffset(content.ActualHeight);
            }
        }

        private Inline[] ParseMessageWithHighlights(string message)
        {
            var inlines = new List<Inline>();
            int currentIndex = 0;
            while (currentIndex < message.Length)
            {
                int atIndex = message.IndexOf('@', currentIndex);
                if (atIndex == -1)
                {
                    // No more '@' references, add the remaining text as normal
                    if (currentIndex < message.Length)
                    {
                        inlines.Add(new Run(message.Substring(currentIndex)));
                    }
                    break;
                }
                else
                {
                    // Add normal text before '@'
                    if (atIndex > currentIndex)
                    {
                        inlines.Add(new Run(message.Substring(currentIndex, atIndex - currentIndex)));
                    }

                    // Find the end of the reference (space or end of line)
                    int endIndex = message.IndexOf(' ', atIndex);
                    if (endIndex == -1) endIndex = message.Length;

                    // Handle worksheet references with range (e.g., @Sheet1!A1:B2)
                    if (message.IndexOf('!', atIndex) is int bangIndex && bangIndex > atIndex && bangIndex < endIndex)
                    {
                        endIndex = message.IndexOf(' ', bangIndex);
                        if (endIndex == -1) endIndex = message.Length;
                    }

                    string referenceText = message.Substring(atIndex, endIndex - atIndex);

                    // Apply bold and blue to the reference run
                    var referenceRun = new Run(referenceText)
                    {
                        Foreground = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#0078D4")),
                        FontWeight = FontWeights.Bold
                    };
                    inlines.Add(referenceRun);

                    currentIndex = endIndex;
                }
            }

            return inlines.ToArray();
        }

        private void UpdateWorksheetList()
        {
            if (excelService != null && WorksheetListPanel != null)
            {
                var worksheets = excelService.GetWorksheetNames();
                WorksheetListPanel.Children.Clear();
                foreach (var worksheet in worksheets)
                {
                    var button = new Button
                    {
                        Content = worksheet,
                        Style = (Style)FindResource("RoundedButton"),
                        Margin = new Thickness(0, 0, 5, 0)
                    };
                    button.Click += WorksheetButton_Click;
                    WorksheetListPanel.Children.Add(button);
                }
            }
        }

        private int GetTextOffset(TextPointer position)
        {
            TextPointer navigator = InputTextBox.Document.ContentStart;
            int offset = 0;
            
            while (navigator.CompareTo(position) < 0)
            {
                if (navigator.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    offset += navigator.GetTextRunLength(LogicalDirection.Forward);
                }
                navigator = navigator.GetNextContextPosition(LogicalDirection.Forward);
                if (navigator == null) break;
            }
            
            return offset;
        }

        private async void SelectRange_Click(object sender, RoutedEventArgs e)
        {
            _logger.LogSystemMessage("=== Starting Range Selection Process ===");
            try
            {
                if (_excelService == null)
                {
                    _logger.LogError("Range Selection: No Excel service available");
                    MessageBox.Show("Please select an Excel file first.", "No Excel File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _logger.LogSystemMessage("Range Selection: Starting Excel connection verification");
                bool isConnected = await VerifyExcelConnection();
                if (!isConnected)
                {
                    _logger.LogError("Range Selection: Excel connection verification failed");
                    MessageBox.Show("Excel connection is not stable. Please try reopening the file.", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                _logger.LogSystemMessage("Range Selection: Excel connection verified successfully");

                // Get the active worksheet with verification
                Excel.Worksheet? activeSheet = null;
                try
                {
                    _logger.LogSystemMessage("Range Selection: Attempting to get active worksheet");
                    activeSheet = _excelService.ActiveSheet;
                    if (activeSheet == null)
                    {
                        _logger.LogError("Range Selection: No active worksheet found");
                        throw new InvalidOperationException("No active worksheet found");
                    }
                    string worksheetName = activeSheet.Name;
                    _logger.LogSystemMessage($"Range Selection: Active worksheet found - {worksheetName}");
                }
                catch (Exception ex)
                {
                    _logger.LogError("Range Selection: Failed to access active worksheet", ex);
                    MessageBox.Show("Cannot access the active worksheet. Please ensure Excel is responding.", "Worksheet Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Make Excel visible and activate it
                _logger.LogSystemMessage("Range Selection: Setting Excel visibility and activation");
                _excelService.Visible = true;
                await Task.Delay(200); // Increased delay for visibility
                _logger.LogSystemMessage("Range Selection: Activating Excel window");
                _excelService.ActiveWindow?.Activate();
                IntPtr excelHwnd = new IntPtr(_excelService.Hwnd);
                _logger.LogSystemMessage($"Range Selection: Excel window handle - {excelHwnd}");
                SetForegroundWindow(excelHwnd);
                await Task.Delay(200); // Increased delay after activation

                // Get the current selection or default to A1
                Excel.Range? currentSelection = null;
                try
                {
                    _logger.LogSystemMessage("Range Selection: Getting current selection");
                    currentSelection = _excelService.Selection ?? activeSheet.Range["A1"];
                    string currentAddress = currentSelection.Address[false, false];
                    _logger.LogSystemMessage($"Range Selection: Current selection - {currentAddress}");
                }
                catch (Exception ex)
                {
                    _logger.LogError("Range Selection: Failed to get current selection", ex);
                    MessageBox.Show("Cannot access the current selection. Please try again.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                try
                {
                    _logger.LogSystemMessage("Range Selection: Showing Excel InputBox for range selection");
                    object? rangeResult = _excelService.InputBox(
                        "Select a range",
                        "Range Selection",
                        currentSelection?.Address[false, false] ?? "A1",
                        -1,
                        -1,
                        null,
                        null,
                        8
                    );
                    _logger.LogSystemMessage("Range Selection: InputBox dialog completed");

                    if (rangeResult is Excel.Range selectedRange)
                    {
                        _logger.LogSystemMessage("Range Selection: Range selected successfully");
                        this.selectedRange = selectedRange.Address[false, false];
                        this.selectedRangeContent = GetRangeContent(selectedRange);
                        _logger.LogSystemMessage($"Range Selection: Selected range - {this.selectedRange}");
                        _logger.LogSystemMessage($"Range Selection: Range content length - {this.selectedRangeContent?.Length ?? 0} characters");
                        UpdateSelectedRangeDisplay();

                        // Get current caret position
                        TextPointer caretPosition = InputTextBox.CaretPosition;
                        _logger.LogSystemMessage("Range Selection: Got caret position");

                        // Check if we need a space before
                        bool needSpaceBefore = false;
                        if (caretPosition.GetTextInRun(LogicalDirection.Backward)?.LastOrDefault() is char lastChar)
                        {
                            needSpaceBefore = !char.IsWhiteSpace(lastChar);
                        }

                        // Build the text to insert
                        string sheetName = activeSheet.Name;
                        string insertText = $"{(needSpaceBefore ? " " : "")}@{sheetName}!{this.selectedRange} ";
                        _logger.LogSystemMessage($"Range Selection: Inserting text - {insertText}");

                        // Insert the text at the caret position
                        TextRange insertRange = new TextRange(caretPosition, caretPosition);
                        insertRange.Text = insertText;
                        _logger.LogSystemMessage("Range Selection: Text inserted successfully");

                        // Move caret to after the inserted text
                        for (int i = 0; i < insertText.Length; i++)
                        {
                            caretPosition = caretPosition.GetPositionAtOffset(1) ?? caretPosition;
                        }
                        InputTextBox.CaretPosition = caretPosition;
                        _logger.LogSystemMessage("Range Selection: Caret position updated");

                        // Apply highlighting
                        HighlightAtSymbols();
                        _logger.LogSystemMessage("Range Selection: Highlighting applied");

                        // Focus back on the input box
                        InputTextBox.Focus();
                        _logger.LogSystemMessage("Range Selection: Input box focused");

                        if (selectedRange != currentSelection)
                        {
                            Marshal.ReleaseComObject(selectedRange);
                            _logger.LogSystemMessage("Range Selection: Released selected range COM object");
                        }
                    }
                    else
                    {
                        _logger.LogSystemMessage("Range Selection: No range selected (user cancelled)");
                    }
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    _logger.LogError("Range Selection: COM Exception during range selection", comEx);
                    Debug.WriteLine($"COM Exception in range selection: {comEx.Message}");
                    return;
                }
                finally
                {
                    _logger.LogSystemMessage("Range Selection: Cleaning up COM objects");
                    if (currentSelection != null)
                    {
                        Marshal.ReleaseComObject(currentSelection);
                        _logger.LogSystemMessage("Range Selection: Released current selection COM object");
                    }
                    if (activeSheet != null)
                    {
                        Marshal.ReleaseComObject(activeSheet);
                        _logger.LogSystemMessage("Range Selection: Released active sheet COM object");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Range Selection: Critical error", ex);
                MessageBox.Show($"Error selecting range: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _logger.LogSystemMessage("=== Range Selection Process Completed ===");
            }
        }

        private string GetRangeContent(Excel.Range range)
        {
            try
            {
                if (range == null) return string.Empty;

                var values = range.Value as object[,];
                if (values == null) return string.Empty;

                StringBuilder content = new StringBuilder();
                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    for (int col = 1; col <= values.GetLength(1); col++)
                    {
                        var value = values[row, col];
                        content.Append(value?.ToString() ?? "").Append("\t");
                    }
                    content.AppendLine();
                }
                return content.ToString().Trim();
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private void UpdateSelectedRangeDisplay()
        {
            if (string.IsNullOrEmpty(selectedRange))
            {
                SelectRangeButton.Content = "Select Range";
                return;
            }

            SelectRangeButton.Content = "Add New Range";
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Maximized;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OnMouseLeftButtonDown_(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void LoadRecentFiles()
        {
            try
            {
                var recentFiles = recentFilesService.GetRecentFiles();
                RecentFilesList.Items.Clear();
                foreach (var recentFile in recentFiles)
                {
                    var button = new Button
                    {
                        Style = (Style)FindResource("FileActionButton"),
                        Tag = recentFile.FullPath
                    };

                    var stackPanel = new StackPanel { Orientation = Orientation.Horizontal };
                    var pathIcon = new WinPath
                    {
                        Data = Geometry.Parse("M13,9V3.5L18.5,9M6,2C4.89,2 4,2.89 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2H6Z"),
                        Fill = (Brush)FindResource("PrimaryColor"),
                        Width = 24,
                        Height = 24,
                        Margin = new Thickness(0, 0, 10, 0)
                    };

                    var textBlock = new TextBlock
                    {
                        Text = recentFile.Name,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center
                    };

                    stackPanel.Children.Add(pathIcon);
                    stackPanel.Children.Add(textBlock);
                    button.Content = stackPanel;
                    button.Click += RecentFile_Click;

                    RecentFilesList.Items.Add(button);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading recent files: {ex.Message}");
            }
        }

        private async Task LoadFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            
            try
            {
                selectedFilePath = filePath;
                AddToRecentFiles(filePath);
                await LoadWorkbook();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddToRecentFiles(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            recentFilesService.AddToRecentFiles(filePath);
            LoadRecentFiles();
        }

        private async void RecentFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string filePath)
            {
                if (File.Exists(filePath))
                {
                    await LoadFile(filePath);
                }
                else
                {
                    MessageBox.Show($"File not found: {filePath}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    recentFilesService.RemoveRecentFile(filePath);
                    LoadRecentFiles();
                }
            }
        }

        private async void NewFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Create a new Excel service if it doesn't exist
                if (excelService == null)
                {
                    excelService = new ExcelService(_logger, Dispatcher.CurrentDispatcher);
                    excelService.WorkbookLoaded += ExcelService_WorkbookLoaded;
                    excelService.WorksheetActivated += ExcelService_WorksheetActivated;
                    excelService.ErrorOccurred += ExcelService_ErrorOccurred;
                    excelService.WorkbookClosed += ExcelService_WorkbookClosed;
                }

                // Update UI to show loading state
                UpdateSelectFileButtonUI("Creating...");

                // Create and load the new workbook
                bool success = await excelService.LoadWorkbookAsync(null, GenieForExcelFolderPath, true);
                if (!success)
                {
                    MessageBox.Show("Failed to create new workbook.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    UpdateSelectFileButtonUI("Select File");
                    return;
                }

                // Create a new workbook name
                string defaultName = $"New_Workbook_{DateTime.Now:yyyyMMdd_HHmmss}.xlsm";
                string newFilePath = IOPath.Combine(GenieForExcelFolderPath, defaultName);

                // Save the new workbook
                excelService.ActiveWorkbook?.SaveAs(newFilePath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                selectedFilePath = newFilePath;
                AddToRecentFiles(newFilePath);

                // Update UI state
                UpdateSelectFileButtonUI(TruncateFileName(IOPath.GetFileName(newFilePath)));
                InputTextBox.IsEnabled = true;
                GenerateButton.IsEnabled = true;
                IsFileSelected = true;
                IsSupportingFileSelectionEnabled = true;

                // Position windows after a delay
                await Task.Delay(1000);
                await Application.Current.Dispatcher.InvokeAsync(() => PositionWindowsAfterInit());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating new file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateSelectFileButtonUI("Select File");
            }
        }

        private void ChatScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // Check if we're at the bottom
            if (e.ExtentHeight != 0)
            {
                bool wasAtBottom = e.ExtentHeight - e.VerticalOffset - e.ViewportHeight <= 1;
                _autoScroll = wasAtBottom;
            }
        }

        private TextPointer GetTextPointerAtCharOffset(FlowDocument doc, int charOffset)
        {
            TextPointer pointer = doc.ContentStart;
            int charsTraversed = 0;

            while (pointer != null && pointer.CompareTo(doc.ContentEnd) < 0)
            {
                if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    string textRun = pointer.GetTextInRun(LogicalDirection.Forward);
                    if (charsTraversed + textRun.Length >= charOffset)
                    {
                        // The target position is within this run
                        return pointer.GetPositionAtOffset(charOffset - charsTraversed);
                    }
                    charsTraversed += textRun.Length;
                    pointer = pointer.GetPositionAtOffset(textRun.Length);
                }
                pointer = pointer.GetNextContextPosition(LogicalDirection.Forward);
            }

            return doc.ContentEnd;
        }

        private void HighlightAtSymbols()
        {
            if (isHighlighting) return;

            try
            {
                isHighlighting = true;
                var text = new TextRange(InputTextBox.Document.ContentStart, InputTextBox.Document.ContentEnd).Text;
                var doc = InputTextBox.Document;

                // If no @ symbols, don't process
                if (!text.Contains('@'))
                {
                    return;
                }

                doc.Blocks.Clear();
                var paragraph = new Paragraph();
                var currentPosition = 0;

                while (currentPosition < text.Length)
                {
                    var atIndex = text.IndexOf('@', currentPosition);
                    if (atIndex == -1)
                    {
                        // Add remaining text
                        if (currentPosition < text.Length)
                        {
                            paragraph.Inlines.Add(new Run(text.Substring(currentPosition)));
                        }
                        break;
                    }

                    // Add text before @
                    if (atIndex > currentPosition)
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(currentPosition, atIndex - currentPosition)));
                    }

                    // Find the end of the reference
                    var spaceIndex = text.IndexOf(' ', atIndex);
                    var newlineIndex = text.IndexOf('\n', atIndex);
                    var endIndex = text.Length;

                    if (spaceIndex != -1 && (newlineIndex == -1 || spaceIndex < newlineIndex))
                    {
                        endIndex = spaceIndex;
                    }
                    else if (newlineIndex != -1)
                    {
                        endIndex = newlineIndex;
                    }

                    // Add highlighted reference
                    var reference = text.Substring(atIndex, endIndex - atIndex);
                    var highlightedText = new Run(reference)
                    {
                        Foreground = new SolidColorBrush((MediaColor)ColorConverter.ConvertFromString("#0078D4"))
                    };
                    paragraph.Inlines.Add(highlightedText);

                    currentPosition = endIndex;
                    if (currentPosition < text.Length)
                    {
                        if (text[currentPosition] == ' ')
                        {
                            paragraph.Inlines.Add(new Run(" "));
                            currentPosition++;
                        }
                        else if (text[currentPosition] == '\n')
                        {
                            doc.Blocks.Add(paragraph);
                            paragraph = new Paragraph();
                            currentPosition++;
                        }
                    }
                }

                doc.Blocks.Add(paragraph);
            }
            finally
            {
                isHighlighting = false;
            }
        }

        private void InputTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isHighlighting) return;

            // Update HasText property first
            var text = new TextRange(InputTextBox.Document.ContentStart, InputTextBox.Document.ContentEnd).Text;
            HasText = !string.IsNullOrWhiteSpace(text);
            UpdateButtonState();

            // Only highlight if there's an @ symbol
            if (text.Contains('@'))
            {
                var caretPosition = InputTextBox.CaretPosition;
                var caretOffset = caretPosition.GetOffsetToPosition(InputTextBox.Document.ContentStart);
                
                HighlightAtSymbols();
                
                // Restore caret position
                var newPosition = GetTextPointerAtCharOffset(InputTextBox.Document, caretOffset);
                if (newPosition != null)
                {
                    InputTextBox.CaretPosition = newPosition;
                }
            }
        }

        private void HandlePathElement(UIElement element)
        {
            if (element is WinPath pathElement)
            {
                var data = pathElement.Data;
                // Use data...
            }
        }

        private void InputTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return && (Keyboard.Modifiers & ModifierKeys.Shift) != ModifierKeys.Shift)
            {
                e.Handled = true;
                GenerateButton_Click(sender, e);
            }
        }

        protected override async void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            await CheckInternetConnection(); // Initial check
            _internetCheckTimer.Start(); // Ensure timer is started
            await PositionWindowsAfterInit();
        }

        private async Task PositionWindowsAfterInit()
        {
            try
            {
                _logger.LogSystemMessage("Starting window positioning");

                // Get screen dimensions
                var screen = System.Windows.Forms.Screen.PrimaryScreen;
                if (screen == null)
                {
                    _logger.LogError("Could not get primary screen information");
                    return;
                }

                int screenWidth = screen.WorkingArea.Width;
                int screenHeight = screen.WorkingArea.Height;
                _logger.LogSystemMessage($"Screen dimensions: {screenWidth}x{screenHeight}");

                // Position Excel window (left side, 60% width)
                if (_excelService?.ActiveWorkbook != null)
                {
                    try
                    {
                        int excelWidth = (int)(screenWidth * 0.6);
                        IntPtr excelHwnd = new IntPtr(_excelService.Hwnd);
                        _logger.LogSystemMessage($"Excel window handle: {excelHwnd}, Target width: {excelWidth}");

                        var excelWindow = _excelService.ActiveWorkbook.Windows[1];
                        if (excelWindow != null)
                        {
                            try
                            {
                                // First, ensure Excel is not maximized
                                _excelService.WindowState = Excel.XlWindowState.xlNormal;
                                excelWindow.WindowState = Excel.XlWindowState.xlNormal;
                                await Task.Delay(100);

                                // Hide Excel while positioning
                                _excelService.Visible = false;
                                await Task.Delay(100);

                                // Reset window position and size
                                ShowWindow(excelHwnd, SW_RESTORE);
                                await Task.Delay(100);

                                // Set Excel window position at (0, 0) with the calculated width and full height
                                SetWindowPos(excelHwnd, HWND_TOP, 0, 0, excelWidth, screenHeight, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
                                await Task.Delay(100);

                                // Force window update
                                UpdateWindow(excelHwnd);
                                RedrawWindow(excelHwnd, IntPtr.Zero, IntPtr.Zero, RedrawWindowFlags.Invalidate | RedrawWindowFlags.Frame);

                                RECT rect;
                                GetWindowRect(excelHwnd, out rect);
                                _logger.LogSystemMessage($"Excel window position: Left={rect.Left}, Top={rect.Top}, Right={rect.Right}, Bottom={rect.Bottom}");

                                // Make Excel visible again
                                _excelService.Visible = true;
                                SetForegroundWindow(excelHwnd);

                                bool isConnected = await VerifyExcelConnection();
                                _logger.LogSystemMessage($"Excel connection status: {(isConnected ? "Connected" : "Not Connected")}");
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(excelWindow);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Error positioning Excel window", ex);
                    }
                }

                // Position Genie window (right side, 40% width)
                _logger.LogSystemMessage("Starting Genie window positioning");
                await Task.Delay(100);

                this.WindowState = WindowState.Normal;
                IntPtr genieHwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
                int genieX = (int)(screenWidth * 0.6);
                int genieWidth = (int)(screenWidth * 0.4);

                _logger.LogSystemMessage($"Setting Genie window position: X={genieX}, Width={genieWidth}, Height={screenHeight}");
                SetWindowPos(genieHwnd, HWND_TOP, genieX, 0, genieWidth, screenHeight, SWP_SHOWWINDOW | SWP_FRAMECHANGED);

                RECT genieRect;
                if (GetWindowRect(genieHwnd, out genieRect))
                {
                    _logger.LogSystemMessage($"Final Genie window position: Left={genieRect.Left}, Top={genieRect.Top}, Right={genieRect.Right}, Bottom={genieRect.Bottom}");
                }

                _logger.LogSystemMessage("Window positioning completed");
            }
            catch (Exception ex)
            {
                _logger.LogError("Error in PositionWindowsAfterInit", ex);
            }
        }

        private async Task<bool> VerifyExcelConnection()
        {
            try
            {
                if (_excelService == null || _excelService.ActiveWorkbook == null)
                {
                    _logger.LogSystemMessage("Excel connection check: No active workbook");
                    IsExcelConnected = false;
                    return false;
                }

                // Test 1: Check Excel process
                _logger.LogSystemMessage("Checking Excel process...");
                if (_excelService.Hwnd == 0)
                {
                    _logger.LogSystemMessage("Excel process check failed: No window handle");
                    IsExcelConnected = false;
                    return false;
                }

                // Test 2: Check workbook access
                _logger.LogSystemMessage("Checking workbook access...");
                string workbookName = _excelService.ActiveWorkbook.Name;
                _logger.LogSystemMessage($"Active workbook name: {workbookName}");

                // Test 3: Check all worksheets
                _logger.LogSystemMessage("Checking all worksheets...");
                var worksheets = _excelService.ActiveWorkbook.Worksheets;
                int worksheetCount = worksheets.Count;
                _logger.LogSystemMessage($"Total worksheets found: {worksheetCount}");

                // If all tests pass, update connection status
                IsExcelConnected = true;
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError("Excel connection verification failed", ex);
                IsExcelConnected = false;
                return false;
            }
        }

        // Add these Win32 API declarations at the class level
        [DllImport("user32.dll")]
        static extern bool UpdateWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_RESTORE = 9;
        const int SW_NORMAL = 1;
        const uint SWP_FRAMECHANGED = 0x0020;

        private void InitializeThemeButton()
        {
            try
            {
                if (ThemeToggleButton != null)
                {
                    ThemeToggleButton.Click -= ThemeToggleButton_Click; // Remove any existing handlers
                    ThemeToggleButton.Click += ThemeToggleButton_Click; // Add the click handler
                }
                isDarkTheme = true; // Set initial theme
                UpdateThemeIcon();
                UpdateTheme(); // Add initial theme update
            }
            catch (Exception ex)
            {
                _logger?.LogError("Error initializing theme button", ex);
            }
        }

        private void ThemeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                e.Handled = true;
                isDarkTheme = !isDarkTheme;
                
                // Update file action buttons first
                var noFileView = FindVisualChildren<Grid>(this).FirstOrDefault(g => g.Name == "NoFileView");
                if (noFileView != null)
                {
                    foreach (var button in FindVisualChildren<Button>(noFileView))
                    {
                        if (button.Style?.ToString()?.Contains("FileActionButton") == true)
                        {
                            button.Background = new SolidColorBrush(isDarkTheme ? 
                                (MediaColor)ColorConverter.ConvertFromString("#2D2D2D") : 
                                (MediaColor)ColorConverter.ConvertFromString("Transparent"));
                            button.BorderBrush = new SolidColorBrush(isDarkTheme ? 
                                (MediaColor)ColorConverter.ConvertFromString("#404040") : 
                                (MediaColor)ColorConverter.ConvertFromString("#E0E0E0"));
                        }
                    }
                }

                UpdateTheme();
                UpdateThemeIcon();
            }
            catch (Exception ex)
            {
                _logger?.LogError("Error in theme toggle button click", ex);
            }
        }

        private void UpdateThemeIcon()
        {
            try
            {
                if (ThemeIcon == null) return;

                // Update the icon color based on theme
                ThemeIcon.Stroke = new SolidColorBrush(isDarkTheme ? 
                    System.Windows.Media.Colors.White : 
                    (MediaColor)ColorConverter.ConvertFromString("#202020"));

                // Moon icon for dark theme, sun icon for light theme
                string iconPath = isDarkTheme
                    ? "M9.37 5.51c-.18.64-.27 1.31-.27 1.99 0 4.08 3.32 7.4 7.4 7.4.68 0 1.35-.09 1.99-.27C17.45 17.19 14.93 19 12 19c-3.86 0-7-3.14-7-7 0-2.93 1.81-5.45 4.37-6.49z" // Moon
                    : "M12 7c-2.76 0-5 2.24-5 5s2.24 5 5 5 5-2.24 5-5-2.24-5-5-5zM2 13h2c.55 0 1-.45 1-1s-.45-1-1-1H2c-.55 0-1 .45-1 1s.45 1 1 1zm18 0h2c.55 0 1-.45 1-1s-.45-1-1-1h-2c-.55 0-1 .45-1 1s.45 1 1 1zM11 2v2c0 .55.45 1 1 1s1-.45 1-1V2c0-.55-.45-1-1-1s-1 .45-1 1zm0 18v2c0 .55.45 1 1 1s1-.45 1-1v-2c0-.55-.45-1-1-1s-1 .45-1 1zM5.99 4.58c-.39-.39-1.03-.39-1.41 0-.39.39-.39 1.03 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0s.39-1.03 0-1.41L5.99 4.58zm12.37 12.37c-.39-.39-1.03-.39-1.41 0-.39.39-.39 1.03 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0 .39-.39.39-1.03 0-1.41l-1.06-1.06zm1.06-10.96c.39-.39.39-1.03 0-1.41-.39-.39-1.03-.39-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06zM7.05 18.36c.39-.39.39-1.03 0-1.41-.39-.39-1.03-.39-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06z"; // Sun

                ThemeIcon.Data = Geometry.Parse(iconPath);
            }
            catch (Exception ex)
            {
                _logger?.LogError("Error updating theme icon", ex);
            }
        }

        private void UpdateTheme()
        {
            var isDark = isDarkTheme;
            var backgroundColor = isDark ? "#1E1E1E" : "#FFFFFF";
            var foregroundColor = isDark ? "#FFFFFF" : "#000000";
            var borderColor = isDark ? "#3A3A3A" : "#E5E5E5";

            // Update window background
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(backgroundColor));

            // Update text colors
            foreach (var textBlock in FindVisualChildren<TextBlock>(this))
            {
                textBlock.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(foregroundColor));
            }

            // Update borders
            foreach (var border in FindVisualChildren<Border>(this))
            {
                border.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(borderColor));
            }

            // Update buttons
            foreach (var button in FindVisualChildren<Button>(this))
            {
                UpdateButtonTheme(button);
            }

            // Update paths
            foreach (var path in FindVisualChildren<WinPath>(this))
            {
                UpdatePathTheme(path);
            }
        }

        private void UpdateButtonTheme(Button button)
        {
            if (button.Content is string)
            {
                button.Foreground = new SolidColorBrush(isDarkTheme ? 
                    MediaColors.White : 
                    (MediaColor)ColorConverter.ConvertFromString("#202020"));
            }

            foreach (var buttonChild in LogicalTreeHelper.GetChildren(button))
            {
                if (buttonChild is WinPath path)
                {
                    UpdatePathTheme(path);
                }
            }
        }

        private void UpdatePathTheme(WinPath path)
        {
            if (!path.Stroke.ToString().Contains("#1D6F42"))
            {
                var color = isDarkTheme ? 
                    MediaColors.White : 
                    (MediaColor)ColorConverter.ConvertFromString("#202020");
                path.Stroke = new SolidColorBrush(color);
                path.Fill = new SolidColorBrush(color);
            }
        }

        private void ScrollLeft_Click(object sender, RoutedEventArgs e)
        {
            ScrollWorksheetList(-100); // Scroll left by 100 pixels
        }

        private void ScrollRight_Click(object sender, RoutedEventArgs e)
        {
            ScrollWorksheetList(100); // Scroll right by 100 pixels
        }

        private void ScrollWorksheetList(double offset)
        {
            if (WorksheetScrollViewer != null)
            {
                double newOffset = WorksheetScrollViewer.HorizontalOffset + offset;
                newOffset = Math.Max(0, Math.Min(newOffset, WorksheetScrollViewer.ScrollableWidth));
                WorksheetScrollViewer.ScrollToHorizontalOffset(newOffset);
                UpdateScrollButtonStates();
            }
        }

        private void UpdateScrollButtonStates()
        {
            if (WorksheetScrollViewer != null)
            {
                // Calculate if scrolling is possible
                bool canScroll = WorksheetScrollViewer.ScrollableWidth > 0;
                NavigationButtonsPanel.Visibility = canScroll ? Visibility.Visible : Visibility.Collapsed;

                if (canScroll)
                {
                    ScrollLeftButton.IsEnabled = WorksheetScrollViewer.HorizontalOffset > 0;
                    ScrollRightButton.IsEnabled = WorksheetScrollViewer.HorizontalOffset < WorksheetScrollViewer.ScrollableWidth;
                }
            }
        }

        // Add a handler for the ScrollViewer's ScrollChanged event
        private void WorksheetScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            UpdateScrollButtonStates();
        }

        // Add these event handlers for worksheet changes
        private void ExcelService_WorksheetRenamed(object? sender, (string oldName, string newName) e)
        {
            _logger?.LogSystemMessage($"Worksheet renamed from {e.oldName} to {e.newName}");
            UpdateWorksheetList();
        }

        private void ExcelService_WorksheetDeleted(object? sender, string worksheetName)
        {
            _logger?.LogSystemMessage($"Worksheet deleted: {worksheetName}");
            UpdateWorksheetList();
        }

        private void ModeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            isProMode = !isProMode;
            UpdateUIBasedOnSubscription();
        }

        private async void AddContext_Click(object sender, RoutedEventArgs e)
        {
            if (supportingFile != null)
            {
                ClearSupportingFile();
            }
            else
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Supported Files|*.docx;*.xlsx;*.pdf",
                    Title = "Select Context File"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    try
                    {
                        if (allowedExtensions.Contains(IOPath.GetExtension(openFileDialog.FileName).ToLower()))
                        {
                            await AddSupportingFile(openFileDialog.FileName);
                        }
                        else
                        {
                            MessageBox.Show($"Only .docx, .xlsx, and .pdf files are supported.", 
                                "Unsupported File Type", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing file: {ex.Message}", 
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private async Task CheckInternetConnection()
        {
            try
            {
                using var client = new HttpClient();
                var response = await client.GetAsync("http://www.google.com");
                NoInternetOverlay.Visibility = response.IsSuccessStatusCode ? Visibility.Collapsed : Visibility.Visible;
            }
            catch
            {
                NoInternetOverlay.Visibility = Visibility.Visible;
            }
        }

        private bool IsUserMessage(TextBlock textBlock)
        {
            var parent = VisualTreeHelper.GetParent(textBlock);
            while (parent != null)
            {
                if (parent is Border border)
                {
                    return border.Background.ToString().Contains("#1D6F42");
                }
                parent = VisualTreeHelper.GetParent(parent);
            }
            return false;
        }

        private bool IsUserMessageBorder(Border border)
        {
            return border.Background.ToString().Contains("#1D6F42");
        }

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj == null) yield break;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = VisualTreeHelper.GetChild(depObj, i);
                
                if (child is T t)
                    yield return t;

                foreach (T childOfChild in FindVisualChildren<T>(child))
                    yield return childOfChild;
            }
        }

        private void CustomInstructionsButton_Click(object sender, RoutedEventArgs e)
        {
            var customInstructionsWindow = new CustomInstructionsWindow(isDarkTheme);
            customInstructionsWindow.Owner = this;
            if (customInstructionsWindow.ShowDialog() == true)
            {
                // Update the chat manager with the new custom instructions
                chatManager.SetCustomInstructions(customInstructionsWindow.CustomInstructions);
            }
        }

        private void AuditFormulasButton_Click(object sender, RoutedEventArgs e)
        {
            string message = "Please audit my file and provide a comprehensive review of:\n" +
                           "1. Data structure and organization\n" +
                           "2. Formula usage and complexity\n" +
                           "3. Potential errors or inconsistencies\n" +
                           "4. Security and data validation\n" +
                           "5. Overall workbook health";

            // Clear the input box and add the message
            SetInputText(message);
            
            // Trigger the generate button click
            GenerateButton_Click(sender, e);
        }

        private void SuggestImprovementsButton_Click(object sender, RoutedEventArgs e)
        {
            string message = "Please analyze this workbook and suggest improvements for:\n" +
                           "1. Overall structure and organization\n" +
                           "2. Data layout and presentation\n" +
                           "3. Formula efficiency and maintainability\n" +
                           "4. Formatting and visual clarity\n" +
                           "5. Performance optimization";

            // Clear the input box and add the message
            SetInputText(message);
            
            // Trigger the generate button click
            GenerateButton_Click(sender, e);
        }

        private void MakePrintReadyButton_Click(object sender, RoutedEventArgs e)
        {
            string message = "Please analyze this workbook and suggest a plan for:\n" +
                           "1. Additional features or functionality\n" +
                           "2. New data analysis capabilities\n" +
                           "3. Automation opportunities\n" +
                           "4. Reporting enhancements\n" +
                           "5. Integration possibilities";

            // Clear the input box and add the message
            SetInputText(message);
            
            // Trigger the generate button click
            GenerateButton_Click(sender, e);
        }

        private void ThemeButton_Click(object sender, RoutedEventArgs e)
        {
            isDarkTheme = !isDarkTheme;
            UpdateThemeIcon();
        }

        private async void ConnectionStatusButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!IsExcelConnected)
                {
                    _logger.LogSystemMessage("\n=== Starting Excel Connection Process ===");
                    _logger.LogSystemMessage($"Current Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                    ShowLoadingIndicator();
                    
                    try
                    {
                        // Step 1: Create workbook
                        _logger.LogSystemMessage("\n=== Step 1: Creating workbook ===");
                        UpdateLoadingStep(1, "Creating Workbook...");
                        var workbookInfo = await _excelService.GetWorkbookInfo();
                        if (workbookInfo != null)
                        {
                            _logger.LogSystemMessage("Successfully retrieved workbook info");
                            var workbookRef = await _palantirService.CreateExcelWorkbook(workbookInfo);
                            _logger.LogSystemMessage($"Successfully created workbook:");
                            _logger.LogSystemMessage($"- Local GUID: {workbookRef.LocalGuid}");
                            _logger.LogSystemMessage($"- Palantir Primary Key: {workbookRef.PalantirPrimaryKey}");
                            await Task.Delay(500);

                            // Step 2: Process connections
                            _logger.LogSystemMessage("\n=== Step 2: Processing connections ===");
                            UpdateLoadingStep(2, "Processing Connections...");
                            try
                            {
                                var connections = await _excelService.GetWorkbookConnections();
                                _logger.LogSystemMessage($"Found {connections?.Count ?? 0} connections");
                                if (connections != null && connections.Count > 0)
                                {
                                    var workbook = _excelService.GetActiveWorkbook();
                                    if (workbook != null)
                                    {
                                        await _palantirService.CreateExcelConnections(workbook, workbookRef.LocalGuid);
                                        _logger.LogSystemMessage("Successfully processed connections");
                                    }
                                }
                                else
                                {
                                    _logger.LogSystemMessage("No connections found to process");
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError($"Error processing connections: {ex.Message}");
                            }
                            await Task.Delay(500);

                            // Step 3: Process worksheets and cells
                            _logger.LogSystemMessage("\n=== Step 3: Processing worksheets and cells ===");
                            try
                            {
                                await ProcessWorksheets();
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError($"Error processing worksheets and cells: {ex.Message}");
                            }

                            // Step 4: Generate Workbook Summary
                            _logger.LogSystemMessage("\n=== Step 4: Generating Workbook Summary ===");
                            UpdateLoadingStep(4, "Generating Workbook Summary...");
                            try
                            {
                                await _palantirService.GenerateWorkbookSummary(workbookRef.PalantirPrimaryKey);
                                _logger.LogSystemMessage("Workbook Summary Generation Status: Started");
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError($"Error generating workbook summary: {ex.Message}");
                            }

                            _logger.LogSystemMessage("\n=== Excel Connection Process Complete ===");
                            HideLoadingIndicator();
                            IsExcelConnected = true;
                            UpdateConnectionButton();
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error during Excel connection process: {ex.Message}");
                        HideLoadingIndicator();
                        MessageBox.Show($"Error connecting to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    await CleanupExcelProcesses();
                    IsExcelConnected = false;
                    UpdateConnectionButton();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ConnectionStatusButton_Click: {ex.Message}");
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void UpdateLoadingStep(int step, string operation = "")
        {
            string stepText = step switch
            {
                1 => "Creating Workbook...",
                2 => "Processing Connections...",
                3 => "Processing Worksheets and Cells...",
                4 => "Generating Workbook Summary...",
                _ => "Processing..."
            };

            if (!string.IsNullOrEmpty(operation))
            {
                stepText = operation;
            }

            _loadingManager.UpdateLoadingText($"Step {step}/4: {stepText}");
            double progress = (step - 1) * (100.0 / 4);
            _loadingManager.UpdateProgress(progress);
        }

        private void ShowLoadingIndicator()
        {
            Dispatcher.Invoke(() =>
            {
                // Reset progress bar and initial text
                LoadingProgressBar.Value = 0;
                CurrentOperationLabel.Text = "Starting indexing process...";
                PalantirProcessingOverlay.Visibility = Visibility.Visible;
            });
        }

        public void HideLoadingIndicator()
        {
            Dispatcher.Invoke(() =>
            {
                PalantirProcessingOverlay.Visibility = Visibility.Collapsed;
                LoadingProgressBar.Value = 0;
                CurrentOperationLabel.Text = "Ready";
            });
        }

        private async void SimulateLoadingSteps()
        {
            try
            {
                // Step 1: Workbook Structure
                for (int i = 0; i <= 33; i++)
                {
                    await Task.Delay(30);
                    UpdateLoadingStep(1, i == 33 ? "Successfully indexed workbook structure" : "");
                }

                // Step 2: Connections
                for (int i = 0; i <= 33; i++)
                {
                    await Task.Delay(30);
                    UpdateLoadingStep(2, i == 33 ? "Successfully indexed connections" : "");
                }

                // Step 3: Worksheets and Cells
                for (int i = 0; i <= 34; i++)
                {
                    await Task.Delay(30);
                    UpdateLoadingStep(3, i == 34 ? "Successfully indexed worksheets and cells" : "");
                }

                // Hide after a short delay
                await Task.Delay(500);
                HideLoadingIndicator();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during loading simulation: {ex.Message}");
                HideLoadingIndicator();
            }
        }

        private async Task ProcessUserInput(object sender, RoutedEventArgs e)
        {
            var userInput = GetInputText();
            if (string.IsNullOrWhiteSpace(userInput))
            {
                return;
            }

            // Clear the input box and add the message
            SetInputText("");
            
            // Trigger the generate button click
            GenerateButton_Click(sender, e);
        }

        private async Task ProcessWorksheets()
        {
            _logger.LogSystemMessage("=== Starting Worksheet Processing ===");
            try
            {
                _logger.LogInformation("Starting worksheet processing...");
                var workbook = _excelService.GetActiveWorkbook();
                if (workbook == null)
                {
                    _logger.LogError("No active workbook found.");
                    return;
                }

                _loadingManager.UpdateLoadingStep(3);
                var workbookRef = await _palantirService.CreateExcelWorkbook(new Dictionary<string, string> { { "workbook_path", workbook.FullName } });
                await _palantirService.CreateExcelWorksheets(workbook, workbookRef);
                _logger.LogInformation("Successfully processed worksheets and cells.");

                _loadingManager.UpdateLoadingStep(4);
                await _palantirService.GenerateWorkbookSummary(workbookRef.PalantirPrimaryKey);
                _logger.LogInformation("Successfully initiated workbook summary generation.");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error processing worksheets: {ex.Message}", ex);
                throw;
            }
        }

        private int _totalCellCount = 0;
        private ObservableCollection<string> _processingWorksheets = new ObservableCollection<string>();
        private ObservableCollection<string> _processingCharts = new ObservableCollection<string>();

        public void ShowProcessingPanel()
        {
            PalantirProcessingOverlay.Visibility = Visibility.Visible;
            ProcessingDetailsPanel.Visibility = Visibility.Visible;
            _totalCellCount = 0;
            _processingWorksheets.Clear();
            _processingCharts.Clear();
            ProcessingWorksheetList.ItemsSource = _processingWorksheets;
            ProcessingChartList.ItemsSource = _processingCharts;
        }

        public void HideProcessingPanel()
        {
            PalantirProcessingOverlay.Visibility = Visibility.Collapsed;
            ProcessingDetailsPanel.Visibility = Visibility.Collapsed;
        }

        public void UpdateWorkbookInfo(string workbookName)
        {
            WorkbookInfoLabel.Text = $"Processing Workbook: {workbookName}";
        }

        public void AddProcessingWorksheet(string worksheetName)
        {
            _processingWorksheets.Add($"Worksheet: {worksheetName}");
        }

        public void UpdateCellCount(int count)
        {
            _totalCellCount += count;
            CellCountLabel.Text = $"Total Cells Processed: {_totalCellCount:N0}";
        }

        public void AddProcessingChart(string chartName)
        {
            _processingCharts.Add($"Chart: {chartName}");
        }

        public void UpdateCurrentOperation(string operation)
        {
            CurrentOperationLabel.Text = operation;
        }

        public void UpdateLoadingProgress(double progress)
        {
            LoadingProgressBar.Value = progress;
        }

        private async Task ProcessWorkbook(Excel.Workbook workbook, PalantirService.EntityReference workbookRef)
        {
            try
            {
                ShowProcessingPanel();
                UpdateWorkbookInfo(workbook.Name);
                UpdateCurrentOperation("Processing workbook structure...");
                UpdateLoadingProgress(10);

                // Process worksheets
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    UpdateCurrentOperation($"Processing worksheet: {worksheet.Name}");
                    AddProcessingWorksheet(worksheet.Name);

                    var worksheetRef = await _palantirService.CreateExcelWorksheet(worksheet, workbookRef.PalantirPrimaryKey);
                    UpdateLoadingProgress(30);

                    // Process cells
                    Excel.Range usedRange = worksheet.UsedRange;
                    if (usedRange != null)
                    {
                        int cellCount = usedRange.Cells.Count;
                        UpdateCellCount(cellCount);
                        await _palantirService.ProcessWorksheetCells(worksheet, worksheetRef.PalantirPrimaryKey);
                    }
                    UpdateLoadingProgress(60);

                    // Process charts
                    Excel.ChartObjects chartObjects = worksheet.ChartObjects() as Excel.ChartObjects;
                    if (chartObjects != null && chartObjects.Count > 0)
                    {
                        foreach (Excel.ChartObject chartObject in chartObjects)
                        {
                            AddProcessingChart(chartObject.Name);
                        }
                        await _palantirService.ProcessWorksheetCharts(worksheet, worksheetRef.PalantirPrimaryKey);
                    }
                    UpdateLoadingProgress(80);
                }

                UpdateCurrentOperation("Generating workbook summary...");
                await _palantirService.GenerateWorkbookSummary(workbookRef.PalantirPrimaryKey);
                UpdateLoadingProgress(100);

                UpdateCurrentOperation("Processing complete!");
                await Task.Delay(2000); // Show completion message briefly
                HideProcessingPanel();
            }
            catch (Exception ex)
            {
                UpdateCurrentOperation($"Error: {ex.Message}");
                _logger.LogError("Error processing workbook", ex);
                throw;
            }
        }

        private void SetupDragAndDrop()
        {
            if (DragDropArea != null)
            {
                DragDropArea.AllowDrop = true;
                DragDropArea.Drop += DragDropArea_Drop;
                DragDropArea.DragEnter += DragDropArea_DragEnter;
                DragDropArea.DragLeave += DragDropArea_DragLeave;
            }
        }

        private void SetupWindowChrome()
        {
            WindowChrome.SetWindowChrome(this, new WindowChrome
            {
                CaptionHeight = 32,
                ResizeBorderThickness = new Thickness(6),
                CornerRadius = new CornerRadius(0),
                GlassFrameThickness = new Thickness(0),
                UseAeroCaptionButtons = false
            });
        }

        private void SetupTheme()
        {
            isDarkTheme = true; // Default to dark theme
            UpdateTheme();
        }

        private void WorkbookExplorer_WorkbookSelected(object sender, RoutedEventArgs e)
        {
            if (sender is ViewModels.WorkbookViewModel workbook)
            {
                // For now, just log the selection
                _logger.LogSystemMessage($"Workbook selected: {workbook.Name}");
            }
        }

        private void WorkbookExplorer_WorksheetSelected(object sender, RoutedEventArgs e)
        {
            if (sender is ViewModels.WorksheetViewModel worksheet)
            {
                // Activate the selected worksheet
                _excelService?.ActivateWorksheet(worksheet.Name);
                _logger.LogSystemMessage($"Worksheet selected: {worksheet.Name}");
            }
        }

        private void WorkbookExplorer_ChartSelected(object sender, RoutedEventArgs e)
        {
            if (sender is ViewModels.ChartViewModel chart)
            {
                // For now, just log the selection
                _logger.LogSystemMessage($"Chart selected: {chart.Name}");
            }
        }

        private async void ProcessDroppedFile(string filePath)
        {
            var extension = IOPath.GetExtension(filePath).ToLower();
            if (extension == ".xlsx" || extension == ".xls" || extension == ".xlsm")
            {
                await LoadFile(filePath);
            }
            else
            {
                MessageBox.Show("Please drop a valid Excel file (.xlsx, .xls, or .xlsm)", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private string GetRelativePath(string fullPath)
        {
            string currentDirectory = Environment.CurrentDirectory;
            if (fullPath.StartsWith(currentDirectory))
            {
                return IOPath.GetRelativePath(currentDirectory, fullPath);
            }
            return fullPath;
        }

    }
}
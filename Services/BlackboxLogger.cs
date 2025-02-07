using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Diagnostics;

namespace ExcelGenie.Services
{
    public class BlackboxLogger
    {
        private readonly object _logLock = new object();
        private string? _logFilePath;
        private readonly StringBuilder _pendingLogs = new StringBuilder();
        private readonly Timer _flushTimer;
        private bool _isDisposed;

        public BlackboxLogger()
        {
            // Create a timer to flush logs every second
            _flushTimer = new Timer(FlushPendingLogs, null, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1));
        }

        public void SetLogFilePath(string logFilePath)
        {
            Debug.WriteLine($"BlackboxLogger: Setting log file path to {logFilePath}");
            try
            {
                lock (_logLock)
                {
                    // Ensure the directory exists
                    string? directory = Path.GetDirectoryName(logFilePath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Debug.WriteLine($"BlackboxLogger: Creating directory {directory}");
                        Directory.CreateDirectory(directory);
                    }

                    // Flush any pending logs to the old file
                    if (_logFilePath != null)
                    {
                        Debug.WriteLine("BlackboxLogger: Flushing pending logs to old file");
                        FlushPendingLogs(null);
                    }

                    _logFilePath = logFilePath;
                    
                    // Create the initial log entry with new timestamp format
                    string formattedTime = DateTime.Now.ToString("hh:mm:sstt").ToLower().Replace(" ", "") + ":";
                    string header = $"=== Log Session Started at {formattedTime} ===\r\n" +
                                  $"Log File: {Path.GetFileName(logFilePath)}\r\n" +
                                  "============================================\r\n\r\n";
                    
                    Debug.WriteLine("BlackboxLogger: Writing header to new log file");
                    File.WriteAllText(_logFilePath, header);
                    Debug.WriteLine($"BlackboxLogger: Header written successfully to {_logFilePath}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"BlackboxLogger Error: Failed to set log file path: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }

        public void LogMethodStart(string methodName, string details)
        {
            AppendLog("START", $"[{methodName}] {details}");
        }

        public void LogMethodSuccess(string methodName, string details)
        {
            AppendLog("SUCCESS", $"[{methodName}] {details}");
        }

        public void LogProgress(string methodName, string details)
        {
            AppendLog("PROGRESS", $"[{methodName}] {details}");
        }

        public void LogUserMessage(string message)
        {
            AppendLog("USER", message);
        }

        public void LogSystemMessage(string message)
        {
            AppendLog("SYSTEM", message);
        }

        public void LogInformation(string message)
        {
            AppendLog("INFO", message);
        }

        public void LogError(string error, Exception? ex = null, string? methodName = null)
        {
            var errorBuilder = new StringBuilder();
            if (methodName != null)
            {
                errorBuilder.AppendLine($"[{methodName}] {error}");
            }
            else
            {
                errorBuilder.AppendLine(error);
            }

            if (ex != null)
            {
                errorBuilder.AppendLine($"Exception Type: {ex.GetType().FullName}");
                errorBuilder.AppendLine($"Message: {ex.Message}");
                if (ex is System.Runtime.InteropServices.COMException comEx)
                {
                    errorBuilder.AppendLine($"COM Error Code: 0x{comEx.ErrorCode:X8}");
                }
                errorBuilder.AppendLine($"Stack Trace: {ex.StackTrace}");
                
                if (ex.InnerException != null)
                {
                    errorBuilder.AppendLine("Inner Exception:");
                    errorBuilder.AppendLine($"Type: {ex.InnerException.GetType().FullName}");
                    errorBuilder.AppendLine($"Message: {ex.InnerException.Message}");
                    errorBuilder.AppendLine($"Stack Trace: {ex.InnerException.StackTrace}");
                }
            }

            AppendLog("ERROR", errorBuilder.ToString());
        }

        public void LogApiRequest(string endpoint, string requestJson, string methodName)
        {
            var logMessage = $"[{methodName}] Endpoint: {endpoint}\r\nRequest:\r\n{requestJson}";
            AppendLog("API REQUEST", logMessage);
        }

        public void LogApiResponse(string endpoint, string responseJson, string methodName)
        {
            var logMessage = $"[{methodName}] Endpoint: {endpoint}\r\nResponse:\r\n{responseJson}";
            AppendLog("API RESPONSE", logMessage);
        }

        public void LogVbaCode(string vbaCode, bool isGenerated = true)
        {
            string type = isGenerated ? "GENERATED VBA CODE" : "EXECUTED VBA CODE";
            AppendLog(type, vbaCode);
        }

        public void LogExcelOperation(string operation, string details, string methodName)
        {
            AppendLog("EXCEL", $"[{methodName}] {operation}: {details}");
        }

        private void AppendLog(string type, string message)
        {
            if (_isDisposed) return;

            try
            {
                // Change timestamp format to "10:27:30pm:"
                string timestamp = DateTime.Now.ToString("hh:mm:sstt")  // "10:27:30 PM"
                                    .ToLower()                  // "10:27:30 pm"
                                    .Replace(" ", "")           // "10:27:30pm"
                                 + ":";                         // "10:27:30pm:"

                string formattedMessage = FormatLogMessage(timestamp, type, message);

                lock (_logLock)
                {
                    if (_logFilePath == null)
                    {
                        Debug.WriteLine("BlackboxLogger Warning: No log file path set");
                        _pendingLogs.Append(formattedMessage);
                        return;
                    }

                    try
                    {
                        // Write directly to file instead of buffering
                        Debug.WriteLine($"BlackboxLogger: Writing log entry of type {type} to {_logFilePath}");
                        File.AppendAllText(_logFilePath, formattedMessage);
                        Debug.WriteLine("BlackboxLogger: Log entry written successfully");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"BlackboxLogger Error: Failed to write to log file: {ex.Message}");
                        Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                        // Store in pending logs if file write fails
                        _pendingLogs.Append(formattedMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"BlackboxLogger Error: Failed to append log: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private string FormatLogMessage(string timestamp, string type, string message)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"[{timestamp}] [{type}]");
            
            // Split message into lines and indent them
            string[] lines = message.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            foreach (string line in lines)
            {
                builder.AppendLine($"    {line}");
            }
            
            builder.AppendLine(); // Add a blank line between entries
            return builder.ToString();
        }

        private void FlushPendingLogs(object? state)
        {
            if (_isDisposed) return;

            lock (_logLock)
            {
                if (_logFilePath == null || _pendingLogs.Length == 0) return;

                try
                {
                    Debug.WriteLine($"BlackboxLogger: Flushing {_pendingLogs.Length} characters to {_logFilePath}");
                    File.AppendAllText(_logFilePath, _pendingLogs.ToString());
                    _pendingLogs.Clear();
                    Debug.WriteLine("BlackboxLogger: Flush completed successfully");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"BlackboxLogger Error: Failed to flush logs: {ex.Message}");
                    Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                }
            }
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            
            _isDisposed = true;
            _flushTimer?.Dispose();
            
            // Final flush of any pending logs
            FlushPendingLogs(null);
        }
    }
} 
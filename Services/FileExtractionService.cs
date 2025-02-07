using System;
using System.Text;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;

namespace ExcelGenie.Services
{
    public class FileExtractionService : IDisposable
    {
        private readonly string[] allowedExtensions = { ".docx", ".xlsx", ".pdf", ".pptx", ".xls" };
        private bool _disposed;
        private readonly HttpClient _httpClient;
        private readonly BlackboxLogger _logger;

        public class GenerationResult
        {
            public string Description { get; set; } = string.Empty;
            public string Code { get; set; } = string.Empty;
        }

        public FileExtractionService(BlackboxLogger logger)
        {
            _logger = logger;
            _httpClient = new HttpClient();
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
                    _httpClient.Dispose();
                }
                _disposed = true;
            }
        }

        ~FileExtractionService()
        {
            Dispose(false);
        }

        public async Task<GenerationResult> GenerateVBACode(
            string userInput,
            string worksheetData,
            string supportingFileContent,
            string? selectedRange = null,
            string? selectedRangeContent = null,
            string? activeWorksheetName = null,
            List<(string message, bool isUser)>? conversationHistory = null,
            string? customInstructions = null)
        {
            // For now, return empty result since we removed OpenAI service
            return new GenerationResult
            {
                Description = "Service temporarily unavailable",
                Code = string.Empty
            };
        }

        public bool IsFileTypeSupported(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLower();
            return Array.Exists(allowedExtensions, ext => ext == extension);
        }

        public async Task<string> ExtractTextFromFile(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The specified file does not exist.", filePath);

            string extension = Path.GetExtension(filePath).ToLower();
            if (!IsFileTypeSupported(extension))
                throw new NotSupportedException($"File type {extension} is not supported.");

            return extension switch
            {
                ".pdf" => await Task.Run(() => ExtractTextFromPdf(filePath)),
                ".docx" => await Task.Run(() => ExtractTextFromDocx(filePath)),
                ".xlsx" => await Task.Run(() => ExtractTextFromXlsx(filePath)),
                ".pptx" => await Task.Run(() => ExtractTextFromPptx(filePath)),
                ".xls" => await Task.Run(() => ExtractTextFromXls(filePath)),
                _ => throw new NotSupportedException($"File type {extension} is not supported.")
            };
        }

        private string ExtractTextFromPdf(string filePath)
        {
            StringBuilder text = new StringBuilder();
            using (PdfReader pdfReader = new PdfReader(filePath))
            using (PdfDocument pdfDoc = new PdfDocument(pdfReader))
            {
                for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                    text.AppendLine(pageText);
                }
            }
            return text.ToString();
        }

        private string ExtractTextFromDocx(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                Body? body = doc.MainDocumentPart?.Document.Body;
                return body?.InnerText ?? string.Empty;
            }
        }

        private string ExtractTextFromPptx(string filePath)
        {
            StringBuilder text = new StringBuilder();
            using (PresentationDocument ppt = PresentationDocument.Open(filePath, false))
            {
                if (ppt.PresentationPart?.Presentation?.SlideIdList != null)
                {
                    foreach (SlideId slideId in ppt.PresentationPart.Presentation.SlideIdList)
                    {
                        var relationshipId = slideId?.RelationshipId;
                        if (string.IsNullOrEmpty(relationshipId)) continue;
                        
                        var part = ppt.PresentationPart.GetPartById(relationshipId!);
                        if (part == null) continue;
                        
                        SlidePart? slide = part as SlidePart;
                        if (slide?.Slide != null)
                        {
                            text.AppendLine(slide.Slide.InnerText);
                        }
                    }
                }
            }
            return text.ToString();
        }

        private string ExtractTextFromXlsx(string filePath)
        {
            StringBuilder text = new StringBuilder();
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = spreadsheet.WorkbookPart;
                if (workbookPart?.Workbook?.Sheets == null) return string.Empty;

                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                foreach (Sheet sheet in workbookPart.Workbook.Sheets.OfType<Sheet>())
                {
                    string? id = sheet.Id?.Value;
                    if (string.IsNullOrEmpty(id)) continue;

                    var worksheetPart = workbookPart.GetPartById(id) as WorksheetPart;
                    if (worksheetPart?.Worksheet == null) continue;

                    var cells = worksheetPart.Worksheet.Descendants<Cell>();
                    foreach (Cell cell in cells)
                    {
                        string value = GetCellValue(cell, sharedStringTable);
                        if (!string.IsNullOrEmpty(value))
                        {
                            text.AppendLine($"{cell.CellReference}: {value}");
                        }
                    }
                }
            }
            return text.ToString();
        }

        private string ExtractTextFromXls(string filePath)
        {
            StringBuilder text = new StringBuilder();
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(file);
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet? sheet = workbook.GetSheetAt(i);
                    if (sheet != null)
                    {
                        for (int row = 0; row <= sheet.LastRowNum; row++)
                        {
                            IRow? sheetRow = sheet.GetRow(row);
                            if (sheetRow != null)
                            {
                                for (int col = 0; col < sheetRow.LastCellNum; col++)
                                {
                                    ICell? cell = sheetRow.GetCell(col);
                                    if (cell != null)
                                    {
                                        string cellRef = new CellReference(row, col).FormatAsString();
                                        string cellValue = GetCellValueAsString(cell);
                                        if (!string.IsNullOrEmpty(cellValue))
                                        {
                                            text.AppendLine($"{cellRef}: {cellValue}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return text.ToString();
        }

        private string GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
        {
            if (cell?.DataType == null || cell.CellValue?.Text == null) return string.Empty;
            
            if (cell.DataType == CellValues.SharedString && sharedStringTable != null)
            {
                if (int.TryParse(cell.CellValue.Text, out int ssid))
                {
                    var element = sharedStringTable.ElementAt(ssid);
                    return element?.InnerText ?? string.Empty;
                }
            }
            return cell.CellValue.Text;
        }

        private string GetCellValueAsString(ICell? cell)
        {
            if (cell == null) return string.Empty;
            
            return cell.CellType switch
            {
                NPOI.SS.UserModel.CellType.Numeric => cell.NumericCellValue.ToString(),
                NPOI.SS.UserModel.CellType.String => cell.StringCellValue ?? string.Empty,
                NPOI.SS.UserModel.CellType.Boolean => cell.BooleanCellValue.ToString(),
                NPOI.SS.UserModel.CellType.Formula => cell.CellFormula ?? string.Empty,
                _ => string.Empty
            };
        }
    }
} 
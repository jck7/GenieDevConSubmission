namespace ExcelGenie.Models
{
    public class UndoActionInfo
    {
        public string? BackupFilePath { get; set; }
        public int ConversationHistoryCount { get; set; }
        public int ChatPanelChildrenCount { get; set; }
    }
} 
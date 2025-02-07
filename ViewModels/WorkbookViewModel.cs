using System.Collections.ObjectModel;

namespace ExcelGenie.ViewModels
{
    public class WorkbookViewModel
    {
        public required string Name { get; set; }
        public required string PalantirId { get; set; }
        public ObservableCollection<WorksheetViewModel> Worksheets { get; } = new();
    }
} 
using System.Collections.ObjectModel;

namespace ExcelGenie.ViewModels
{
    public class WorksheetViewModel
    {
        public required string Name { get; set; }
        public required string PalantirId { get; set; }
        public ObservableCollection<ChartViewModel> Charts { get; } = new();
    }
} 
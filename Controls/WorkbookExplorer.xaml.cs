using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using ExcelGenie.ViewModels;

namespace ExcelGenie.Controls
{
    public partial class WorkbookExplorer : UserControl
    {
        public ObservableCollection<WorkbookViewModel> Workbooks { get; set; }
            = new ObservableCollection<WorkbookViewModel>();

        // Events for selection
        public event RoutedEventHandler? WorkbookSelected;
        public event RoutedEventHandler? WorksheetSelected;
        public event RoutedEventHandler? ChartSelected;

        public WorkbookExplorer()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void ObjectTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is WorkbookViewModel wb)
            {
                WorkbookSelected?.Invoke(wb, new RoutedEventArgs());
            }
            else if (e.NewValue is WorksheetViewModel ws)
            {
                WorksheetSelected?.Invoke(ws, new RoutedEventArgs());
            }
            else if (e.NewValue is ChartViewModel chart)
            {
                ChartSelected?.Invoke(chart, new RoutedEventArgs());
            }
        }
    }
} 
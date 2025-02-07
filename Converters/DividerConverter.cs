using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelGenie.Converters
{
    public class DividerConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isDarkTheme)
            {
                return new SolidColorBrush(isDarkTheme ? Color.FromRgb(42, 42, 42) : Color.FromRgb(224, 224, 224));
            }
            return new SolidColorBrush(Color.FromRgb(42, 42, 42));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 
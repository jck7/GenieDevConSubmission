using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelGenie.Converters
{
    public class BorderHoverConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isDarkTheme)
            {
                return new SolidColorBrush(isDarkTheme ? Color.FromRgb(74, 74, 74) : Color.FromRgb(208, 208, 208));
            }
            return new SolidColorBrush(Color.FromRgb(74, 74, 74));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 
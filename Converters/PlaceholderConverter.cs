using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelGenie.Converters
{
    public class PlaceholderConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isDarkTheme)
            {
                return new SolidColorBrush(isDarkTheme ? 
                    Color.FromRgb(80, 80, 80) : 
                    Color.FromRgb(160, 160, 160));
            }
            return new SolidColorBrush(Color.FromRgb(80, 80, 80));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 
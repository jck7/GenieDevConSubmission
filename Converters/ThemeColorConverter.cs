using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelGenie.Converters
{
    public class ThemeColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isDarkTheme)
            {
                return isDarkTheme ? new SolidColorBrush(Colors.White) : new SolidColorBrush(Colors.Black);
            }
            return new SolidColorBrush(Colors.Black);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 
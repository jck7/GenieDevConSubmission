using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelGenie.Converters
{
    public class BackgroundConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isDarkTheme)
            {
                return isDarkTheme ? new SolidColorBrush(Color.FromRgb(30, 30, 30)) : new SolidColorBrush(Colors.White);
            }
            return new SolidColorBrush(Colors.White);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 
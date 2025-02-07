using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelGenie.Converters
{
    public class BorderConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isDarkTheme)
            {
                return new SolidColorBrush(isDarkTheme ? 
                    Color.FromRgb(64, 64, 64) : 
                    Color.FromRgb(224, 224, 224));
            }
            return new SolidColorBrush(Color.FromRgb(64, 64, 64));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 
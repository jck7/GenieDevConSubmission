using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace ExcelGenie.Converters
{
    public class FirstLineConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string text)
            {
                // Split on newline characters and take the first segment.
                var firstLine = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).FirstOrDefault();
                return firstLine ?? text;
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 
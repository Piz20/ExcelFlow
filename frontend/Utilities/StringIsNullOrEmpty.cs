using System;
using System.Globalization;
using System.Windows.Data;

namespace ExcelFlow.Utilities // Ensure this namespace matches your using statement in GenerationView.xaml.cs
{
    public class StringIsNullOrEmptyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Returns true if the string is null or empty, false otherwise.
            return string.IsNullOrEmpty(value as string);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
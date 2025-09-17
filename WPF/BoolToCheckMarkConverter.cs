using System;
using System.Globalization;
using System.Windows.Data;

// 네임스페이스를 WPF와 통일합니다.
namespace PureGIS_Geo_QC.WPF
{
    public class BoolToCheckMarkConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool booleanValue)
            {
                return booleanValue ? "✓" : "✗";
            }
            return "✗";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
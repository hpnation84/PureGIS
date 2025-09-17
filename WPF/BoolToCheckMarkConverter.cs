using System;
using System.Globalization;
using System.Windows.Data;
using PureGIS_Geo_QC.Exports.Models;
using PureGIS_Geo_QC.Models;

namespace PureGIS_Geo_QC.WPF
{
    /// <summary>
    /// Boolean 값을 체크마크 문자로 변환하는 컨버터
    /// </summary>
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

    /// <summary>
    /// ColumnValidationResult를 비고 메시지로 변환하는 컨버터
    /// </summary>
    public class RemarksConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ColumnValidationResult result)
            {
                return ReportData.GetRemarks(result);
            }
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
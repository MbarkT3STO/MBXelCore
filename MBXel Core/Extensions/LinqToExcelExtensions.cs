using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace MBXel_Core.Extensions
{
    /// <summary>
    /// Represent a bunch of extension methods for LinqToExcel
    /// </summary>
    public static class LinqToExcelExtensions
    {
        public static int ToInt(this Cell cell) => int.Parse(cell.Value.ToString());
        public static double ToDouble(this Cell cell) => double.Parse(cell.Value.ToString());
        public static float ToFloat(this Cell cell) => float.Parse(cell.Value.ToString());
        public static decimal ToDecimal(this Cell cell) => decimal.Parse(cell.Value.ToString());
        public static DateTime ToDateTime(this Cell cell) => DateTime.Parse(cell.Value.ToString());
        public static TimeSpan ToTime(this Cell cell) => TimeSpan.Parse(cell.Value.ToString());
    }
}

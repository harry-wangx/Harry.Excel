using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Harry.Extensions.EPPlus
{
    public static class EnumerableExtensions
    {
        public static void WriteToExcel<T>(this IEnumerable<T> data, Stream stream, Action<EnumerableDataOptions<T>> optionsAction = null)
        {
            Check.NotNull(data, nameof(data));
            Check.NotNull(stream, nameof(stream));

            using (ExcelPackage doc = new ExcelPackage())
            {
                doc.LoadFromEnumerableData(data, optionsAction);
                doc.SaveAs(stream);
            }
        }
        public static ExcelPackage WriteToExcel<T>(this IEnumerable<T> data, Action<EnumerableDataOptions<T>> optionsAction = null)
        {
            Check.NotNull(data, nameof(data));

            return new ExcelPackage().LoadFromEnumerableData(data, optionsAction);
        }
    }
}

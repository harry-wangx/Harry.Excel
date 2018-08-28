﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Harry.Extensions.EPPlus
{
    public static class DataTableExtensions
    {
        public static void WriteToExcel(this DataTable dt, Stream stream, Action<DataTableOptions> dtOptionsAction = null)
        {
            Check.NotNull(dt, nameof(dt));
            Check.NotNull(stream, nameof(stream));

            using (ExcelPackage doc = new ExcelPackage())
            {
                doc.LoadFromDataTable(dt, dtOptionsAction);
                doc.SaveAs(stream);
            }
        }
        public static byte[] WriteToExcel(this DataTable dt, Action<DataTableOptions> dtOptionsAction = null)
        {
            Check.NotNull(dt, nameof(dt));

            using (var stream = new MemoryStream())
            using (ExcelPackage doc = new ExcelPackage())
            {
                doc.LoadFromDataTable(dt, dtOptionsAction);
                doc.SaveAs(stream);
                return stream.GetBuffer();
            }
        }
    }
}

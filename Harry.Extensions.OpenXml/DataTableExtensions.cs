using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Harry.Extensions.OpenXml
{
    public static class DataTableExtensions
    {
        public static void WriteToExcel(this DataTable dt, Stream stream, int pageSize = 50000)
        {
            Check.NotNull(dt, nameof(dt));
            Check.NotNull(stream, nameof(stream));

            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                doc.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                doc.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                doc.Write(dt, pageSize);
            }
        }

        public static void WriteToExcel(this DataTable dt, string path, int pageSize = 50000)
        {
            Check.NotNull(dt, nameof(dt));
            Check.NotNullOrEmpty(path, nameof(path));

            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                doc.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                doc.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                doc.Write(dt, pageSize);
            }
        }

        public static byte[] WriteToExcel(this DataTable dt, int pageSize = 50000)
        {
            Check.NotNull(dt, nameof(dt));

            using (MemoryStream ms = new MemoryStream())
            {
                dt.WriteToExcel(ms, pageSize);
                return ms.GetBuffer();
            }
        }



    }
}

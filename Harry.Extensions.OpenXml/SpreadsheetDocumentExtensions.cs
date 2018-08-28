using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Harry.Extensions.OpenXml
{
    public static class SpreadsheetDocumentExtensions
    {
        /// <summary>
        /// 向excel写入数据表
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="dt"></param>
        /// <param name="pageSize">每页写入多少条数据</param>
        public static void Write(this SpreadsheetDocument doc, DataTable dt, int pageSize = 50000)
        {
            Harry.Check.NotNull(doc, nameof(doc));
            Harry.Check.NotNull(dt, nameof(dt));

            if (pageSize <= 0)
                throw new Exception($"{nameof(pageSize)}必须为正整数");

            int pageIndex = 0;
            Worksheet currentSheet = null;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (i % pageSize == 0)
                {
                    //写入表头
                    currentSheet = createSheet(doc, dt);

                }

                //写入数据行
                //writeRow(currentSheet, dt.Rows[i]);
            }
        }

        /// <summary>
        /// 创建一张新表，并写入表头信息
        /// </summary>
        private static Worksheet createSheet(SpreadsheetDocument doc, DataTable dt)
        {
            //var workbookPart = doc.AddWorkbookPart();
            //doc.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            //doc.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

            var sheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            sheetPart.Worksheet = new Worksheet();

            //sheetPart.Worksheet.

            return sheetPart.Worksheet;
        }

        //private static IRow writeRow(ISheet sheet, DataRow row)
        //{
        //    return null;
        //}
    }
}

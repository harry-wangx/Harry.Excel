using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Harry.Extensions.EPPlus
{
    public static class ExcelPackageExtensions
    {
        /// <summary>
        /// 向excel写入数据表
        /// </summary>
        /// <param name="pageSize">每页写入多少条数据</param>
        public static void LoadFromDataTable(this ExcelPackage doc, DataTable dt, Action<DataTableOptions> dtOptionsAction)
        {
            Harry.Check.NotNull(doc, nameof(doc));
            Harry.Check.NotNull(dt, nameof(dt));

            var dtOptions = new DataTableOptions();
            dtOptionsAction?.Invoke(dtOptions);
            dtOptions.WorkbookAction?.Invoke(doc.Workbook);

            //doc.Workbook.Worksheets.Add($"{dt.TableName}")
            //    .Cells[1, 1].LoadFromDataTable(dt, false);
            //return;

            int rownum = 0;
            ExcelWorksheet currentSheet = null;
            int currentSheetIndex = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (i % dtOptions.PageSize == 0)
                {
                    //写入表头
                    rownum = 1;
                    currentSheet = createSheet(doc, dt, ref rownum, $"{dt.TableName}({(++currentSheetIndex).ToString()})", dtOptions);
                }

                //写入数据行
                writeRow(currentSheet, dt.Rows[i], rownum++, dtOptions);
            }
        }

        public static string GetValidSheetName(this ExcelPackage doc, string name)
        {
            Harry.Check.NotNull(doc, nameof(doc));
            Harry.Check.NotNullOrEmpty(name, nameof(name));

            var dicNames = doc.Workbook.Worksheets
                .Select(m => m.Name)
                .ToDictionary(m => m.ToUpperInvariant());

            if (!dicNames.ContainsKey(name.ToUpperInvariant()))
            {
                return name;
            }

            string sheetname;
            int index = 1;
            do
            {
                sheetname = $"{name}({index.ToString()})";
                if (!dicNames.ContainsKey(sheetname.ToUpperInvariant()))
                {
                    return sheetname;
                }

            } while (index++ < int.MaxValue);
            throw new Exception("恭喜您中了大奖，奖品为本人签名照一张。");
        }

        public static string GetValidSheetName(this ExcelPackage doc)
        {
            Harry.Check.NotNull(doc, nameof(doc));

            var dicNames = doc.Workbook.Worksheets
                .Select(m => m.Name)
                .ToDictionary(m => m.ToUpperInvariant());

            string sheetname;
            int index = 1;
            do
            {
                sheetname = "Sheet" + index.ToString();
                if (!dicNames.ContainsKey(sheetname.ToUpperInvariant()))
                {
                    return sheetname;
                }

            } while (index++ < int.MaxValue);
            throw new Exception("恭喜您中了大奖，奖品为本人签名照一张。");
        }

        /// <summary>
        /// 创建一张新表，并写入表头信息
        /// </summary>
        private static ExcelWorksheet createSheet(ExcelPackage doc, DataTable dt, ref int rownum, string sheetName, DataTableOptions dtOptions)
        {
            ExcelWorksheet worksheet = doc.Workbook.Worksheets.Add(doc.GetValidSheetName(sheetName));
            ////写入表名称
            //var titleCell = worksheet.Cells[rownum,  1];
            //titleCell.Value = dt.TableName;

            ////设置标题样式
            //dtOptions.TitleStyleAction?.Invoke(titleCell.Style);

            //rownum++;

            //写入字段信息
            foreach (DataColumn column in dt.Columns)
            {
                var header = worksheet.Cells[rownum, column.Ordinal + 1];
                header.Value = column.Caption;
                var excelColumn = worksheet.Column(column.Ordinal + 1);

                dtOptions.HeaderCellAction?.Invoke(column,excelColumn, header);
            }
            //var fieldNameCells = worksheet.Cells[rownum, 1, rownum, dt.Columns.Count];
            //dtOptions.HeaderStyleAction?.Invoke(fieldNameCells.Style);

            if (dtOptions.HeaderHeight != null)
            {
                //设置header行高
                worksheet.Row(rownum).Height = dtOptions.HeaderHeight.Value;
            }

            //冻结首行
            if (dtOptions.FrozenHeader)
            {
                worksheet.View.FreezePanes(rownum + 1, 1);
            }

            rownum++;

            return worksheet;
        }

        private static void writeRow(ExcelWorksheet sheet, DataRow dr, int rownum, DataTableOptions dtOptions)
        {
            if (dtOptions.DataHeight != null)
            {
                //设置行高
                sheet.Row(rownum).Height = dtOptions.DataHeight.Value;
            }
            foreach (DataColumn column in dr.Table.Columns)
            {
                var cell = sheet.Cells[rownum, column.Ordinal + 1];
                cell.Value = dr[column];

                ValidateCellStyle(cell, column.DataType);

                //设置单元格样式
                dtOptions.DataCellAction?.Invoke(column, cell);
            }
        }

        private static void ValidateCellStyle(ExcelRange cell, Type dataType)
        {
            switch (dataType.ToString())
            {
                case "System.DateTime":
                    if (string.IsNullOrEmpty(cell.Style.Numberformat.Format)
                        || "General".Equals(cell.Style.Numberformat.Format))
                    {
                        cell.Style.Numberformat.Format = "yyyy/m/d h:mm";
                    }
                    break;

                //case "System.String":

                //    break;

                //case "System.Boolean"://布尔型

                //    break;
                //case "System.SByte":
                //case "System.Byte":
                //case "System.Int16":
                //case "System.UInt16":
                //case "System.Int32":

                //    break;
                //case "System.UInt32":
                //case "System.Int64":
                //case "System.UInt64":
                //case "System.Decimal":
                //case "System.Double":

                //    break;
                default:
                    break;
            }
        }

    }
}

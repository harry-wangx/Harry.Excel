using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace Harry.Extensions.EPPlus
{
    public static class ExcelPackageExtensions
    {
        /// <summary>
        /// 获取可用Sheet名称
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="name"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 获取可用Sheet名称
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
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
        /// 向excel写入数据表
        /// </summary>
        public static ExcelPackage LoadFromDataTable(this ExcelPackage doc, DataTable dt, Action<DataTableOptions> dtOptionsAction)
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
            int currentSheetIndex = 0;
            ExcelWorksheet currentSheet = null;

            if (dt.Rows != null && dt.Rows.Count > 0)
            {
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
            else
            {
                //无数据，写个表头然后退出
                //写入表头
                rownum = 1;
                currentSheet = createSheet(doc, dt, ref rownum, $"{dt.TableName}({(++currentSheetIndex).ToString()})", dtOptions);
            }

            return doc;
        }

        /// <summary>
        /// 向excel写入数据集合
        /// </summary>
        public static ExcelPackage LoadFromEnumerableData<T>(this ExcelPackage doc, IEnumerable<T> data, Action<EnumerableDataOptions<T>> optionsAction)
        {
            Harry.Check.NotNull(doc, nameof(doc));
            Harry.Check.NotNull(data, nameof(data));

            var options = new EnumerableDataOptions<T>();
            optionsAction?.Invoke(options);

            if (string.IsNullOrEmpty(options.TableName))
                options.TableName = "Sheet";

            options.WorkbookAction?.Invoke(doc.Workbook);


            //doc.Workbook.Worksheets.Add($"{dt.TableName}")
            //    .Cells[1, 1].LoadFromDataTable(dt, false);
            //return;

            int rownum = 0;
            int currentSheetIndex = 0;
            ExcelWorksheet currentSheet = null;

            rownum = 1;
            currentSheet = createSheetForEnumerableData(doc, ref rownum, $"{options.TableName}({(++currentSheetIndex).ToString()})", options);

            int i = 0;
            foreach (var item in data)
            {
                //写入数据行
                if (options.DataHeight != null)
                {
                    //设置行高
                    currentSheet.Row(rownum).Height = options.DataHeight.Value;
                }

                options.DataRowAction?.Invoke(currentSheet, rownum, item);

                rownum++;
                i++;

                if (i % options.PageSize == 0)
                {
                    //写入表头
                    rownum = 1;
                    currentSheet = createSheetForEnumerableData(doc, ref rownum, $"{options.TableName}({(++currentSheetIndex).ToString()})", options);
                }
            }

            return doc;
        }

        public static byte[] GetBuffer(this ExcelPackage doc)
        {
            Harry.Check.NotNull(doc, nameof(doc));

            using (var ms = new MemoryStream())
            {
                doc.SaveAs(ms);

                ms.Position = 0;
                byte[] buffer = new byte[ms.Length];
                ms.Read(buffer, 0, (int)ms.Length);

                return buffer;
            }
        }

        //创建一张新表，并写入表头信息
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

                dtOptions.HeaderCellAction?.Invoke(column, excelColumn, header);
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
                var value = dr[column];
                if (value != null && value != DBNull.Value)
                {
                    cell.Value = dr[column];
                }

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

        private static ExcelWorksheet createSheetForEnumerableData<T>(ExcelPackage doc, ref int rownum, string sheetName, EnumerableDataOptions<T> options)
        {
            ExcelWorksheet worksheet = doc.Workbook.Worksheets.Add(doc.GetValidSheetName(sheetName));
            
            ////写入表名称
            //var titleCell = worksheet.Cells[rownum,  1];
            //titleCell.Value = dt.TableName;

            ////设置标题样式
            //dtOptions.TitleStyleAction?.Invoke(titleCell.Style);

            //rownum++;

            //写入表头信息
            //if (options.HeaderRowActions != null && options.HeaderRowActions.Count > 0)
            //{
            //    for (int i = 1; i <= options.HeaderRowActions.Count; i++)
            //    {
            //        var header = worksheet.Cells[rownum, i];
            //        var excelColumn = worksheet.Column(i);
            //        options.HeaderRowActions[i].Invoke(excelColumn, header);
            //    }
            //}

            options.HeaderRowAction?.Invoke(worksheet, rownum);

            if (options.HeaderHeight != null)
            {
                //设置header行高
                worksheet.Row(rownum).Height = options.HeaderHeight.Value;
            }

            //冻结首行
            if (options.FrozenHeader)
            {
                worksheet.View.FreezePanes(rownum + 1, 1);
            }

            rownum++;

            return worksheet;
        }

    }
}

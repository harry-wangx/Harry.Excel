using Harry.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Harry.Extensions.EPPlus;
using System.Drawing;
using OfficeOpenXml.Style;

namespace Harry.Extensions.EPPlus.Test
{
    public class DataTableExtensionsTest
    {
        [Test]
        public void DatableToExcel()
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            path = System.IO.Path.Combine(path, "test_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_DataTable.xlsx");
            Console.WriteLine(path);

            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                Helper.GetDataTable().WriteToExcel(fs, SetOptions);
            }
        }

        [Test]
        public void DatableToExcelWithMem()
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            path = System.IO.Path.Combine(path, "test_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_DataTable.xlsx");
            Console.WriteLine(path);

            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                var data = Helper.GetDataTable().WriteToExcel(SetOptions);
                fs.Write(data, 0, data.Length);
            }
        }

        private void SetOptions(DataTableOptions options)
        {
            options.HeaderHeight = 33;
            options.DataHeight = 31;

            //设置表头样式
            options.HeaderCellAction = (dataColumn, excelColumn, cell) =>
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(164, 50, 83));
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.Font.Size = 11;
                cell.Style.Font.Name = "微软雅黑";
                cell.Style.Font.Bold = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.WrapText = true;

                switch (dataColumn.ColumnName)
                {
                    case "Id":
                        excelColumn.Width = 8;
                        break;
                    case "Name":
                        excelColumn.Width = 12;
                        break;
                    case "Birthday":
                        excelColumn.Width = 15;
                        break;
                }
            };

            //设置数据单元格样式
            options.DataCellAction = (dataColumn, cell) =>
            {
                cell.Style.Font.Size = 11;
                cell.Style.Font.Name = "微软雅黑";
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.WrapText = true;
            };

            //dtStyle.ExcelWorksheetAction = sheet => { };
        }
    }
}

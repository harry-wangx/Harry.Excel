using Harry.Excel;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace Harry.Extensions.EPPlus.Test
{
    public class ExcelPackageExtensionsTest
    {
        [Test]
        public void LoadFromEnumerableData()
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            path = System.IO.Path.Combine(path, "test_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_Data.xlsx");
            Console.WriteLine(path);

            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                var data = new ExcelPackage().LoadFromEnumerableData<Person>(Helper.GetPersons(), options =>
                {
                    //options.PageSize = 1;
                    options.HeaderRowAction = (sheet, rownum) =>
                    {
                        sheet.Cells[rownum, 1].Value = "编号";
                        sheet.Cells[rownum, 2].Value = "姓名";
                        sheet.Cells[rownum, 3].Value = "生日";
                        var style = sheet.Cells[rownum, 1, rownum, 3].Style;
                        style.Fill.PatternType = ExcelFillStyle.Solid;
                        style.Fill.BackgroundColor.SetColor(Color.FromArgb(164, 50, 83));
                        style.Font.Color.SetColor(Color.White);
                        style.Font.Size = 11;
                        style.Font.Name = "微软雅黑";
                        style.Font.Bold = true;
                        style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        style.WrapText = true;
                    };
                    options.DataRowAction = (sheet, rownum, item) =>
                    {
                        sheet.Cells[rownum, 1].Value = item.Id;
                        sheet.Cells[rownum, 2].Value = item.Name;
                        sheet.Cells[rownum, 3].Value = item.Birthday;
                        sheet.Cells[rownum, 3].Style.Numberformat.Format = "yyyy/m/d";
                        var style = sheet.Cells[rownum, 1, rownum, 3].Style;
                        style.Font.Size = 11;
                        style.Font.Name = "微软雅黑";
                        style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        style.WrapText = true;
                    };
                })
                .GetBuffer();
                fs.Write(data, 0, data.Length);
            }
        }
    }
}

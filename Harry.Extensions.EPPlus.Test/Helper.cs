using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Harry.Excel
{
    class Helper
    {
        public static DataTable GetDataTable()
        {
            DataTable dt = new DataTable("数据测试表");
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Birthday", typeof(DateTime));

            dt.Rows.Add(1, "张三", new DateTime(1988, 5, 7));
            dt.Rows.Add(2, "李四", new DateTime(1992, 6, 8));

            return dt;
        }

        public static List<Person> GetPersons()
        {
            List<Person> results = new List<Person>();

            results.Add(new Person() { Id = 1, Name = "张三", Birthday = new DateTime(1988, 5, 7) });
            results.Add(new Person() { Id = 2, Name = "李四", Birthday = new DateTime(1992, 6, 8) });

            return results;
        }
    }
}

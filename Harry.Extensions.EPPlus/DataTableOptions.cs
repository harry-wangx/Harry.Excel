﻿using OfficeOpenXml;
using System;
using System.Data;

namespace Harry.Extensions.EPPlus
{
    public class DataTableOptions
    {
        /// <summary>
        /// 
        /// </summary>
        public Action<ExcelWorkbook> WorkbookAction { get; set; }

        private int pageSize = 50000;
        /// <summary>
        /// 每页数据行数
        /// </summary>
        public int PageSize
        {
            get => pageSize;
            set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(PageSize), "PageSize必须大于0");
                }
                pageSize = value;
            }
        }


        /// <summary>
        /// 表头单元格样式
        /// </summary>
        public Action<DataColumn, ExcelColumn, ExcelRange> HeaderCellAction { get; set; }

        /// <summary>
        /// 数据单元格样式
        /// </summary>
        public Action<DataColumn, ExcelRange> DataCellAction { get; set; }

        /// <summary>
        /// 表头行高
        /// </summary>
        public double? HeaderHeight { get; set; }

        /// <summary>
        /// 是否冻结表头
        /// </summary>
        public bool FrozenHeader { get; set; } = true;

        /// <summary>
        /// 数据行高
        /// </summary>
        public double? DataHeight { get; set; }

    }
}

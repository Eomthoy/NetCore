using System;

namespace Core.Common.Helper
{
    /// <summary>
    /// Excel导入
    /// </summary>
    public class ExcelColumAttribute : Attribute
    {
        //public ExcelColumAttribute();

        public string ColumName { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;
using System.Reflection;
using NPOI.XSSF.UserModel;

namespace Core.Common.Helper
{
    /// <summary>
    /// Excel数据获取
    /// </summary>
    public class ExcelHelper
    {
        #region Excel数据获取

        //数据所在的行数，初始值为第firstRowColumn+1行
        private static int row = 0;

        /// <summary>
        /// 读取excel数据，返回List<T>数据集
        /// </summary>
        /// <typeparam name="T">entity</typeparam>
        /// <param name="fileName">文件路径</param>
        /// <param name="sheetName">表名，默认为第一张表</param>
        /// <param name="firstRowColumn">列名出现的行数，默认为第一行</param>
        /// <returns></returns>
        public static List<T> GetAll<T>(string fileName, string sheetName = null, int firstRowColumn = 1) where T : new()
        {
            row = firstRowColumn + 1;
            List<T> result = new List<T>();
            //获取原始数据
            DataTable dt = ReadExcelToDataTable(fileName, sheetName, firstRowColumn - 1);
            //获取所有匹配的列
            List<string> Colums = ExistList<T>(dt, new T());
            //数据循环赋值
            if (Colums.Count > 0)
            {
                foreach (DataRow item in dt.Rows)
                {
                    T model = new T();
                    model = AddModel<T>(item, model, Colums);
                    result.Add(model);
                    row++;
                }
            }
            return result;
        }
        /// <summary>
        /// 将excel文件内容读取到DataTable数据表中
        /// </summary>
        /// <param name="fileName">文件完整路径名</param>
        /// <param name="sheetName">指定读取excel工作薄sheet的名称</param>
        /// <param name="firstRowColumn">列名出现的行数</param>
        /// <returns>DataTable数据表</returns>
        public static DataTable ReadExcelToDataTable(string fileName, string sheetName, int firstRowColumn)
        {
            //定义要返回的datatable对象
            DataTable data = new DataTable();
            //excel工作表
            ISheet sheet = null;
            try
            {
                if (!File.Exists(fileName))
                {
                    return null;
                }
                //根据指定路径读取文件
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    //根据文件流创建excel数据结构
                    IWorkbook workbook = null;
                    if (Path.GetExtension(fileName) == ".xls")
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else if (Path.GetExtension(fileName) == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    //如果有指定工作表名称
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        sheet = workbook.GetSheet(sheetName);
                        //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                        if (sheet == null)
                        {
                            sheet = workbook.GetSheetAt(0);
                        }
                    }
                    else
                    {
                        //如果没有指定的sheetName，则尝试获取第一个sheet
                        sheet = workbook.GetSheetAt(0);
                    }
                    if (sheet != null)
                    {
                        #region 获取列名

                        IRow firstRow = sheet.GetRow(firstRowColumn);
                        //一行最后一个cell的编号 即总的列数
                        int cellCount = firstRow.LastCellNum;
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        #endregion

                        #region 获取数据

                        //数据开始行(排除列名行)
                        int startRow = firstRowColumn + 1;
                        //最后一列的标号
                        int rowCount = sheet.LastRowNum;
                        //循环数据
                        for (int i = startRow; i <= rowCount; ++i)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null) continue; //没有数据的行默认是null

                            DataRow dataRow = data.NewRow();
                            for (int j = row.FirstCellNum; j < cellCount; ++j)
                            {
                                if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                    dataRow[j] = row.GetCell(j).ToString();
                            }
                            data.Rows.Add(dataRow);
                        }
                        #endregion

                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// 返回一个传入的Model
        /// </summary>
        /// <typeparam name="Model"></typeparam>
        /// <param name="dt">行数据源</param>
        /// <param name="model">传入的Model</param>
        /// <param name="Exist">当前Model中的值</param>
        /// <returns>将行数据转化成Model返回</returns>
        public static Model AddModel<Model>(DataRow dt, Model model, List<string> Exist)
        {
            try
            {
                Type ModelType = typeof(Model);
                PropertyInfo[] ModelTypeProp = ModelType.GetProperties();
                //循环给我们要附加的列进行赋值
                if (Exist.Count > 0)
                {
                    foreach (PropertyInfo propertyInfo in ModelTypeProp)
                    {
                        ExcelColumAttribute attribute = propertyInfo.GetCustomAttributes(false).Where(o => o.GetType() == typeof(ExcelColumAttribute)).FirstOrDefault() as ExcelColumAttribute;
                        if (attribute != null)
                        {
                            if (Exist.Contains(attribute.ColumName.Trim().ToLower()) && !string.IsNullOrEmpty(dt[attribute.ColumName].ToString()))
                            {
                                try
                                {
                                    string tp = propertyInfo.PropertyType.Name;//获得对象的属性类型
                                                                               //处理时间格式
                                    if (tp.Contains("Nullable"))
                                    {
                                        Type[] TTModel = propertyInfo.PropertyType.GetGenericArguments();
                                        string TypeName = TTModel[0].FullName.ToLower();
                                        if (TypeName.Contains("int32"))
                                        {
                                            int value = Convert.ToInt32(dt[attribute.ColumName]);
                                            propertyInfo.SetValue(model, value, null);//赋值的对象
                                        }
                                        else if (TypeName.Contains("datetime"))
                                        {
                                            DateTime value = Convert.ToDateTime(dt[attribute.ColumName]);
                                            propertyInfo.SetValue(model, value, null);//赋值的对象
                                        }
                                        else if (TypeName.Contains("double"))
                                        {
                                            double value = Convert.ToDouble(dt[attribute.ColumName]);
                                            propertyInfo.SetValue(model, value, null);//赋值的对象
                                        }
                                        else
                                        {
                                            string value = dt[attribute.ColumName].ToString();
                                            if (!string.IsNullOrEmpty(value))
                                            {
                                                propertyInfo.SetValue(model, value, null);//赋值的对象
                                            }
                                        }
                                    }
                                    //处理bool
                                    else if (tp.Contains("Boolean"))
                                    {
                                        bool value = dt[attribute.ColumName].ToString() == "真" ? true : false;
                                        propertyInfo.SetValue(model, value, null);//赋值的对象
                                    }
                                    else if (tp.Trim().ToLower().Contains("int"))
                                    {
                                        int value = Convert.ToInt32(dt[attribute.ColumName].ToString());
                                        propertyInfo.SetValue(model, value, null);//赋值的对象
                                    }
                                    else if (tp.Trim().ToLower().Contains("datetime"))
                                    {
                                        DateTime value = Convert.ToDateTime(dt[attribute.ColumName]);
                                        propertyInfo.SetValue(model, value, null);//赋值的对象
                                    }
                                    else if (tp.Trim().ToLower().Contains("double"))
                                    {
                                        double value = Convert.ToDouble(dt[attribute.ColumName]);
                                        propertyInfo.SetValue(model, value, null);//赋值的对象
                                    }
                                    //其他格式
                                    else
                                    {
                                        string value = dt[attribute.ColumName].ToString();
                                        propertyInfo.SetValue(model, value, null);//赋值的对象
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception($"第【{row}】行，【{attribute.ColumName}】数据格式错误！");
                                }
                            }
                        }
                    }
                }
                return model;
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// 传入的DataTable和model的列比对 返回model中存在的列的集合
        /// </summary>
        /// <typeparam name="Model"></typeparam>
        /// <param name="dt">行数据源</param>
        /// <param name="model">传入的Model</param>
        /// <returns> 返回model中存在的列的集合</returns>
        public static List<string> ExistList<Model>(DataTable dt, Model model)
        {
            try
            {
                DataColumnCollection columns = dt.Columns;

                Type ModelType = typeof(Model);
                PropertyInfo[] ModelTypeProp = ModelType.GetProperties();
                //取出当前传入的Model的字段
                List<string> NameList = new List<string>();
                foreach (PropertyInfo item in ModelTypeProp)
                {
                    ExcelColumAttribute attribute = item.GetCustomAttributes(false).Where(o => o.GetType() == typeof(ExcelColumAttribute)).FirstOrDefault() as ExcelColumAttribute;
                    if (attribute != null)
                        NameList.Add(attribute.ColumName);
                }
                //取出数据库有的字段(数据库的字段必须和Model的字段一直 大小写都要一致)
                List<string> Exist = new List<string>();
                foreach (DataColumn item in columns)
                {
                    string ColumName = item.ToString();
                    if (NameList.Contains(ColumName))
                    {
                        Exist.Add(ColumName);
                    }
                }
                return Exist;
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw new Exception(ex.Message);
            }
        }
        #endregion
    }
}



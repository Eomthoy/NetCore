using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;

//Excel
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace Core.Common.Helper
{
    /// <summary>
    /// 通用数据访问类
    /// </summary>
    public class SqlHelper
    {

        #region 数据源：SqlServer数据库

        //数据库连接字符串
        public static string connString = "";

        /// <summary>
        /// 返回单个值
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static object ExecuteScalar(string sql, CommandType commandType, params SqlParameter[] parameters)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandType = commandType;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        conn.Open();
                        return cmd.ExecuteScalar();
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw new Exception(ex.Message);
            }
        }

        #region 增删改
        /// <summary>
        /// 执行增删改，返回受影响行数
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static int Execute(string sql, CommandType commandType, params SqlParameter[] parameters)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandType = commandType;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        conn.Open();
                        return cmd.ExecuteNonQuery();
                    }
                }
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

        #region SqlDataReader
        /// <summary>
        /// 执行查询，返回SqlDataReader对象
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static SqlDataReader ExecuteReader(string sql, CommandType commandType, params SqlParameter[] parameters)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandType = commandType;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        conn.Open();
                        return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    }
                }
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
        /// 执行查询，获取List<T>集合
        /// </summary>
        /// <typeparam name="T">entity</typeparam>
        /// <param name="reader">data reader</param>
        /// <returns>entity</returns>
        public static List<T> GetAllFromReader<T>(string sql, CommandType commandType, params SqlParameter[] parameters) where T : new()
        {
            SqlDataReader reader = ExecuteReader(sql, commandType, parameters);
            List<T> list = new List<T>();
            using (reader)
            {
                while (reader.Read())
                {
                    T model = ReaderToEntity<T>(reader);
                    list.Add(model);
                }
            }
            return list;
        }
        /// <summary>
        /// 执行查询，获取T单个对象
        /// </summary>
        /// <typeparam name="T">entity</typeparam>
        /// <param name="reader">data reader</param>
        /// <returns>entity</returns>
        public static T GetFromReader<T>(string sql, CommandType commandType, params SqlParameter[] parameters) where T : new()
        {
            SqlDataReader reader = ExecuteReader(sql, commandType, parameters);
            T model = new T();
            using (reader)
            {
                while (reader.Read())
                {
                    model = ReaderToEntity<T>(reader);
                }
            }
            return model;
        }
        /// <summary>
        /// Reader To Entity
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        private static T ReaderToEntity<T>(SqlDataReader data) where T : new()
        {
            T val = new T();
            foreach (PropertyInfo propertyInfo in typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                //字段属性（字段名、字段类型）
                string propertyName = propertyInfo.Name;
                string propertyType = propertyInfo.PropertyType.ToString().Trim().ToLower();
                //赋值
                var obj = new object();
                obj = data[propertyName];
                if (obj != DBNull.Value && obj != null)
                {
                    while (propertyInfo.GetSetMethod() != (MethodInfo)null)
                    {
                        propertyInfo.SetValue(val, obj, null);
                    }
                }
            }
            return val;
        }
        #endregion

        #region DataTable

        /// <summary>
        /// 执行查询，返回DataTable对象
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static DataTable ExecuteDataTable(string sql, CommandType commandType, params SqlParameter[] parameters)
        {
            try
            {
                using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connString))
                {
                    DataTable dt = new DataTable();
                    adapter.SelectCommand.CommandType = commandType;
                    if (parameters != null)
                    {
                        adapter.SelectCommand.Parameters.AddRange(parameters);
                    }
                    adapter.Fill(dt);
                    return dt;
                }
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
        /// 执行查询，获取数据集List<T>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static List<T> GetAll<T>(string sql, CommandType commandType, params SqlParameter[] parameters) where T : new()
        {
            try
            {
                DataTable dt = ExecuteDataTable(sql, commandType, parameters);
                List<T> list = new List<T>();
                foreach (DataRow row in dt.Rows)
                {
                    T model = DataTableToEntity<T>(row, dt.Columns);
                    list.Add(model);
                }

                return list;
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
        /// 执行查询，获取单个对象T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static T Get<T>(string sql, CommandType commandType, params SqlParameter[] parameters) where T : new()
        {
            try
            {
                DataTable dt = ExecuteDataTable(sql, commandType, parameters);
                T model = new T();
                foreach (DataRow row in dt.Rows)
                {
                    model = DataTableToEntity<T>(row, dt.Columns);
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
        /// DataTable To Entity
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        private static T DataTableToEntity<T>(DataRow data, DataColumnCollection columns) where T : new()
        {
            T model = new T();
            foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
            {
                //字段属性
                string propName = propertyInfo.Name;
                string propType = propertyInfo.PropertyType.Name.Trim().ToLower();
                //赋值
                if (columns.Contains(propName))
                {
                    object val = data[propName];
                    if (val != DBNull.Value && val != null)
                    {
                        propertyInfo.SetValue(model, val, null);
                    }
                }
            }

            return model;
        }
        #endregion

        #endregion

        #region 数据源：excel

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
        public static List<T> GetAllFromExcel<T>(string fileName, string sheetName = null, int firstRowColumn = 1) where T : new()
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


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using System.Data;
using NPOI.SS.Util;
using NPOI.SS.UserModel;

namespace Core.Common.Helper
{
    /// <summary>
    /// 通用数据访问类
    /// </summary>
    public class ExcelHelper
    {
        static HSSFWorkbook hssfworkbook;

        public static void WirteExcel()
        {
            InitializeWorkbook();

            //写标题文本
            HSSFSheet sheet1 = hssfworkbook.CreateSheet("校外人员劳务领取表") as HSSFSheet;
            HSSFCell cellTitle = sheet1.CreateRow(0).CreateCell(0) as HSSFCell;
            cellTitle.SetCellValue("校外人员劳务领取表");

            //设置标题行样式
            HSSFCellStyle style = (HSSFCellStyle)hssfworkbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            HSSFFont font = hssfworkbook.CreateFont() as HSSFFont;
            font.FontHeight = 20 * 20;
            style.SetFont(font);

            cellTitle.CellStyle = style;

            //合并标题行
            sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 9));

            DataTable dt = GetData();
            IRow row;
            ICell cell;
            //HSSFCellStyle celStyle = getCellStyle();

            HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFClientAnchor anchor;
            HSSFSimpleShape line;
            int rowIndex;
            //表头数据
            rowIndex = 1;
            row = sheet1.CreateRow(rowIndex);
            cell = row.CreateCell(0);
            cell.SetCellValue("编号");
            cell = row.CreateCell(1);
            cell.SetCellValue("姓名");
            cell = row.CreateCell(2);
            cell.SetCellValue("身份证号");
            cell = row.CreateCell(3);
            cell.SetCellValue("工作单位");
            cell = row.CreateCell(4);
            cell.SetCellValue("银行卡号");
            cell = row.CreateCell(5);
            cell.SetCellValue("工作时间");
            cell = row.CreateCell(6);
            cell.SetCellValue("工作内容");
            cell = row.CreateCell(7);
            cell.SetCellValue("发放标准");
            cell = row.CreateCell(8);
            cell.SetCellValue("发放金额");
            cell = row.CreateCell(9);
            cell.SetCellValue("领款人签字");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //表头数据
                rowIndex = 3 * (i + 1);
                row = sheet1.CreateRow(rowIndex);

                cell = row.CreateCell(0);
                cell.SetCellValue("姓名");
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(1);
                cell.SetCellValue("基本工资");
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(2);
                cell.SetCellValue("住房公积金");
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(3);
                cell.SetCellValue("绩效奖金");
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(4);
                cell.SetCellValue("社保扣款");
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(5);
                cell.SetCellValue("代扣个税");
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(6);
                cell.SetCellValue("实发工资");
                //cell.CellStyle = celStyle;


                DataRow dr = dt.Rows[i];
                //设置值和计算公式
                row = sheet1.CreateRow(rowIndex + 1);
                cell = row.CreateCell(0);
                cell.SetCellValue(dr["FName"].ToString());
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(1);
                cell.SetCellValue((double)dr["FBasicSalary"]);
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(2);
                cell.SetCellValue((double)dr["FAccumulationFund"]);
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(3);
                cell.SetCellValue((double)dr["FBonus"]);
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(4);
                cell.SetCellFormula(String.Format("$B{0}*0.08", rowIndex + 2));
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(5);
                cell.SetCellFormula(String.Format("SUM($B{0}:$D{0})*0.1", rowIndex + 2));
                //cell.CellStyle = celStyle;

                cell = row.CreateCell(6);
                cell.SetCellFormula(String.Format("SUM($B{0}:$D{0})-SUM($E{0}:$F{0})", rowIndex + 2));
                //cell.CellStyle = celStyle;


                //绘制分隔线
                //sheet1.AddMergedRegion(new Region(rowIndex + 2, 0, rowIndex + 2, 6));
                //anchor = new HSSFClientAnchor(0, 125, 1023, 125, 0, rowIndex + 2, 6, rowIndex + 2);
                //line = patriarch.CreateSimpleShape(anchor);
                //line.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
                //line.LineStyle = HSSFShape.LINESTYLE_DASHGEL;
            }
            WriteToFile();
        }
        static DataTable GetData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("FName", typeof(System.String));
            dt.Columns.Add("FBasicSalary", typeof(System.Double));
            dt.Columns.Add("FAccumulationFund", typeof(System.Double));
            dt.Columns.Add("FBonus", typeof(System.Double));

            dt.Rows.Add("令狐冲", 6000, 1000, 2000);
            dt.Rows.Add("任盈盈", 7000, 1000, 2500);
            dt.Rows.Add("林平之", 5000, 1000, 1500);
            dt.Rows.Add("岳灵珊", 4000, 1000, 900);
            dt.Rows.Add("任我行", 4000, 1000, 800);
            dt.Rows.Add("风清扬", 9000, 5000, 3000);

            return dt;
        }


        //static HSSFCellStyle getCellStyle()
        //{
        //    HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle();
        //    cellStyle.BorderBottom = HSSFCellStyle.BORDER_THIN;
        //    cellStyle.BorderLeft = HSSFCellStyle.BORDER_THIN;
        //    cellStyle.BorderRight = HSSFCellStyle.BORDER_THIN;
        //    cellStyle.BorderTop = HSSFCellStyle.BORDER_THIN;
        //    return cellStyle;
        //}

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(Directory.GetCurrentDirectory() + @"/test.xls", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}



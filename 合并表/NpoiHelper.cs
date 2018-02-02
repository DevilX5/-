using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 合并表
{
    public class NpoiHelper
    {
        public static List<string> GetAllSheet(string filePath)
        {
            var lst = new List<string>();
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                var workbook = new HSSFWorkbook(fs);
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    lst.Add(workbook.GetSheetAt(i).SheetName);
                }
            }
            return lst;
        }
        public static void FillDataSet(DataSet ds, FileAndSheets fileAndSheets)
        {
            using (FileStream fs = new FileStream(fileAndSheets.FileName, FileMode.Open, FileAccess.ReadWrite))
            {
                var workbook = new HSSFWorkbook(fs);
                foreach (var name in fileAndSheets.SheetNames)
                {
                    var dt = FillDataTable(workbook, name);
                    ds.Tables.Add(dt);
                }
            }
        }
        public static DataTable GetSingleSheetData(FileAndSheets fileAndSheets)
        {
            using (FileStream fs = new FileStream(fileAndSheets.FileName, FileMode.Open, FileAccess.ReadWrite))
            {
                var workbook = new HSSFWorkbook(fs);
                return FillDataTable(workbook, fileAndSheets.CurrentSheetName);
            }
        }
        public static void ReadExcelToDataSet(DataSet ds, FileAndSheets fileAndSheets)
        {
            DataRow dr;
            ICell objCellValue;
            string cellValue;
            using (FileStream fs = new FileStream(fileAndSheets.FileName, FileMode.Open, FileAccess.ReadWrite))
            {
                var workbook = new HSSFWorkbook(fs);

                foreach (var name in fileAndSheets.SheetNames)
                {
                    var sheet = workbook.GetSheet(name);
                    if (sheet == null) continue;
                    var dt = new DataTable();
                    var rowCount = sheet.LastRowNum;
                    var firstRow = sheet.GetRow(0);
                    var columnCount = firstRow.LastCellNum;
                    for (int i = 0; i < columnCount; i++)
                    {
                        objCellValue = firstRow.GetCell(i);
                        cellValue = objCellValue == null ? "" : objCellValue.StringCellValue;
                        dt.Columns.Add(cellValue);
                    }
                    for (int k = 1; k <= rowCount; k++)
                    {
                        var currentRow = sheet.GetRow(k);
                        dr = dt.NewRow();
                        for (int j = 1; j <= columnCount; j++)
                        {
                            objCellValue = currentRow.GetCell(j - 1);
                            if (objCellValue == null)
                            {
                                dr[j - 1] = "";
                            }
                            else
                            {
                                switch (objCellValue.CellType)
                                {
                                    case CellType.Numeric:
                                        short format = objCellValue.CellStyle.DataFormat;
                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                            dr[j - 1] = objCellValue.DateCellValue;
                                        else
                                            dr[j - 1] = objCellValue.NumericCellValue;
                                        break;
                                    case CellType.String:
                                        dr[j - 1] = objCellValue.StringCellValue;
                                        break;
                                    default:
                                        dr[j - 1] = "";
                                        break;
                                }
                            }
                            //cellValue = objCellValue == null ? "" : objCellValue.StringCellValue;
                            //dr[j - 1] = cellValue;
                        }
                        dt.Rows.Add(dr);
                    }
                    ds.Tables.Add(dt);
                }
            }
        }
        static DataTable FillDataTable(HSSFWorkbook workbook, string sheetname)
        {
            DataRow dr;
            ICell objCellValue;
            string cellValue;
            var sheet = workbook.GetSheet(sheetname);
            var dt = new DataTable();
            if (sheet == null) return dt;
            var rowCount = sheet.LastRowNum;
            var firstRow = sheet.GetRow(0);
            var columnCount = firstRow.LastCellNum;
            for (int i = 0; i < columnCount; i++)
            {
                objCellValue = firstRow.GetCell(i);
                cellValue = objCellValue == null ? "" : objCellValue.StringCellValue;
                dt.Columns.Add(cellValue);
            }
            for (int k = 1; k <= rowCount; k++)
            {
                var currentRow = sheet.GetRow(k);
                dr = dt.NewRow();
                for (int j = 1; j <= columnCount; j++)
                {
                    objCellValue = currentRow.GetCell(j - 1);
                    if (objCellValue == null)
                    {
                        dr[j - 1] = "";
                    }
                    else
                    {
                        switch (objCellValue.CellType)
                        {
                            case CellType.Formula:
                                HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(objCellValue.Sheet.Workbook);
                                e.EvaluateInCell(objCellValue);
                                dr[j-1] = objCellValue.ToString();
                                break;
                            case CellType.Numeric:
                                //short format = objCellValue.CellStyle.DataFormat;
                                //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                //if (format == 14 || format == 31 || format == 57 || format == 58)
                                if(DateUtil.IsCellDateFormatted(objCellValue))
                                    dr[j - 1] = objCellValue.DateCellValue;
                                else
                                    dr[j - 1] = objCellValue.NumericCellValue;
                                break;
                            case CellType.String:
                                dr[j - 1] = objCellValue.StringCellValue;
                                break;
                            default:
                                dr[j - 1] = objCellValue.ToString();
                                break;
                        }
                    }
                    //cellValue = objCellValue == null ? "" : objCellValue.StringCellValue;
                    //dr[j - 1] = cellValue;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}

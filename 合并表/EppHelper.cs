using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace 合并表
{
    public class EppHelper
    {



        public static bool ExportByDt(string path, DataTable Dt, string SheetName = "Sheet1", bool WithTitle = true,string startcell="A1")
        {
            string startstr = string.Join("", new Regex("[a-zA-z]").Matches(startcell).Cast<Match>().Select(n => n.Value));
            string endstr = string.Join("", new Regex("[0-9]").Matches(startcell).Cast<Match>().Select(n => n.Value));
            var col = (startstr.Length-1)*26+ startstr.ToUpper().Last() - 64;
            var row = Int32.Parse(endstr);
            var result = false;
            FileInfo newFile = new FileInfo(path);
            using (var pkg = new ExcelPackage(newFile))
            {
                var ws = pkg.Workbook.Worksheets.FirstOrDefault(n => n.Name.Equals(SheetName));
                if (ws == null)
                {
                    ws = pkg.Workbook.Worksheets.Add(SheetName);
                }
                ws.Cells[row,col].LoadFromDataTable(Dt, WithTitle);
                var toRow = Dt.Rows.Count + (WithTitle ? 1 : 0);
                var toCol=Dt.Columns.Count;

                ws.Cells[row,col,toRow,toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[row,col,toRow,toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[row,col,toRow,toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[row,col,toRow,toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row,col,toRow,toCol].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                ws.Cells[row,col,toRow,toCol].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                ws.Cells[row,col,toRow,toCol].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                ws.Cells[row, col, Dt.Rows.Count, Dt.Columns.Count].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                pkg.Save();
                result = true;
            }
            return result;
        }
        public static List<string> GetAllSheet(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                using (var package = new ExcelPackage())
                {
                    package.Load(fs);
                    return package.Workbook.Worksheets.Select(n => n.Name).ToList();
                }
            }
        }
        public static void ReadExcelToDataSet(DataSet ds,FileAndSheets fileAndSheets)
        {
            DataRow dr;
            object objCellValue;
            string cellValue;
            using (FileStream fs = new FileStream(fileAndSheets.FileName, FileMode.Open, FileAccess.ReadWrite))
            using (var package = new ExcelPackage())
            {
                package.Load(fs);
                foreach (var sheet in package.Workbook.Worksheets.Where(n => fileAndSheets.SheetNames.Contains(n.Name)))
                {
                    if (sheet.Dimension == null) continue;
                    var columnCount = sheet.Dimension.End.Column;
                    var rowCount = sheet.Dimension.End.Row;
                    if (rowCount > 0)
                    {
                        var dt = new DataTable();
                        for (int j = 0; j < columnCount; j++)//设置DataTable列名  
                        {
                            objCellValue = sheet.Cells[1, j + 1].Value;
                            cellValue = objCellValue == null ? "" : objCellValue.ToString();
                            dt.Columns.Add(cellValue, typeof(string));
                        }
                        for (int i = 2; i <= rowCount; i++)
                        {
                            dr = dt.NewRow();
                            for (int j = 1; j <= columnCount; j++)
                            {
                                objCellValue = sheet.Cells[i, j].Value;
                                cellValue = objCellValue == null ? "" : objCellValue.ToString();
                                dr[j - 1] = cellValue;
                            }
                            dt.Rows.Add(dr);
                        }
                        ds.Tables.Add(dt);
                    }
                }
            }
        }
        public static DataTable GetSingleSheetData(FileAndSheets fileAndSheets)
        {
            DataRow dr;
            object objCellValue;
            string cellValue;
            using (FileStream fs = new FileStream(fileAndSheets.FileName, FileMode.Open, FileAccess.ReadWrite))
            using (var package = new ExcelPackage())
            {
                package.Load(fs);
                var sheet = package.Workbook.Worksheets[fileAndSheets.CurrentSheetName];
                var dt = new DataTable();
                if (sheet.Dimension == null) return dt;
                var columnCount = sheet.Dimension.End.Column;
                var rowCount = sheet.Dimension.End.Row;
                if (rowCount > 0)
                {
                    for (int j = 0; j < columnCount; j++)//设置DataTable列名  
                    {
                        objCellValue = sheet.Cells[1, j + 1].Value;
                        cellValue = objCellValue == null ? "" : objCellValue.ToString();
                        dt.Columns.Add(cellValue, typeof(string));
                    }
                    for (int i = 2; i <= rowCount; i++)
                    {
                        dr = dt.NewRow();
                        for (int j = 1; j <= columnCount; j++)
                        {
                            objCellValue = sheet.Cells[i, j].Value;
                            cellValue = objCellValue == null ? "" : objCellValue.ToString();
                            dr[j - 1] = cellValue;
                        }
                        dt.Rows.Add(dr);
                    }
                }
                return dt;
            }
        }
        public static bool ExportByModel<T>(string path, IEnumerable<T> ModelList, string SheetName = "Sheet1", bool WithTitle = true)
        {
            var result = false;
            FileInfo newFile = new FileInfo(path);
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(path);
            }
            using (var pkg = new ExcelPackage(newFile))
            {
                var ws = pkg.Workbook.Worksheets.Add(SheetName);
                ws.Cells[1, 1].LoadFromCollection<T>(ModelList, WithTitle);
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                pkg.Save();
                result = true;
            }
            return result;
        }
    }
}

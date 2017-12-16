using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 合并表
{
    public class EppHelper
    {
        public static bool ExportByDt(string path, DataTable Dt, string SheetName = "Sheet1", bool WithTitle = true)
        {
            var result = false;
            FileInfo newFile = new FileInfo(path);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(path);
            }
            using (var pkg = new ExcelPackage(newFile))
            {
                var ws = pkg.Workbook.Worksheets.Add(SheetName);
                ws.Cells[1, 1].LoadFromDataTable(Dt, WithTitle);
                //ws.Cells[ws.Dimension.Address].AutoFitColumns();
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

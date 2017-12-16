using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 合并表
{
    public static class ExcelHelper
    {
        public static List<string> GetSheetList(string filepath)
        {
            if (Path.GetExtension(filepath).Equals(".xlsx"))
            {
                return EppHelper.GetAllSheet(filepath);
            }
            else
            {
                return NpoiHelper.GetAllSheet(filepath);
            }
        }
        public static void SetDataSet(DataSet ds,FileAndSheets fs)
        {
            if (Path.GetExtension(fs.FileName).Equals(".xlsx"))
            {
                EppHelper.ReadExcelToDataSet(ds,fs);
            }
            else
            {
                NpoiHelper.ReadExcelToDataSet(ds,fs);
            }
        }
    }
}

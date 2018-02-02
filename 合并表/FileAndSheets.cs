using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 合并表
{
    public class FileAndSheets
    {
        public FileAndSheets()
        {
            SheetNames = new List<string>();
        }
        public string FileName { get; set; }
        public List<string> SheetNames { get; set; }
        public string CurrentSheetName { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitExcelFiles
{
    public class SplitFile
    {

        public string FileName { get; set; }
        public string SheetName { get; set; }
        public string FirstCell { get; set; }
        public string Prefixe { get; set; }
        public string Suffixe { get; set; }
        public string SelExtension { get; set; }
        public string KeyColName { get; set; }
        public string OutputFolder { get; set; }
    }
}

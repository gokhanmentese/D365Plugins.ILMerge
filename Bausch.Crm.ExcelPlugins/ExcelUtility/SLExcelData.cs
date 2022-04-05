using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crm.ExcelPlugins
{
    public class SLExcelStatus
    {
        public string Message { get; set; }
        public bool Success
        {
            get { return string.IsNullOrWhiteSpace(Message); }
        }
    }

    public class SLExcelData
    {
        public SLExcelStatus Status { get; set; }
        public Columns ColumnConfigurations { get; set; }
        public List<string> Headers { get; set; }
        public List<List<string>> DataRows { get; set; }
        public string SheetName { get; set; }

        public SLExcelData()
        {
            Status = new SLExcelStatus();
            Headers = new List<string>();
            DataRows = new List<List<string>>();
        }
    }

    public class ExcelColumn
    {
        public ExcelColumn()
        {
            IsIdentity = false;
        }

        public int Index { get; set; }
        public string Name { get; set; }
        public string Text { get; set; }
        public bool IsIdentity { get; set; }
    }
}

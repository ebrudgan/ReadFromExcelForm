using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadFromExcelForm
{
    class ExcelSheet
    {
        public int RowNum { get; set; }
        public string FullName { get; set; }
        public ActionType Action { get; set; }
        public DateTime Date { get; set; }
        public string Terminal { get; set; }

    }

    public enum ActionType
    {
        In,
        Out,
    }

}


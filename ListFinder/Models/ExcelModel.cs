using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListFinder.Models
{
    class ExcelModel
    {
        public ExcelModel()
        {

        }

        public ExcelModel(string fileName)
        {
            Application = new Excel.Application();
            Workbook = Application.Workbooks.Open(fileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet = (Excel.Worksheet) Workbook.Worksheets.get_Item(1);
            Range = Worksheet.UsedRange;
        }

        public bool IsValid()
        {
            return Application != null && Workbook != null && Worksheet != null && Range != null;
        }

        public Excel._Application Application { get; set; }

        public Excel.Workbook Workbook { get; set; }

        public Excel.Worksheet Worksheet { get; set; }

        public Excel.Range Range { get; set; }
    }
}

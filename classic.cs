using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel_12
{
    public class classic
    {
        public void nameSbv()
        {
            Excel.Application Xlapp = new Microsoft.Office.Interop.Excel.Application(); //initialize the component

            Excel.Workbook Book1;
            Excel.Worksheet Sheet1;
            object misValue = System.Reflection.Missing.Value;
            Book1 = Xlapp.Workbooks.Add(misValue);
            Sheet1 = (Excel.Worksheet) Book1.Worksheets.get_Item(1);
            Sheet1.Cells[1,1] = "ID";
            Sheet1.Cells[1,2] = "Name"; 


            Book1.SaveAs("C:\\Users\\user\\source\\repos\\Excel 12\\try3.xls",Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            Book1.Close(true, misValue,misValue);
            Xlapp.Quit();

            Marshal.ReleaseComObject(Book1);
            Marshal.ReleaseComObject(Sheet1);
            Marshal.ReleaseComObject(Xlapp);
        }


    }
}

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcelToPDFUsingCSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new Application();
            var pathToOpen = @"C:\Users\SPlay\Documents\products.xlsx";
            app.Workbooks.Open(pathToOpen);
            app.ActiveWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF);
            app.ActiveWorkbook.Close(false);
            app.Quit();
        }
    }
}

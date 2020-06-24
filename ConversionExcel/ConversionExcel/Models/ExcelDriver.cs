using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConversionExcel.Models
{
    public class ExcelDriver
    {
        public void Execute()
        {
            var readExcel = new ExcelDriverCore(@"C:\Users\aoike\Desktop\読込Excel.xlsx");
            var outputExcel = new ExcelDriverCore(@"C:\Users\aoike\Desktop\出力Excel.xlsx");

            using (var readPackage = new ExcelPackage(readExcel.FileInfo))
            using (var outputpackage = new ExcelPackage(outputExcel.FileInfo))
            {
                
            }
        }
    }
}
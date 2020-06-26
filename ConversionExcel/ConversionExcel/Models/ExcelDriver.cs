using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace ConversionExcel.Models
{
    public class ExcelDriver
    {
        public void Execute()
        //public void Execute(string readPath, string outputPath)
        {
            var readPath = @"C:\Users\aoike\Desktop\読込Excel.xlsx";
            var outputPath = @"C:\Users\aoike\Desktop\出力Excel.xlsx";

            if (!File.Exists(readPath)) return;

            var readExcel = new ExcelDriverCore(readPath);
            var outputExcel = new ExcelDriverCore(outputPath);

            using (var readPackage = new ExcelPackage(readExcel.FileInfo))
            using (var outputpackage = new ExcelPackage(outputExcel.FileInfo))
            {
                outputExcel.NewCreate(outputPath, outputpackage);

                // これ以下を生成していく
                outputExcel.Writing(outputpackage, "Sheet1", "A1", "hello");
            }
        }
    }
}
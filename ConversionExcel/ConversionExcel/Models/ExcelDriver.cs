using ConversionExcel.Enum;
using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Management;

namespace ConversionExcel.Models
{
    public class ExcelDriver
    {
        //public Results ReadConfiguration()
        //{

        //}
        public Results Execute(Parent model)
        {
            var readPath = model.ReadPath;
            var outputPath = model.OutputPath;

            if (!File.Exists(readPath)) return new Results() { Message = ConstValue.NOT_EXISTS_READ_EXCEL, HasError = true };

            var readExcel = new ExcelDriverCore(readPath);
            var outputExcel = new ExcelDriverCore(outputPath);
            var results = new Results();

            using (var readPackage = new ExcelPackage(readExcel.FileInfo))
            using (var outputPackage = new ExcelPackage(outputExcel.FileInfo))
            {
                outputExcel.NewCreate(outputPath, outputPackage);

                results = ExecuteCore(readExcel, outputExcel, readPackage, outputPackage, model);
            }

            return results;
        }
        private Results ExecuteCore(ExcelDriverCore readExcel, ExcelDriverCore outputExcel, ExcelPackage readPackage, ExcelPackage outputPackage, Parent model)
        {
            var count = 0;
            foreach (var process in model.Processes)
            {
                count++;

                if (process.Shori == ConstValue.WRITING)
                {
                    var results = outputExcel.Writing(outputPackage, process.Arg1, process.Arg2, process.Arg3);
                    if (results.HasError)
                    {
                        results.Message = ConstValue.Processing_Content + count + ":" + results.Message;
                        return results;
                    }
                    continue;
                }
            }

            return new Results() { Message = ConstValue.SUCCESS, HasError = false };
        }
    }
}
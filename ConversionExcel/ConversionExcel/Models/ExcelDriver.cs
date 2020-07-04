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
        public Results ReadConfiguration(string path)
        {
            if (!File.Exists(path)) return new Results() { Message = ConstValue.NOT_EXISTS_FILE };

            var configurationExcel = new ExcelDriverCore(path);
            var parent = new Parent();
            parent.Processes = new List<Process>();
            var process = new Process();

            using (var configurationPackage = new ExcelPackage(configurationExcel.FileInfo))
            {
                try
                {
                    parent.ReadPath = configurationExcel.Reading(configurationPackage, "実行設定", "B1");
                    parent.OutputPath = configurationExcel.Reading(configurationPackage, "実行設定", "B2");
                    var rowCount = 5;
                    while (!string.IsNullOrEmpty(process.Shori = configurationExcel.Reading(configurationPackage, "実行設定", "A" + rowCount)))
                    {
                        process.Arg1 = configurationExcel.Reading(configurationPackage, "実行設定", "B" + rowCount);
                        process.Arg2 = configurationExcel.Reading(configurationPackage, "実行設定", "C" + rowCount);
                        process.Arg3 = configurationExcel.Reading(configurationPackage, "実行設定", "D" + rowCount);
                        process.Arg4 = configurationExcel.Reading(configurationPackage, "実行設定", "E" + rowCount);
                        process.Arg5 = configurationExcel.Reading(configurationPackage, "実行設定", "F" + rowCount);
                        parent.Processes.Add(process);
                        process = new Process();
                        rowCount++;
                    }
                }
                catch (Exception e)
                {
                    return new Results() { Message = e.Message };
                }
            }
            return new Results() { Message = ConstValue.SUCCESS, Parent = parent };
        }
        public Results Execute(Parent parent)
        {
            var readPath = parent.ReadPath;
            var outputPath = parent.OutputPath;

            if (!File.Exists(readPath)) return new Results() { Message = ConstValue.NOT_EXISTS_FILE };

            var readExcel = new ExcelDriverCore(readPath);
            var outputExcel = new ExcelDriverCore(outputPath);
            var results = new Results();

            using (var readPackage = new ExcelPackage(readExcel.FileInfo))
            using (var outputPackage = new ExcelPackage(outputExcel.FileInfo))
            {
                try
                {
                    outputExcel.NewCreate(outputPath, outputPackage);
                }
                catch (Exception e)
                {
                    return new Results() { Message = e.InnerException.ToString() };
                }

                results = ExecuteCore(readExcel, outputExcel, readPackage, outputPackage, parent);
            }

            return results;
        }
        private Results ExecuteCore(ExcelDriverCore readExcel, ExcelDriverCore outputExcel, ExcelPackage readPackage, ExcelPackage outputPackage, Parent parent)
        {
            var count = 0;
            try
            {
                foreach (var process in parent.Processes)
                {
                    count++;

                    // 処理をどんどん増やしていく
                    if (process.Shori == null)
                    {
                        continue;
                    }
                    if (process.Shori == ConstValue.WRITING)
                    {
                        outputExcel.Writing(outputPackage, process.Arg1, process.Arg2, process.Arg3);
                        continue;
                    }
                }
            }
            catch (Exception e)
            {
                return new Results() { Message = ConstValue.PROCESSING_CONTENT + count + "：" + e.Message };
            }

            return new Results() { Message = ConstValue.SUCCESS };
        }
    }
}
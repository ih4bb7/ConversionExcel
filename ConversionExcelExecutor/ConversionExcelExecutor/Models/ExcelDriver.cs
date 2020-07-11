using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace ConversionExcelExecutor.Models
{
    public class ExcelDriver
    {
        public Results Save(Parent parent)
        {
            if (!File.Exists(parent.ConfigurationPath)) return new Results() { Message = ConstValue.NOT_EXISTS_CONFIGRATIONFILE, IsFile = false };

            var configurationFileInfo = new FileInfo(parent.ConfigurationPath);
            var configurationExcel = new ExcelDriverCore(configurationFileInfo);
            var results = SaveCore(configurationExcel, parent);

            return results;
        }
        public Results ReadConfiguration(string path)
        {
            if (!File.Exists(path)) return new Results() { Message = ConstValue.NOT_EXISTS_CONFIGRATIONFILE, IsFile = false };

            var configurationFileInfo = new FileInfo(path);
            var configurationExcel = new ExcelDriverCore(configurationFileInfo);
            var parent = new Parent();
            parent.Processes = new List<Process>();
            var process = new Process();

            try
            {
                parent.ReadPath = configurationExcel.Reading("実行設定", "B1");
                parent.WritePath = configurationExcel.Reading("実行設定", "B2");
                var rowCount = 5;
                while (!string.IsNullOrEmpty(process.Shori = configurationExcel.Reading("実行設定", "A" + rowCount)))
                {
                    process.Arg1 = configurationExcel.Reading("実行設定", "B" + rowCount);
                    process.Arg2 = configurationExcel.Reading("実行設定", "C" + rowCount);
                    process.Arg3 = configurationExcel.Reading("実行設定", "D" + rowCount);
                    process.Arg4 = configurationExcel.Reading("実行設定", "E" + rowCount);
                    process.Arg5 = configurationExcel.Reading("実行設定", "F" + rowCount);
                    parent.Processes.Add(process);
                    process = new Process();
                    rowCount++;
                }
            }
            catch (Exception e)
            {
                return new Results() { Message = e.Message, HasError = true };
            }
            finally
            {
                configurationExcel.Dispose();
            }

            return new Results() { Message = ConstValue.SUCCESS, Parent = parent };
        }
        public Results Execute(Parent parent)
        {
            var readPath = parent.ReadPath;
            var outputPath = parent.WritePath;

            if (!File.Exists(readPath)) return new Results() { Message = ConstValue.NOT_EXISTS_READFILE };

            var readFileInfo = new FileInfo(readPath);
            var outputFileInfo = new FileInfo(outputPath);
            var readExcel = new ExcelDriverCore(readFileInfo);
            var writeExcel = new ExcelDriverCore(outputFileInfo);

            try
            {
                writeExcel.NewCreate(outputPath);
            }
            catch (Exception e)
            {
                return new Results() { Message = e.InnerException.ToString() };
            }

            var results = ExecuteCore(readExcel, writeExcel, parent);

            return results;
        }
        private Results ExecuteCore(ExcelDriverCore readExcel, ExcelDriverCore writeExcel, Parent parent)
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
                        writeExcel.Writing(process.Arg1, process.Arg2, process.Arg3);
                        continue;
                    }
                    if (process.Shori == ConstValue.CELLCOPY_AND_PASTE)
                    {
                        var value = readExcel.Reading(process.Arg1, process.Arg2);
                        writeExcel.Writing(process.Arg3, process.Arg4, value);
                        continue;
                    }
                    if (process.Shori == ConstValue.ROWCOPY_AND_PASTE)
                    {
                        var value = readExcel.RowCopy(process.Arg1, int.Parse(process.Arg2));
                        writeExcel.RowPaste(process.Arg3, int.Parse(process.Arg4), value);
                        continue;
                    }
                    // 処理をどんどん増やしていく
                }
            }
            catch (Exception e)
            {
                return new Results() { Message = ConstValue.PROCESSING_CONTENT + count + "：" + e.Message };
            }
            finally
            {
                readExcel.Dispose();
                writeExcel.Dispose();
            }

            return new Results() { Message = ConstValue.SUCCESS };
        }
        private Results SaveCore(ExcelDriverCore configurationExcel, Parent parent)
        {
            try
            {
                for (int i = 0; i < parent.Processes.Count; i++)
                {
                    configurationExcel.Writing("実行設定", "A" + (i + 5), parent.Processes[i].Shori == null ? "" : parent.Processes[i].Shori);
                    configurationExcel.Writing("実行設定", "B" + (i + 5), parent.Processes[i].Arg1 == null ? "" : parent.Processes[i].Arg1);
                    configurationExcel.Writing("実行設定", "C" + (i + 5), parent.Processes[i].Arg2 == null ? "" : parent.Processes[i].Arg2);
                    configurationExcel.Writing("実行設定", "D" + (i + 5), parent.Processes[i].Arg3 == null ? "" : parent.Processes[i].Arg3);
                    configurationExcel.Writing("実行設定", "E" + (i + 5), parent.Processes[i].Arg4 == null ? "" : parent.Processes[i].Arg4);
                    configurationExcel.Writing("実行設定", "F" + (i + 5), parent.Processes[i].Arg5 == null ? "" : parent.Processes[i].Arg5);
                }
            }
            catch (Exception e)
            {
                return new Results() { Message = e.Message, HasError = true };
            }
            finally
            {
                configurationExcel.Dispose();
            }

            return new Results() { Message = ConstValue.SUCCESS };
        }
    }
}
using ConversionExcelExecutor.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConversionExcelExecutor
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelDriver = new ExcelDriver();
            var configuration = excelDriver.ReadConfiguration("設定Excel.xlsx");
            if (!configuration.IsFile)
            {
                ConsoleWriteLine(configuration.Message);
                return;
            }

            var results = excelDriver.Execute(configuration.Parent);
            ConsoleWriteLine(results.Message);
        }
        private static void ConsoleWriteLine(string message)
        {
            Console.WriteLine(message);
            Console.WriteLine("続行するには何かキーを押してください．．．");
            Console.ReadKey();
        }
    }
}

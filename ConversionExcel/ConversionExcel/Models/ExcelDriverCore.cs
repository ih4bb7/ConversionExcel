using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;

namespace ConversionExcel.Models
{
    public class ExcelDriverCore
    {
        public ExcelPackage ExcelPackage { get; set; }

        public ExcelDriverCore(FileInfo fileInfo)
        {
            ExcelPackage = new ExcelPackage(fileInfo);
        }
        /// <summary>
        /// 新規作成
        /// </summary>
        public void NewCreate(string path)
        {
            if (File.Exists(path)) return;

            ExcelPackage.Workbook.Worksheets.Add("Sheet1");
            ExcelPackage.Save();
        }
        /// <summary>
        /// 書き込み
        /// </summary>
        public void Writing(string sheetName, string cell, string value)
        {
            var sheet = ExcelPackage.Workbook.Worksheets[sheetName];
            sheet.Cells[cell].Value = value;
            ExcelPackage.Save();
            return;
        }
        /// <summary>
        /// 読み込み
        /// </summary>
        public string Reading(string sheetName, string cell)
        {
            var sheet = ExcelPackage.Workbook.Worksheets[sheetName];
            return sheet.Cells[cell].Text;
        }
        /// <summary>
        /// 解放
        /// </summary>
        public void Dispose()
        {
            ExcelPackage.Dispose();
        }
        /// <summary>
        /// 行コピー
        /// </summary>
        public ForRowCopyAndPaste RowCopy(string sheetName, int rowNum)
        {
            var result = new ForRowCopyAndPaste();
            var sheet = ExcelPackage.Workbook.Worksheets[sheetName];

            for (int i = 0; i < sheet.Dimension.Columns; i++)
            {
                result.rowValueData.Add(sheet.Cells[rowNum, i + 1].Value);
                result.rowStyleData.Add(sheet.Cells[rowNum, i + 1].StyleID);
            }

            return result;
        }
        /// <summary>
        /// 行ペースト
        /// </summary>
        public void RowPaste(string sheetName, int rowNum, ForRowCopyAndPaste value)
        {
            var sheet = ExcelPackage.Workbook.Worksheets[sheetName];

            for (int i = 0; i < value.rowValueData.Count; i++)
            {
                sheet.Cells[rowNum, i + 1].Value = value.rowValueData[i];
                sheet.Cells[rowNum, i + 1].StyleID = value.rowStyleData[i];
            }

            ExcelPackage.Save();
        }
        /// <summary>
        /// 行コピペ
        /// </summary>
        public static void RowCopyAndPaste(ExcelPackage sourcePackage, string sourceSheetName, ExcelPackage destPackage, string destSheetName, int sourceRowNum, int destRowNum)
        {
            var rowValueData = new List<object>();
            var rowStyleData = new List<int>();

            var sourceWorksheet = sourcePackage.Workbook.Worksheets[sourceSheetName];

            for (int i = 0; i < sourceWorksheet.Dimension.Columns; i++)
            {
                rowValueData.Add(sourceWorksheet.Cells[sourceRowNum, i + 1].Value);
                rowStyleData.Add(sourceWorksheet.Cells[sourceRowNum, i + 1].StyleID);
            }

            var destWorksheet = destPackage.Workbook.Worksheets[destSheetName];

            for (int i = 0; i < rowValueData.Count; i++)
            {
                destWorksheet.Cells[destRowNum, i + 1].Value = rowValueData[i];
                destWorksheet.Cells[destRowNum, i + 1].StyleID = rowStyleData[i];
            }

            destPackage.Save();
        }
        public class ForRowCopyAndPaste
        {
            public List<object> rowValueData;
            public List<int> rowStyleData;

            public ForRowCopyAndPaste()
            {
                rowValueData = new List<object>();
                rowStyleData = new List<int>();
            }
        }










        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////


        public static void RangeCopyAndPaste(ExcelPackage sourcePackage, string sourceSheetName, ExcelPackage destPackage, string destSheetName, int fromRow_Source, int fromCol_Source, int toRow_Source, int toCol_Source, int destRowNum, int destColumnNum)
        {
            CopyAndPaste(sourcePackage, sourceSheetName, destPackage, destSheetName, fromRow_Source, fromCol_Source, toRow_Source, toCol_Source, destRowNum, destColumnNum, null, null);
        }

        public static void RangeCopyAndPaste(ExcelPackage sourcePackage, string sourceSheetName, ExcelPackage destPackage, string destSheetName, string sourceAddress, string destAddress)
        {
            CopyAndPaste(sourcePackage, sourceSheetName, destPackage, destSheetName, 0, 0, 0, 0, 0, 0, sourceAddress, destAddress);
        }

        public static void SpecifyPrintArea(ExcelWorksheet worksheet, string address)
        {
            worksheet.PrinterSettings.PrintArea = worksheet.Cells[address];
        }

        public static void SpecifyPrintArea(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            worksheet.PrinterSettings.PrintArea = worksheet.Cells[fromRow, fromCol, toRow, toCol];
        }

        public static ExcelPackage CreateExcelPackage(FileInfo fileInfo, string password)
        {
            if (string.IsNullOrEmpty(password))
            {
                return new ExcelPackage(fileInfo);
            }

            return new ExcelPackage(fileInfo, password);
        }

        public static void BottomBorder(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = borderStyle;
        }

        public static void BottomBorder(ExcelWorksheet worksheet, string address, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[address].Style.Border.Bottom.Style = borderStyle;
        }

        public static void BorderAround(ExcelWorksheet worksheet, string address, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[address].Style.Border.BorderAround(borderStyle);
        }

        public static void TopBorder(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = borderStyle;
        }

        public static void TopBorder(ExcelWorksheet worksheet, string address, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[address].Style.Border.Top.Style = borderStyle;
        }

        public static void RightBorder(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = borderStyle;
        }

        public static void RightBorder(ExcelWorksheet worksheet, string address, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[address].Style.Border.Right.Style = borderStyle;
        }

        public static void LeftBorder(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = borderStyle;
        }

        public static void LeftBorder(ExcelWorksheet worksheet, string address, ExcelBorderStyle borderStyle)
        {
            worksheet.Cells[address].Style.Border.Left.Style = borderStyle;
        }

        public static void SetNumberFormat(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol, string format)
        {
            worksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Numberformat.Format = format;
        }

        public static void SetNumberFormat(ExcelWorksheet worksheet, string address, string format)
        {
            worksheet.Cells[address].Style.Numberformat.Format = format;
        }

        public static void SetHorizontalAlignment(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol, ExcelHorizontalAlignment horizontalAlignment)
        {
            worksheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = horizontalAlignment;
        }

        public static void SetHorizontalAlignment(ExcelWorksheet worksheet, string address, ExcelHorizontalAlignment horizontalAlignment)
        {
            worksheet.Cells[address].Style.HorizontalAlignment = horizontalAlignment;
        }

        public static void SetFontStyleAndSize(ExcelWorksheet worksheet, int fromRow, int fromCol, int toRow, int toCol, string fontStyle, int fontSize)
        {
            worksheet.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new System.Drawing.Font(fontStyle, fontSize));
        }

        public static void SetFontStyleAndSize(ExcelWorksheet worksheet, string address, string fontStyle, int fontSize)
        {
            worksheet.Cells[address].Style.Font.SetFromFont(new System.Drawing.Font(fontStyle, fontSize));
        }

        public static void FontBold(ExcelWorksheet worksheet, string address)
        {
            worksheet.Cells[address].Style.Font.Bold = true;
        }

        public static void CellMerge(ExcelWorksheet worksheet, string address)
        {
            worksheet.Cells[address].Merge = true;
        }

        public static void SetCellValue(ExcelWorksheet worksheet, string cell, string value)
        {
            worksheet.Cells[cell].Value = value;
        }

        public static void SetCellFormula(ExcelWorksheet worksheet, string cell, string value)
        {
            worksheet.Cells[cell].Formula = value;
        }

        public static string GetCellValue(ExcelWorksheet worksheet, string cell)
        {
            var value = worksheet.Cells[cell].Value ?? "";
            return value.ToString();
        }

        public static void DeleteColumn(ExcelWorksheet worksheet, int column)
        {
            DeleteColumn(worksheet, column, 1);
        }

        public static void DeleteColumn(ExcelWorksheet worksheet, int columnFrom, int columns)
        {
            for (int i = 0; i < columns; i++)
            {
                worksheet.DeleteColumn(columnFrom);
            }
        }

        public static void HiddenColumn(ExcelWorksheet worksheet, int column)
        {
            worksheet.Column(column).Hidden = true;
        }

        public static void HiddenColumn(ExcelWorksheet worksheet, int[] columns)
        {
            for (int i = 0; i < columns.Length; i++)
            {
                worksheet.Column(columns[i]).Hidden = true;
            }
        }

        public static void HiddenColumn(ExcelWorksheet worksheet, int columnFrom, int columns)
        {
            for (int i = 0; i < columns; i++)
            {
                worksheet.Column(columnFrom).Hidden = true;
                columnFrom++;
            }
        }

        //public static void HiddenColumn(ExcelWorksheet worksheet, string columnFrom, string columnTo)
        //{
        //    var columnNoFrom = columnFrom.ToColumnNumber();
        //    var columns = columnTo.ToColumnNumber() - columnNoFrom + 1;

        //    for (int i = 0; i < columns; i++)
        //    {
        //        worksheet.Column(columnNoFrom).Hidden = true;
        //        columnNoFrom++;
        //    }
        //}

        public static void HiddenRow(ExcelWorksheet worksheet, int rowFrom, int rows)
        {
            for (int i = 0; i < rows; i++)
            {
                worksheet.Row(rowFrom).Hidden = true;
                rowFrom++;
            }
        }

        public static void InsertRow(ExcelWorksheet worksheet, int rowFrom, int rows)
        {
            worksheet.InsertRow(rowFrom, rows);
        }

        public static void InsertColumn(ExcelWorksheet worksheet, int columnFrom, int columns)
        {
            worksheet.InsertColumn(columnFrom, columns);
        }

        public static void SetTopMargin(ExcelWorksheet worksheet, decimal topMargin)
        {
            topMargin = topMargin / (decimal)2.54;
            worksheet.PrinterSettings.TopMargin = topMargin;
        }

        public static void SetBottomMargin(ExcelWorksheet worksheet, decimal bottomMargin)
        {
            bottomMargin = bottomMargin / (decimal)2.54;
            worksheet.PrinterSettings.BottomMargin = bottomMargin;
        }

        public static void SetRightMargin(ExcelWorksheet worksheet, decimal rightMargin)
        {
            rightMargin = rightMargin / (decimal)2.54;
            worksheet.PrinterSettings.RightMargin = rightMargin;
        }

        public static void SetLeftMargin(ExcelWorksheet worksheet, decimal leftMargin)
        {
            leftMargin = leftMargin / (decimal)2.54;
            worksheet.PrinterSettings.LeftMargin = leftMargin;
        }

        public static void SetHeaderMargin(ExcelWorksheet worksheet, decimal headerMargin)
        {
            headerMargin = headerMargin / (decimal)2.54;
            worksheet.PrinterSettings.HeaderMargin = headerMargin;
        }

        public static void SetFooterMargin(ExcelWorksheet worksheet, decimal footerMaegin)
        {
            footerMaegin = footerMaegin / (decimal)2.54;
            worksheet.PrinterSettings.FooterMargin = footerMaegin;
        }

        public static void SetPrintTitle(ExcelWorksheet worksheet, string address)
        {
            worksheet.PrinterSettings.RepeatRows = new ExcelAddress(address);
        }

        public static void FitToHeight(ExcelWorksheet worksheet)
        {
            worksheet.PrinterSettings.FitToHeight = 1;
        }

        public static void SetPaperSize(ExcelWorksheet worksheet, ePaperSize paperSize)
        {
            worksheet.PrinterSettings.PaperSize = paperSize;
        }

        public static void SetPrintHorizontalCentered(ExcelWorksheet worksheet)
        {
            worksheet.PrinterSettings.HorizontalCentered = true;
        }

        public static void SetPrintScale(ExcelWorksheet worksheet, int scale)
        {
            worksheet.PrinterSettings.Scale = scale;
        }

        public static void AutoFitColumns(ExcelWorksheet worksheet, string address)
        {
            worksheet.Cells[address].AutoFitColumns();
        }

        public static void SetColumnWidth(ExcelWorksheet worksheet, int column, double width)
        {
            worksheet.Column(column).Width = width;
        }

        public static void ClearCell(ExcelWorksheet worksheet, string address)
        {
            worksheet.Cells[address].Clear();
        }

        public static void IsShowGridLines(ExcelWorksheet worksheet, bool isShowGridLines)
        {
            worksheet.View.ShowGridLines = isShowGridLines;
        }

        public static void SetCharacterColor(ExcelWorksheet worksheet, string address, Color color)
        {
            worksheet.Cells[address].Style.Font.Color.SetColor(color);
        }

        public static void SetBorderBottomColor(ExcelWorksheet worksheet, string address, Color color)
        {
            //worksheet.Cells[address].Style.Border.Bottom.Color.SetColor(color);
        }

        public static void SetBorderLeftColor(ExcelWorksheet worksheet, string address, Color color)
        {
            //worksheet.Cells[address].Style.Border.Left.Color.SetColor(color);
        }

        public static void SetBackGroundColor(ExcelWorksheet worksheet, string address, Color color)
        {
            worksheet.Cells[address].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[address].Style.Fill.BackgroundColor.SetColor(color);
        }

        private static void CopyAndPaste(ExcelPackage sourcePackage, string sourceSheetName, ExcelPackage destPackage, string destSheetName, int fromRow_Source, int fromCol_Source, int toRow_Source, int toCol_Source, int destRowNum, int destColumnNum, string sourceAddress, string destAddress)
        {
            var sourceWorksheet = sourcePackage.Workbook.Worksheets[sourceSheetName];
            var destWorksheet = destPackage.Workbook.Worksheets[destSheetName];

            if (string.IsNullOrEmpty(sourceAddress) && string.IsNullOrEmpty(destAddress))
            {
                sourceWorksheet.Cells[fromRow_Source, fromCol_Source, toRow_Source, toCol_Source].Copy(destWorksheet.Cells[destRowNum, destColumnNum]);
            }
            else
            {
                sourceWorksheet.Cells[sourceAddress].Copy(destWorksheet.Cells[destAddress]);
            }

            destPackage.Save();
        }

        private static void SetRowValueString(ExcelWorksheet worksheet, int rowIndex, List<string> dataList)
        {
            for (var i = 0; i < dataList.Count; i++)
            {
                worksheet.Cells[rowIndex, i + 1].Value = dataList[i];
            }
        }


    }
}
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversionExcelExecutor.Models
{
    public static class ConstValue
    {
        public static string SUCCESS = "成功";
        public static string CHANGE_READFILE = "読み込みExcelを読み込みなおしてください";
        public static string NOT_EXISTS_READFILE = "読み込みExcelがありません";
        public static string NOT_EXISTS_CONFIGRATIONFILE = "設定Excelが存在しません";
        public static string PROCESSING_CONTENT = "処理内容";

        #region 処理内容

        public static string WRITING = "書き込み";
        public static string CELLCOPY_AND_PASTE = "セルコピペ";
        public static string ROWCOPY_AND_PASTE = "行コピペ";

        #endregion
    }
}
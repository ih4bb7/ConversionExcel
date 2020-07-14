using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConversionExcelExecutor.Models
{
    public static class ConstValue
    {
        public static string SUCCESS = "成功";
        public static string NOT_EXISTS_CONFIGRATIONFILE = "設定Excelが存在しません";
        public static string PROCESSING_CONTENT = "処理内容";

        #region 処理内容

        public static string WRITING = "書き込み";
        public static string CELLCOPY_AND_PASTE = "セルコピペ";
        public static string ROWCOPY_AND_PASTE = "行コピペ";
        public static string NUMBERWRITING = "数字書き込み";
        public static string FORMULAWRITING = "関数書き込み";
        public static string RANGECOPY_AND_PASTE = "範囲コピペ";

        #endregion
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConversionExcelExecutor.Models
{
    public class Results
    {
        public string Message;
        public Parent Parent;
        public string PartialView;
        public bool IsFile = true;
        public bool HasError;
        public string Path;
    }
}
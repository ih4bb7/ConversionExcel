using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversionExcelExecutor.Models
{
    public class Results
    {
        public string Message;
        public Parent Parent;
        public string PartialView;
        public bool IsFile = true;
        public bool HasError;
    }
}
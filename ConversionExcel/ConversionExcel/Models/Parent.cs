﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConversionExcel.Models
{
    public class Parent
    {
        public string ReadPath { get; set; }
        public List<Process> Processes { get; set; }
        public string OutputPath { get; set; }
    }
}
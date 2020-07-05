﻿using ConversionExcel.Models;
using ConversionExcel.Enum;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.HtmlControls;
using Microsoft.Ajax.Utilities;
using System.Web.Helpers;
using System.Text;
using System.IO;

namespace ConversionExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Count = 1;
            var parent = new Parent()
            {
                ReadPath = "",
                OutputPath = "",
                Processes = new List<Process>
                {
                    new Process(),
                },
            };
            return View(parent);
        }
        public ActionResult add_Click(int count)
        {
            ViewBag.Count = count + 1;
            return PartialView("_Processes", new Process());
        }
        public JsonResult execute_Click(Parent parent)
        {
            var excelDriver = new ExcelDriver();
            var results = excelDriver.Execute(parent);
            return Json(new { result = results.Message });
        }
        public JsonResult readConfiguration_Click(string path)
        {
            if (!System.IO.File.Exists(path)) return Json(new { result = new Results() { IsFile = false } });
            var excelDriver = new ExcelDriver();
            var results = excelDriver.ReadConfiguration(path);
            if (results.HasError) return Json(new { result = results });
            results.PartialView = CreatePartialView();
            return Json(new { result = results });
        }
        public JsonResult save_Click(Parent parent)
        {
            var excelDriver = new ExcelDriver();
            var results = excelDriver.Save(parent);
            return Json(new { result = results.Message });
        }
        private string CreatePartialView()
        {
            var partialView = new StringBuilder();
            partialView.Append("<div class='container' id='process_Count'>" + Environment.NewLine);
            partialView.Append("    <div class='form-group row'>" + Environment.NewLine);
            partialView.Append("        <label for='' class='col-sm-2 col-form-label'>処理内容 Count：</label>" + Environment.NewLine);
            partialView.Append("        <div class='col-sm-10'>" + Environment.NewLine);
            partialView.Append("            <select id='shori_Count' class='form-control' onchange='selectChange()'>" + Environment.NewLine);
            partialView.Append("                <option></option>" + Environment.NewLine);
            partialView.Append("                <option>書き込み</option>" + Environment.NewLine);
            partialView.Append("            </select>" + Environment.NewLine);
            partialView.Append("            <p class='form-inline'>" + Environment.NewLine);
            partialView.Append("                <input type='text' class='form-control' id='argument1_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                <input type='text' class='form-control' id='argument2_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                <input type='text' class='form-control' id='argument3_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                <input type='text' class='form-control' id='argument4_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                <input type='text' class='form-control' id='argument5_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("            </p>" + Environment.NewLine);
            partialView.Append("        </div>" + Environment.NewLine);
            partialView.Append("    </div>" + Environment.NewLine);
            partialView.Append("</div>" + Environment.NewLine);

            return partialView.ToString();
        }
    }
}
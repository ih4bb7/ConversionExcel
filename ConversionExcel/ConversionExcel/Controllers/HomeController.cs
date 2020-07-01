using ConversionExcel.Models;
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
            return Json(new{ result = results.Message });
        }

        //public JsonResult readConfiguration_Click()
        //{
            
        //    return Json(new{ result = results.Message });
        //}
    }
}
using ConversionExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.HtmlControls;

namespace ConversionExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var parent = new Parent()
            {
                ReadPath = "aaa",
                Processes = new List<Process>
                {
                    new Process() {Content = "2月", Arg1 = "arg1", Arg2 = "arg2", Arg3 = "arg3", Arg4 = "arg4", Arg5 = "arg5"},
                    new Process() {Content = "3月", Arg1 = "arg6", Arg2 = "arg7", Arg3 = "arg8", Arg4 = "arg9", Arg5 = "arg10"},
                },
                OutputPath = "bb",
            };
            return View(parent);
        }

        public ActionResult btnAdd_Click()
        {
            return PartialView("Processes", new Process());
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
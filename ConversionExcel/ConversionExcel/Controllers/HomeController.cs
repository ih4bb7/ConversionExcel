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
        public ActionResult Index(string data)
        {
            return View();
        }

        public ActionResult btnAdd_Click()
        {
            var newdivs = new HtmlGenericControl("DIV");
            newdivs.Attributes.Add("class", "maindivs");

            //maindivs.Controls.Add(newdivs);
            return View("Index");
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
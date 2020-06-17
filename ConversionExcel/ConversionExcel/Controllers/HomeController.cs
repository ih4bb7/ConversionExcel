﻿using ConversionExcel.Models;
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
            ViewBag.Count = 1;
            var parent = new Parent()
            {
                ReadPath = "",
                Processes = new List<Process>
                {
                    new Process(),
                },
                OutputPath = "",
            };
            return View(parent);
        }

        public ActionResult btnAdd_Click(int count)
        {
            ViewBag.Count = count + 1;
            return PartialView("_Processes", new Process());
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
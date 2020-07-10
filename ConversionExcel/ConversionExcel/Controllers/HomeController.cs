using ConversionExcel.Models;
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
using System.Security.Cryptography;

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
                WritePath = "",
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
            results.Path = Path.GetFileName(Path.GetDirectoryName(parent.WritePath)) + "/" + Path.GetFileName(parent.WritePath);
            if (results.IsFile && !results.HasError) results.Message += Environment.NewLine + "ダウンロードが終了してからOKを押してください";
            return Json(new { result = results });
        }
        public JsonResult UploadReadFileForExecute()
        {
            return Upload(Request.Files[0], ConstValue.NOT_EXISTS_READFILE);
        }
        public JsonResult UploadWriteFileForExecute()
        {
            var file = Request.Files[0];
            var fileName = string.Empty;
            var filePath = string.Empty;
            var dir = Path.Combine(@"C:\作業\Kelpex\Kelpex\Upload", DateTime.Now.ToString("yyyyMMddHHmmss"));
            Directory.CreateDirectory(dir);
            if (file.ContentLength == 0)
            {
                fileName = "書き込みExcel.xlsx";
                filePath = Path.Combine(dir, fileName);
            }
            else
            {
                fileName = Path.GetFileName(file.FileName);
                filePath = Path.Combine(dir, fileName);
                file.SaveAs(filePath);
            }

            return Json(new { result = new Results() { Path = filePath } });
        }
        public JsonResult UploadForReadConfiguration()
        {
            return Upload(Request.Files[0], "");
        }
        private JsonResult Upload(HttpPostedFileBase file, string constValue)
        {
            if (file.ContentLength == 0) return Json(new { result = new Results() { IsFile = false, Message = constValue } });

            var dir = Path.Combine(@"C:\作業\Kelpex\Kelpex\Upload", DateTime.Now.ToString("yyyyMMddHHmmss"));
            Directory.CreateDirectory(dir);
            var filePath = Path.Combine(dir, Path.GetFileName(file.FileName));

            file.SaveAs(filePath);
            return Json(new { result = new Results() { Path = filePath } });
        }
        public JsonResult DeleteAfterSave(string path)
        {
            var filePath = Path.Combine(@"C:\作業\Kelpex\Kelpex\Upload", path);
            System.IO.File.Delete(filePath);
            Directory.Delete(Path.GetDirectoryName(filePath));
            return Json(new { result = new Results() });
        }
        public JsonResult DeleteAfterExecute(string readPath, string writePath)
        {
            var readFilePath = Path.Combine(@"C:\作業\Kelpex\Kelpex\Upload", readPath);
            var writeFilePath = Path.Combine(@"C:\作業\Kelpex\Kelpex\Upload", writePath);
            System.IO.File.Delete(readFilePath);
            System.IO.File.Delete(writeFilePath);
            Directory.Delete(Path.GetDirectoryName(readFilePath));
            return Json(new { result = new Results() });
        }
        public JsonResult readConfiguration_Click(string path)
        {
            var excelDriver = new ExcelDriver();
            var results = excelDriver.ReadConfiguration(path);
            if (results.HasError || !results.IsFile) return Json(new { result = results });
            results.PartialView = CreatePartialView();
            System.IO.File.Delete(path);
            Directory.Delete(Path.GetDirectoryName(path));
            return Json(new { result = results });
        }
        public JsonResult save_Click(Parent parent)
        {
            var datetimeNow = DateTime.Now.ToString("yyyyMMddHHmmss");
            var dir = Path.Combine(@"C:\作業\Kelpex\Kelpex\Upload", datetimeNow);
            Directory.CreateDirectory(dir);
            var filePath = Path.Combine(dir, "設定Excel.xlsx");
            System.IO.File.Copy(@"C:\作業\Kelpex\Kelpex\Executor\設定Excel.xlsx", filePath);
            parent.ConfigurationPath = filePath;
            var excelDriver = new ExcelDriver();
            var results = excelDriver.Save(parent);
            results.Path = datetimeNow;
            results.Message += Environment.NewLine + "ダウンロードが終了してからOKを押してください";
            return Json(new { result = results });
        }
        private string CreatePartialView()
        {
            var partialView = new StringBuilder();
            partialView.Append("<div class='container' id='process_Count'>" + Environment.NewLine);
            partialView.Append("    <div class='form-group row'>" + Environment.NewLine);
            partialView.Append("        <label for='' class='col-sm-2 col-form-label'>処理内容 Count：</label>" + Environment.NewLine);
            partialView.Append("        <div class='col-sm-10'>" + Environment.NewLine);
            partialView.Append("            <form>" + Environment.NewLine);
            partialView.Append("                <div class='col-sm-12'>" + Environment.NewLine);
            partialView.Append("                    <select id='shori_Count' class='form-control' onchange='selectChange()'>" + Environment.NewLine);
            partialView.Append("                        <option></option>" + Environment.NewLine);
            partialView.Append("                        <option>書き込み</option>" + Environment.NewLine);
            partialView.Append("                    </select>" + Environment.NewLine);
            partialView.Append("                    <p class='form-inline'>" + Environment.NewLine);
            partialView.Append("                        <input type='text' class='form-control' id='argument1_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                        <input type='text' class='form-control' id='argument2_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                        <input type='text' class='form-control' id='argument3_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                        <input type='text' class='form-control' id='argument4_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                        <input type='text' class='form-control' id='argument5_Count' style='width:19.6%' readonly='readonly'>" + Environment.NewLine);
            partialView.Append("                    </p>" + Environment.NewLine);
            partialView.Append("                </div>" + Environment.NewLine);
            partialView.Append("            </form>" + Environment.NewLine);
            partialView.Append("        </div>" + Environment.NewLine);
            partialView.Append("    </div>" + Environment.NewLine);
            partialView.Append("</div>" + Environment.NewLine);

            return partialView.ToString();
        }
    }
}
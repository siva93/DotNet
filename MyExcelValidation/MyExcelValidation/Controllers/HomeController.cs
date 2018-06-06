using ClosedXML.Excel;
using MyExcelValidation.Models;
using System;
using System.Linq;
using System.Web.Mvc;
using System.Data;
using System.IO;

namespace MyExcelValidation.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        
        public  ActionResult UploadValidationTemplate()
        {
            return PartialView("_UploadValidationTemplate");
        }

        [HttpPost]
        public ActionResult UploadValidationTemplate(TemplateFileDTO templateFileDTO)
        {
            if (!ModelState.IsValid)
                return View(templateFileDTO);
            try
            {
                string path = Server.MapPath("~/App_Data/");
                string filename = Path.GetFileName(templateFileDTO.SheetName);
                templateFileDTO.Sheet.SaveAs(Path.Combine(path, filename));
            }catch(Exception ex)
            {

            }
            return Index();
        }

        [HttpGet]
        public ActionResult ValidateFile()
        {
            return PartialView("_ValidateFile");
        }

        [HttpPost]
        public ActionResult ValidateFile(ValidationFileDTO ValidationFileDTO)
        {
            try
            {
                if (!ModelState.IsValid || Request == null)
                {
                    return View(ValidationFileDTO);
                }
                var uploadedFile = ValidationFileDTO.Sheet;
                if (uploadedFile != null && uploadedFile.ContentLength > 0 && !string.IsNullOrEmpty(uploadedFile.FileName))
                {
                    using (XLWorkbook workBook = new XLWorkbook(uploadedFile.InputStream))
                    {
                        var dataTable = Helper.Helper.ConvertToDataTable(workBook, ValidationFileDTO.SheetName);
                        if (dataTable != null)
                        {
                          // var invalidData = Helper.Helper.ValidatRules(dataTable);
                            //if (invalidData.Any())
                            //{
                                using (XLWorkbook export = new XLWorkbook())
                                {
                                    //var errorData = invalidData.CopyToDataTable();
                                    export.Worksheets.Add(dataTable, "Validation Result");
                                    using (MemoryStream stream = new MemoryStream())
                                    {
                                        export.SaveAs(stream);
                                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ErrorValidation.xlsx");
                                    }
                                }
                            //}
                        }
                    }
                }

                return View("ValidationResult");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            return View();
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
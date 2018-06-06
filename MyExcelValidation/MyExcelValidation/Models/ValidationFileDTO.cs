
using System.ComponentModel.DataAnnotations;
using System.Web;

namespace MyExcelValidation.Models
{
    public class ValidationFileDTO
    {
        
        [Display(Name = "Sheet Name")]
        [Required(ErrorMessage = "Please enter the sheet name")]
        public string SheetName { get; set; }

        [DataType(DataType.Upload)]
        [Display(Name = "Upload File")]
        [Required(ErrorMessage = "Please choose file to upload.")]
        public HttpPostedFileBase Sheet { get; set; }

        [Display(Name = "Sheet Type")]
        //[Required(ErrorMessage = "Please select the sheet type")]
        public string ValidationSheetName { get; set; }
    }
}
using ExcelDataExtract.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelDataExtract.Controllers
{
    public class ExcelController : Controller
    {   
        private readonly IExcelLogicService _excelService;

        public ExcelController(IExcelLogicService excelService)
        {
            _excelService = excelService;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ExtractExcelData(IFormFile postedFile)
        {
            ViewBag.Success = _excelService.GetExcelData(postedFile);
            return View("Index");
        }

    }
}

using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.ComponentModel.DataAnnotations;

namespace ExcelFile.Controllers
{
    public class Person
       
    {
        public Person()
        {
            
        }
        [Display(Name="ایدی")]
        public int Id { get; set; }
        [Display(Name = "نام")]
        public string Name { get; set; }
        [Display(Name = "سن")]
        public int Age { get; set; }
    }

    [ApiController]
    [Route("[controller]")]
    public class ListExportController : ControllerBase
    {
        private readonly IExcelService _exellService;

        public ListExportController(IExcelService exellService)
        {
            this._exellService = exellService;
        }



        [HttpPost("ExportList")]
        public IActionResult ExportList()
        {

            var data = new List<Person>
        {
            new Person { Id = 1, Name = "John Doe", Age = 30 },
            new Person { Id = 2, Name = "Jane Smith", Age = 25 },
            new Person { Id = 3, Name = "Alice Johnson", Age = 35 }
        };


            var result= _exellService.ExportToExcel(data,true,true);
            return File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Export.xlsx");



        }
        [HttpPost("ImportFromExcel")]
        public IActionResult ImportFromExcel(IFormFile formFile)
        {

            var result = _exellService.ImportFromExcel<Person>(formFile);
            return Ok(result);



        }
    }
}

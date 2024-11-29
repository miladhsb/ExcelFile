using Microsoft.AspNetCore.Mvc;

namespace ExcelFile
{
    public interface IExcelService
    {
        byte[] ExportToExcel<T>([FromBody] List<T> data, bool createTable = false, bool rightToLeft = false);
        List<T> ImportFromExcel<T>(IFormFile file) where T : new();
    }
}
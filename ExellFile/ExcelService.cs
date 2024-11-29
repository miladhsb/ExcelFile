using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace ExcelFile
{
    public class ExcelService : IExcelService
    {

        public List<T> ImportFromExcel<T>(IFormFile file) where T : new()
        {
            if (file == null || file.Length == 0)
                throw new ArgumentNullException("File is null or empty.");




            var result = new List<T>();

            try
            {
                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);

                    using (var workbook = new XLWorkbook(stream))
                    {
                        var worksheet = workbook.Worksheet(1); // اولین شیت اکسل
                        var properties = typeof(T).GetProperties(); // پراپرتی‌های مدل T

                        // فرض بر این است که اولین ردیف شامل عناوین ستون‌ها است
                        var headers = new Dictionary<string, int>();
                        foreach (var cell in worksheet.Row(1).CellsUsed())
                        {
                            headers[cell.Value.ToString()] = cell.Address.ColumnNumber;
                        }
                        // خواندن داده‌ها از سطرها
                        var rows = worksheet.RowsUsed().Skip(1); // فرض می‌کنیم سطر اول هدر است

                        foreach (var row in rows)
                        {
                            var obj = new T(); // آبجکت جدید از نوع T
                            foreach (var property in properties)
                            {
                                var displayAttr = property.GetCustomAttributes(typeof(DisplayAttribute), true)
                                    .FirstOrDefault() as DisplayAttribute;

                                var header = displayAttr != null ? displayAttr.Name : property.Name;


                               
                                var cell = row.Cell(headers[header]);

                                if (property.PropertyType == typeof(int))
                                {
                                   
                                    property.SetValue(obj, cell.GetValue<int>());
                                }
                                if (property.PropertyType == typeof(double))
                                {

                                    property.SetValue(obj, cell.GetValue<double>());
                                }
                                if (property.PropertyType == typeof(DateTime))
                                {

                                    property.SetValue(obj, cell.GetValue<DateTime>());
                                }
                                if (property.PropertyType == typeof(string))
                                {

                                    property.SetValue(obj, cell.GetValue<string>());
                                }
                                if (property.PropertyType == typeof(TimeSpan))
                                {

                                    property.SetValue(obj, cell.GetValue<TimeSpan>());
                                }
                                 
                            
                                
                            }
                            result.Add(obj);
                        }
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }







        public byte[] ExportToExcel<T>([FromBody] List<T> data, bool createTable = false, bool rightToLeft = false)
        {
            if (data == null || !data.Any())
                throw new ArgumentNullException("data");

            using (var workbook = new XLWorkbook())
            {
                workbook.RightToLeft = rightToLeft;
                var tt = typeof(T).Name;
                var worksheet = workbook.Worksheets.Add("Worksheet01");

                // 1. گرفتن پراپرتی‌های کلاس T با Reflection
                var properties = typeof(T).GetProperties();

                // 2. اضافه کردن هدرها به اکسل
                for (int i = 0; i < properties.Length; i++)
                {
                    // بررسی Display Attribute برای پراپرتی
                    var displayAttribute = properties[i].GetCustomAttribute<DisplayAttribute>(true);


                    var header = displayAttribute != null ? displayAttribute.Name : properties[i].Name;
                    worksheet.Cell(1, i + 1).Value = header; // نام پراپرتی یا مقدار فارسی
                }

                // 3. پر کردن داده‌ها به صورت پویا
                for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
                {
                    var item = data[rowIndex];
                    for (int colIndex = 0; colIndex < properties.Length; colIndex++)
                    {
                        var value = properties[colIndex].GetValue(item); // گرفتن مقدار پراپرتی

                        worksheet.Cell(rowIndex + 2, colIndex + 1).Value = XLCellValue.FromObject(value); // مقدار یا رشته خالی
                    }
                }

                if (createTable)
                {
                    // 4. تبدیل داده‌ها به جدول
                    var range = worksheet.Range(1, 1, data.Count + 1, properties.Length);
                    var table = range.CreateTable();
                    table.Theme = XLTableTheme.TableStyleMedium9;
                }

                // 5. تنظیمات خروجی
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    return stream.ToArray();
                }
            }
        }
    }

}

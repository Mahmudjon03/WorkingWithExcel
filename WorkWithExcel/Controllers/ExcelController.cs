using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using WorkWithExcel;
using WorkWithExcel.Model;

namespace YourNamespace.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        [HttpPost]
        public IActionResult GenerateExcel(UserData userData)
        {
            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Данные пользователя");
                worksheet.Cells[1, 1].Value = "FName";
                worksheet.Cells[1, 2].Value = "lName";
                worksheet.Cells[1, 3].Value = "Email";
                worksheet.Cells[2, 1].Value = userData.FName;
                worksheet.Cells[2, 2].Value = userData.LName;
                worksheet.Cells[2, 3].Value = userData.Email;
                var memoryStream = new MemoryStream();
                excelPackage.SaveAs(memoryStream);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "UserData.xlsx");
            }
        }


      
    }
}
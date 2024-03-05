using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using WorkWithExcel.Model;
using WorkWithExcel.Services;

namespace WorkWithExcel.Controllers
{
        [ApiController]
        [Route("[controller]")]
    public class ExcelDateController:ControllerBase
    { private readonly DataForExcel _service;

        public ExcelDateController()
        {
            _service = new DataForExcel();
        }

        [HttpPost("Date")]
        public IActionResult GenetrateDate()
        {
           var model = _service.GetData();
            using (var excelPackege = new ExcelPackage())
            {
                int n = 2;
                var sheet = excelPackege.Workbook.Worksheets.Add("Data");
                sheet.Cells[1, 1].Value = "RowNumber";
                sheet.Cells[1, 2].Value = "MerchantName";
                sheet.Cells[1, 3].Value = "Amnt";
                sheet.Cells[1, 4].Value = "Reward";
                sheet.Cells[1, 5].Value = "Cnt";
                sheet.Cells["A1:E1"].Style.Font.Bold = true;
                sheet.Cells[1, 1, model.Count+1, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                sheet.Cells[1, 1, model.Count+1, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                foreach (var item in model)
                {
                    sheet.Cells[n, 1].Value = item.RowNumber;
                    sheet.Cells[n, 2].Value = item.merchat;
                    sheet.Cells[n, 3].Value = item.amnt;
                    sheet.Cells[n, 4].Value = item.reward;
                    sheet.Cells[n, 5].Value = item.cnt;
                    n++;
                }
                var memoryStream = new MemoryStream();
                excelPackege.SaveAs(memoryStream);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");
            }
        }

    }
}

using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using WorkWithExcel.Model;
using WorkWithExcel.Services;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace WorkWithExcel.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelDateController : ControllerBase
    {
        private readonly DataForExcel _service;
       
        public ExcelDateController()
        {
            _service = new DataForExcel();
        }

        [HttpPost("Date")]
        public IActionResult MerchantReconciliationReport(DateTime fromDate, DateTime toDate)
        {
            var model = _service.GetData(fromDate, toDate);
            using (var excelPackege = new ExcelPackage())
            {
                int n = 2;
                var sheet = excelPackege.Workbook.Worksheets.Add("Data");
                sheet.Cells[1, 1].Value = "RowNumber";
                sheet.Cells[1, 2].Value = "MerchantName";
                sheet.Cells[1, 3].Value = "Amnt";
                sheet.Cells[1, 4].Value = "Reward";
                sheet.Cells[1, 5].Value = "CreditAmount";
                sheet.Cells[1, 6].Value = "Cnt";
                sheet.Cells["A1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A1:F1"].Style.Font.Bold = true;
                sheet.Cells[1, 1, model.Count + 1, 6].Style.Border.BorderAround(ExcelBorderStyle.Double);
                sheet.Cells[1, 1, model.Count + 1, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                foreach (var item in model)
                {
                    sheet.Cells[n, 1].Value = item.RowNumber;
                    sheet.Cells[n, 2].Value = item.MerchantName;
                    sheet.Cells[n, 3].Value = item.Amnt;
                    sheet.Cells[n, 4].Value = item.Reward;
                    sheet.Cells[n, 5].Value = item.CreditAmount;
                    sheet.Cells[n, 6].Value = item.Cnt;
                    n++;
                }

                var memoryStream = new MemoryStream();
                excelPackege.SaveAs(memoryStream);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");
            }
            // return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");

        }
        [HttpPost("RegisterOfMerchantPayments")]
        public IActionResult RegisterOfMerchantPayments([FromQuery] DateTime fromDate, [FromQuery] DateTime toDate)
        {
            var model = _service.GetDateExcel(fromDate, toDate);
            using (var excelPackege = new ExcelPackage())
            {
                int n = 2;
                var sheet = excelPackege.Workbook.Worksheets.Add("Data");
                sheet.Cells[1, 1].Value = "Row Number";
                sheet.Cells[1, 2].Value = "Order Number";
                sheet.Cells[1, 3].Value = "Merchant Name";
                sheet.Cells[1, 4].Value = "Sum";
                sheet.Cells[1, 5].Value = "Date";
                sheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A:C"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["E:F"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[1, 6].Value = "Pan Card ";
                sheet.Cells["A1:F1"].Style.Font.Bold = true;
                sheet.Cells[1, 1, model.Count + 1, 6].Style.Border.BorderAround(ExcelBorderStyle.Double);
                sheet.Cells[1, 1, model.Count + 1, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                sheet.Cells.AutoFitColumns();

                foreach (var item in model)
                {
                    sheet.Cells[n, 1].Value = item.RowNumber;
                    sheet.Cells[n, 2].Value = item.OrderNumber;
                    sheet.Cells[n, 3].Value = item.MerchantName;
                    sheet.Cells[n, 4].Value = item.Sum.ToString();
                    sheet.Cells[n, 5].Value = item.Date;
                    sheet.Cells[n, 6].Value = item.panCard;

                    n++;
                }

                var memoryStream = new MemoryStream();
                excelPackege.SaveAs(memoryStream);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");
            }
        }

    }
}

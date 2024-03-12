using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using WorkWithExcel.Model;
using WorkWithExcel.Services;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style.XmlAccess;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.ChartDrawing;
using System.Drawing;

namespace WorkWithExcel.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelDateController : ControllerBase
    {
        private readonly DataForExcel _service;
        private readonly ExcelImport _excel;
       
        public ExcelDateController()
        {
            _service = new DataForExcel();
            _excel = new ExcelImport();
        }

        [HttpPost("Date")]
        public IActionResult MerchantReconciliationReport(DateTime fromDate, DateTime toDate)
        {
          
            var excelPackege = _excel.ExcelService2(fromDate ,toDate);
            var memoryStream = new MemoryStream();
            excelPackege.SaveAs(memoryStream);
            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");
            
            // return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");

        }
        [HttpPost("RegisterOfMerchantPayments")]
        public IActionResult RegisterOfMerchantPayments([FromQuery] DateTime fromDate, [FromQuery] DateTime toDate)
        {
            var excelPackege = _excel.ExcelService(fromDate, toDate);
            MemoryStream memoryStream = new MemoryStream();
            excelPackege.SaveAs(memoryStream);
            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");
            //var model = _service.GetDateExcel(fromDate, toDate);

            //using (var excelPackege = new ExcelPackage())
            //{
            //    int n = 6;
            //    var sheet = excelPackege.Workbook.Worksheets.Add("Data");
            //    ExcelRange excelRange = sheet.Cells["A1:F1"];
            //    excelRange.Merge = true;
            //    excelRange.Value = "Акты сверки мерчантов";
            //    excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    ExcelRange excelDate = sheet.Cells["A2:F2"];
            //    excelDate.Merge = true;
            //    excelDate.Value = $"за период с {fromDate.ToString("yyyy:MM:dd")} по {toDate.ToString("yyyy:MM:dd")}";

            //    excelDate.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    sheet.Cells[5, 1].Value = "#";
            //    sheet.Cells[5, 2].Value = "Номер заказа";
            //    sheet.Cells[5, 3].Value = "Нименование мерчанта";
            //    sheet.Cells[5, 4].Value = "Сумма";
            //    sheet.Cells[5, 5].Value = "Дата";
            //    sheet.Cells[5, 6].Value = "  ПАН-карты  ";

            //    sheet.Cells["A5:F5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    sheet.Cells["A5:F5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCE8FF"));

            //    sheet.Cells["A:C"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    sheet.Cells["E:F"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    sheet.Cells["A1:F1"].Style.Font.Bold = true;
            //    sheet.Cells[5, 1, model.Count + 6, 6].Style.Border.BorderAround(ExcelBorderStyle.Double);
            //    sheet.Cells[5, 1, model.Count + 6, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //     foreach (var item in model)
            //    {
            //        sheet.Cells[n, 1].Value = item.RowNumber;
            //        sheet.Cells[n, 2].Value = item.OrderNumber;
            //        sheet.Cells[n, 3].Value = item.MerchantName;
            //        sheet.Cells[n, 4].Value = item.Sum.ToString();
            //        sheet.Cells[n, 5].Value = item.Date;
            //        sheet.Cells[n, 6].Value = item.panCard;

            //        n++;
            //    }

            //    sheet.Cells[$"D{n}"].Formula = $"=count(D6:D46)";
            //    sheet.Cells[$"A{n}:F{n}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    sheet.Cells[$"A{n}:F{n}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCE8FF"));
            //    sheet.Cells[$"D1:D{n}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            //    sheet.Cells[$"A5:F5"].AutoFitColumns();
            //    sheet.Cells[$"D6:E6"].AutoFitColumns();

            //    var memoryStream = new MemoryStream();
            //    excelPackege.SaveAs(memoryStream);
            //    return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelForm.xlsx");
            //}
        }

    }   
}

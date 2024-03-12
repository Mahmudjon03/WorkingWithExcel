using DocumentFormat.OpenXml.EMMA;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using WorkWithExcel.Services;

namespace WorkWithExcel
{
    public class ExcelImport
    {

        private readonly DataForExcel _service;

        public ExcelImport()
        {
            _service = new DataForExcel();
        }
        public ExcelPackage ExcelService(DateTime fromDate, DateTime toDate)
        {
            var model = _service.GetDateExcel(fromDate, toDate);
            var excelPackege = new ExcelPackage();

            int n = 6;
            var sheet = excelPackege.Workbook.Worksheets.Add("Data");
            ExcelRange excelRange = sheet.Cells["A1:F1"];
            excelRange.Merge = true;
            excelRange.Value = "Акты сверки мерчантов";
            excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange excelDate = sheet.Cells["A2:F2"];
            excelDate.Merge = true;
            excelDate.Value = $"за период с {fromDate.ToString("yyyy:MM:dd")} по {toDate.ToString("yyyy:MM:dd")}";

            excelDate.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[5, 1].Value = "#";
            sheet.Cells[5, 2].Value = "Номер заказа";
            sheet.Cells[5, 3].Value = "Нименование мерчанта";
            sheet.Cells[5, 4].Value = "Сумма";
            sheet.Cells[5, 5].Value = "Дата";
            sheet.Cells[5, 6].Value = "  ПАН - карты  ";

            sheet.Cells["A5:F5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A5:F5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCE8FF"));

            sheet.Cells["A:C"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["E:F"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["A1:F1"].Style.Font.Bold = true;
            sheet.Cells[5, 1, model.Count + 6, 6].Style.Border.BorderAround(ExcelBorderStyle.Double);
            sheet.Cells[5, 1, model.Count + 6, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            foreach (var item in model)
            {
                sheet.Cells[n, 1].Value = item.RowNumber;
                sheet.Cells[n, 2].Value = item.OrderNumber;
                sheet.Cells[n, 3].Value = item.MerchantName;
                sheet.Cells[n, 4].Value = item.Sum;
                sheet.Cells[n, 5].Value = item.Date;
                sheet.Cells[n, 6].Value = item.panCard;

                n++;
            }

            sheet.Cells[$"D{n}"].Formula = $"=SUM(D6:D{n - 1})";
            sheet.Cells[$"A{n}:F{n}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[$"A{n}:F{n}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCE8FF"));
            sheet.Cells[$"D1:D{n}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells[$"A5:F5"].AutoFitColumns();
            sheet.Cells[$"E6:F6"].AutoFitColumns(22);

            return excelPackege;
        }
        public ExcelPackage ExcelService2(DateTime fromDate, DateTime toDate)
        {
            var model = _service.GetData(fromDate, toDate);
            var excelPackege = new ExcelPackage();

            int n = 6;

            var sheet = excelPackege.Workbook.Worksheets.Add("Data");

            ExcelRange excelRange = sheet.Cells["A1:F1"];
            excelRange.Merge = true;
            excelRange.Value = "Акты сверки мерчантов";
            excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange excelDate = sheet.Cells["A2:F2"];
            excelDate.Merge = true;
            excelDate.Value = $"за период с {fromDate.ToString("yyyy:MM:dd")} по {toDate.ToString("yyyy:MM:dd")}";
            excelDate.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Cells[5, 1].Value = "Id";
            sheet.Cells[5, 2].Value = "Наименование мерчанта";
            sheet.Cells[5, 3].Value = "Сумма платежей";
            sheet.Cells[5, 4].Value = "Сумма вознаграждения";
            sheet.Cells[5, 5].Value = "Сумма к зачислению";
            sheet.Cells[5, 6].Value = "Количество платежей";

            sheet.Cells["A5:F5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Cells["A5:F5"].Style.Font.Bold = true;

            sheet.Cells["A5:F5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A5:F5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCE8FF"));

            sheet.Cells[5, 1, model.Count + 6, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            sheet.Cells[5, 1, model.Count + 6, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
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
            sheet.Cells[$"C{n}"].Formula = $"=SUM(C6:C{n - 1})";
            sheet.Cells[$"D{n}"].Formula = $"=SUM(D6:D{n - 1})";
            sheet.Cells[$"E{n}"].Formula = $"=SUM(E6:E{n - 1})";
            sheet.Cells[$"F{n}"].Formula = $"=SUM(F6:F{n - 1})";
            sheet.Cells[$"A{n}:F{n}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[$"A{n}:F{n}"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCE8FF"));
            sheet.Cells[$"A5:F5"].AutoFitColumns();


            return excelPackege;


        }
    }
}

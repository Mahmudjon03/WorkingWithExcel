using OfficeOpenXml;
using OfficeOpenXml.Style;


public class Program
{
    private static void Main(string[] args)
    {
        var data = new Data()
        {
            name = new string[] { "ali", "amin", "maga", "firuz","dima"},
            age = new[] { 13, 21, 23, 45, 24},
            addres = new[] { "dushanbe", "kulob", "vahdat", "norak" ,"102"}
        };
        int row = 2;
        var RN = new Random();
        FileInfo file = new($@"{RN.Next(1,100)}Testt.xlsx");
        using (var packege = new ExcelPackage(file))
        {
            var sheet = packege.Workbook.Worksheets.Add("My Sheet");
            sheet.Cells["A1"].Value = "id";
            sheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B1"].Value = "name";
            sheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["C1"].Value = "age";
            sheet.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["D1"].Value = "Address";
            sheet.Cells["D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Cells["A1:C1"].Style.Font.Bold = true;
             sheet.Cells[1, 1, 6, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);
            sheet.Cells[1, 1, 6, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
           


            for (int i = 0;i< data.name.Length; i++)
            {
             sheet.Cells[row, 1].Value = i;
             sheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
             sheet.Cells[row,2].Value = data.name[i];
             sheet.Cells[row, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
             sheet.Cells[row, 3].Value = data.age[i];
             sheet.Cells[row,3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
             sheet.Cells[row, 4].Value = data.addres[i];
             sheet.Cells[row, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                row++;
            }
            packege.Save();
        };

       
    }
}
class Data
{
    public string[] name { get; set; }
    public string[] addres { get; set; }
    public int[] age { get; set; }
}
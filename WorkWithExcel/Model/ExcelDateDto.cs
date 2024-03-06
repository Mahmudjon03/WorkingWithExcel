namespace WorkWithExcel.Model
{
    public class ExcelDateDto
    {
        public int RowNumber { get; set; }
        public string OrderNumber { get; set; }
        public string MerchantName { get; set; }
        public decimal Sum { get; set; }
        public string Date { get; set; }
        public string? panCard { get; set; }
    }
}

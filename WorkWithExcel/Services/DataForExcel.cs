using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System.Data;
using System.Data.SqlClient;
using WorkWithExcel.Model;

namespace WorkWithExcel.Services
{
    public class DataForExcel : DataContext
    {

        public List<ExcelData> GetData(DateTime fromDate, DateTime toDate)
        {
            string sql = "select " +
                "ROW_NUMBER() OVER (ORDER BY m.id) AS RowNumber, " +
                " m.Name as MerchantName " +
                ",ROUND(CAST(sum(o.total_amount)AS NUMERIC(19, 4)), 4) as Amnt" +
                ",ROUND(CAST(sum((1.5*(o.total_amount))/100) AS NUMERIC(19, 4)), 4) as Reward" +
                ",ROUND(CAST(sum((o.total_amount - (1.5*(o.total_amount))/100)) AS NUMERIC(19, 4)), 4) as CreditAmount" +
                ",count(o.id) as Cnt" +
                " from operations o   " +
                "inner join operation_params op " +
                " on o.id = op.operation_id  inner join merchants m" +
                " on m.Id = o.maker" +
                " where o.[status_id] = 2 and " +
                "date between @fromDate and @toDate " +
                "group by m.[Name],m.id " +
                "order by m.[Name]";
            var pr = new SqlParameter[]
            {
                new SqlParameter("@fromDate",fromDate),
                new SqlParameter("@toDate",toDate),
            };
            var result = new List<ExcelData>();

            using (var sqlObject = Exec(sql, pr, TypeReturn.DataTable, TypeCommand.SqlQuery))
            {
                var dataTable = sqlObject.DataTable;
                if (dataTable.Rows.Count != 0)
                {
                    result = (from DataRow dataRow in dataTable.Rows
                              select new ExcelData()
                              {
                                  RowNumber = (int)(long)dataRow["RowNumber"],
                                  MerchantName = dataRow["MerchantName"].ToString() ?? "",
                                  Amnt = (decimal)dataRow["Amnt"],
                                  Reward = (decimal)dataRow["Reward"],
                                  CreditAmount = (decimal)dataRow["CreditAmount"],
                                  Cnt = (decimal)dataRow["CreditAmount"]
                              }).ToList();
                }
            }

            return result;
        }
        public List<ExcelDateDto> GetDateExcel(DateTime fromDate, DateTime toDate)
        {
            string sql = "  select       " +
                " ROW_NUMBER() OVER (ORDER BY operations.id) AS RowNumber,  " +
                " operations.ext_id as 'Номер заказа',  " +
                " m.Name as 'Нименование мерчанта'," +
                " operations.total_amount as 'Сумма'," +
                " operations.ext_date as 'Дата'" +
                " , op.param.value('(/param/card/number/node())[1]', 'varchar(16)') as 'ПАН-карты'" +
                "  from operations" +
                " inner join operation_params op" +
                " on operations.id = op.operation_id" +
                " inner join merchants m" +
                " on m.Id = operations.maker" +
                " where operations.status_id = 1 and" +
                " operations.date between @fromDate and @toDate " +
                " order by operations.date; ";
            var pr = new SqlParameter[]
            {
                new SqlParameter("@fromDate",fromDate),
                new SqlParameter("@toDate",toDate)
            };
            var result = new List<ExcelDateDto>();
            using (var sqlObject = Exec(sql, pr, TypeReturn.DataTable, TypeCommand.SqlQuery))
            {
                var dataTable = sqlObject.DataTable;
                if (dataTable.Rows.Count != 0)
                {
                    result = (from DataRow dataRow in dataTable.Rows
                              select new ExcelDateDto()
                              {
                                  RowNumber = (int)(long)dataRow["RowNumber"],
                                  OrderNumber = dataRow["Номер заказа"].ToString(),
                                  MerchantName = dataRow["Нименование мерчанта"].ToString(),
                                  Sum = (decimal)dataRow["Сумма"],
                                  Date = dataRow["Дата"].ToString(),
                                  panCard = dataRow["ПАН-карты"].ToString()
                              }
                        ).ToList();
                }
            }
            return result;
        }
    }

}

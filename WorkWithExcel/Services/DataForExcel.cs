using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System.Data;
using System.Data.SqlClient;
using WorkWithExcel.Model;

namespace WorkWithExcel.Services
{
    public class DataForExcel:DataContext
    {

        public List<ExcelData> GetData() 
        {
            string sql = "select " +
                "ROW_NUMBER() OVER (ORDER BY m.id) AS RowNumber," +
                "m.Name as MerchantName" +
                ",ROUND(CAST(sum(o.total_amount)AS NUMERIC(19, 4)), 4) as Amnt" +
                ",ROUND(CAST(sum((1.5*(o.total_amount))/100) AS NUMERIC(19, 4)), 4) as Reward" +
                ",ROUND(CAST(sum((o.total_amount - (1.5*(o.total_amount))/100)) AS NUMERIC(19, 4)), 4) as CreditAmount" +
                ",count(o.id) as Cnt" +
                "  from operation o " +
                " inner join operation_param op" +
                "  on o.id = op.oper_id  inner join merchant m" +
                "  on m.Id = 2 or m.id = 1 or m.id = 3 or m.id = 4" +
                "  where o.[status] = 10 and" +
                " o.[date] >= '20200000' and o.[date] >= '2022000'" +
                "  group by m.[Name],m.id" +
                "  order by m.[Name]";
           
            var result = new List<ExcelData>();
            using ( var sqlObject = Exec(sql, new SqlParameter[] { }, TypeReturn.DataTable, TypeCommand.SqlQuery))
            {
                var dataTable = sqlObject.DataTable;
                if (dataTable.Rows.Count != 0)
                {
                    result = (from DataRow dataRow in dataTable.Rows
                              select new ExcelData()
                              {
                                  RowNumber = (int)(long)dataRow["RowNumber"],
                                  merchat = dataRow["MerchantName"].ToString()??"",
                                  amnt = (decimal)dataRow["Amnt"],
                                  reward = (decimal)dataRow["Reward"],
                                  cnt = (decimal)dataRow["CreditAmount"]
                              }).ToList();
                }
            }

            return result;
        }
    }
}

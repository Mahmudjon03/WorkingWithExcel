using System.Data;
using System.Data.SqlClient;

namespace WorkWithExcel
{
    public class DataContext
    {
        private readonly string _connect;
        public DataContext()
        {
            _connect = "server=192.168.0.6;database=PayTest;user id=APayTest;password=PayTest2024;multipleactiveresultsets=true;";
        }      
        protected virtual SqlObjects? Exec(string sql, SqlParameter[] sqlParam, TypeReturn typeReturn, TypeCommand typeCommand)
        {
            SqlConnection sqlConnection = new SqlConnection("server=192.168.0.6;database=PayTest;user id=APayTest;password=PayTest2024;multipleactiveresultsets=true;");
            sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand(sql, sqlConnection);

            if (typeCommand == TypeCommand.SqlQuery)
            {
                sqlCommand.CommandType = CommandType.Text;
            }
            else
            {
                sqlCommand.CommandType = CommandType.StoredProcedure;
            }

            if (sqlParam != null)
            {
                for (int i = 0; i < sqlParam.Length; i++)
                {
                    sqlCommand.Parameters.Add(sqlParam[i]);
                }
            }

            if (typeReturn == TypeReturn.Empty)
            {
                sqlCommand.ExecuteNonQuery();
                return null;
            }

            if (typeReturn == TypeReturn.SqlDataReader)
            {
                return new SqlObjects()
                {
                    Connection = sqlConnection,
                    Command = sqlCommand,
                    Reader = sqlCommand.ExecuteReader()
                };
            }

            if (typeReturn == TypeReturn.DataTable)
            {
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                DataTable dataTable = new DataTable();
                dataTable.Load(sqlDataReader);

                return new SqlObjects()
                {
                    Connection = sqlConnection,
                    Command = sqlCommand,
                    DataTable = dataTable,
                    Reader = sqlDataReader
                };
            }

            if (typeReturn == TypeReturn.DataSet)
            {
                SqlDataAdapter dataAdapter = new SqlDataAdapter(sqlCommand);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);

                return new SqlObjects()
                {
                    Connection = sqlConnection,
                    Command = sqlCommand,
                    DataSet = ds
                };
            }

            throw new ShukrMoliyaException("Unknown error");
        }

        protected virtual ProcedureStatusCode GetProcedureStatusCode(int resultCode)
        {
            switch (resultCode)
            {
                case -1:
                    return ProcedureStatusCode.Success;
                default:
                    return ProcedureStatusCode.Error;
            }
        }

        public class SqlObjects : IDisposable
        {
            private SqlConnection? _connection;
            private SqlCommand? _command;
            private SqlDataReader? _reader;
            private DataTable? _dataTable;
            private DataSet? _dataSet;
            private bool _disposed;

            public DataSet? DataSet
            {
                get { return _dataSet; }
                set { _dataSet = value; }
            }

            public SqlConnection? Connection
            {
                get { return _connection; }
                set { _connection = value; }
            }

            public SqlCommand? Command
            {
                get { return _command; }
                set { _command = value; }
            }

            public SqlDataReader? Reader
            {
                get { return _reader; }
                set { _reader = value; }
            }

            public DataTable? DataTable
            {
                get { return _dataTable; }
                set { _dataTable = value; }
            }

            public SqlObjects()
            {
                _disposed = false;
            }

            protected void Dispose(bool disposing)
            {
                if (_disposed)
                {
                    return;
                }

                if (disposing)
                {
                    _connection?.Dispose();
                    _command?.Dispose();
                    _dataTable?.Dispose();
                    _dataSet?.Dispose();
                }

                _reader?.Close();
                _disposed = true;
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            ~SqlObjects()
            {
                Dispose(false);
            }
        }
    }
    public enum TypeCommand : int
    {
        SqlQuery = 0,
        SqlStoredProcedure = 1
    }
    public enum TypeReturn : int
    {
        DataTable = 0,
        SqlDataReader = 1,
        DataSet = 2,
        Empty = 3
    }
    public enum ProcedureStatusCode : int
    {
        Success = -1,
        Error = 0
    }
}


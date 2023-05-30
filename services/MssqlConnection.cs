using System.Data.SqlClient;
using Serilog;


namespace AlarmPeople.Bcp;

public class MssqlConnection : IDisposable
{
    public string DSN =>
        $"Server={Server};Database={Database};User Id={User};Password={_password};Integrated Security=false;MultipleActiveResultSets=true;";


    public MssqlConnection(string server, string user, string password, string database)
    {
        this.Server = server;
        User = user;
        _password = password;
        Database = database;

        _conn = new SqlConnection(DSN);
    }

    public void Open() => _conn.Open();
    
    public SqlCommand CreateCommand() => _conn.CreateCommand();
    

    public bool HasTable(string tableName)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = $"select 1 from sys.tables where name = '{tableName}'";
        using var reader = cmd.ExecuteReader();
        var res = false; // assume table does not exist
        while(reader.Read())
        {
            res = reader.IsDBNull(0) ? false : reader.GetInt32(0) == 1;
        }
        Log.Information("Table {tbl} does " + (res ? "exist" : "not exist"), tableName);
        return res;
    }

    public void DropTable(string tableName)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = $"if object_id('{tableName}') is not null drop table [{tableName}]";
        Log.Information("Dropping table {tableName} if it exists", tableName);
        cmd.ExecuteNonQuery();
    }

    public void CreateTable(string tableName, BcpController ctrl)
    {
        var colDefs = string.Join(",", ctrl.SqlDef());
        var cmdText = $"create table [{tableName}] ( {colDefs} )";
        Log.Information("Creating table {t}", tableName);
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = cmdText;
        cmd.ExecuteNonQuery();
    }

    public void TruncateTable(string tableName)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = $"truncate table [{tableName}]";
        Log.Information("Truncating table {tableName}", tableName);
        cmd.ExecuteNonQuery();
    }


    public string Server { get; }
    public string User { get; }

    private string _password;

    public string Database { get; }

    private SqlConnection _conn;


    // boolean variable to ensure dispose
    // method executes only once
    private bool disposedValue;

    // Gets called by the below dispose method
    // resource cleaning happens here
    protected virtual void Dispose(bool disposing)
    {
        // check if already disposed
        if (!disposedValue)
        {
            if (disposing)
            {
                // free managed objects here
                // TotalFiles = 0;
                _conn.Close();
            }
            // fileObject = null;

            // set the bool value to true
            disposedValue = true;
        }
    }

    // The consumer object can call
    // the below dispose method
    public void Dispose()
    {
        // Invoke the above virtual
        // dispose(bool disposing) method
        Dispose(disposing: true);

        // Notify the garbage collector
        // about the cleaning event
        GC.SuppressFinalize(this);
    }
}
using System.Data.SqlClient;
using System.Diagnostics;
using CommandLine;
using CommandLine.Text;
using Serilog;

/*
  bcp_excel import [excel-filename] into [database]..[tablename] -S <server> -U <user> -P <password>
  bcp_excel export [database]..[tablename] into [excel-filename] -S <server> -U <user> -P <password>
*/

/*
First impl.
Scan header for cols
create table from header with "varchar(256) null"
*/

namespace AlarmPeople.Bcp;


class Program
{
    const string ProgramHelpText = @"
bcp_excel 0.4.0.0
Copyright (c) 2023-2024 AlarmPeople A/S
USAGE:
Basic export:
  bcp_excel export mydb..mytbl into my_excel.xlsx -U sa -S MSSQL1 -P pw123
  bcp_excel export mydb..mytbl into my_excel.xlsx -U sa -S MSSQL1 -P pw123 --force true

  -S, --server                   SQL Server
  -U, --user                     SQL user login name
  -P, --password                 SQL user login password
  --force [false|true]           Overwrite existing file (default is false)
  --help                         Display this help screen.
  --version                      Display version information.
  Database and Table (pos. 0)    [database]..[tablename]
  Action (pos. 1)                into
  Excel file (pos. 2)            Excel filename

  [tablename] may be prefixed to allow custom exports
    if [tablename] is prefixed with 'SQL:', the query will be executed and the result exported
    if [tablename] is prefixed with 'FILE:', the content of the file be executed and the result exported

 
Basic import:
  bcp_excel import my_excel.xlsx into mydb..mytbl -U sa -S MSSQL1 -P pw123
  bcp_excel import my_excel.xlsx into mydb..mytbl -U sa -S MSSQL1 -P pw123 --create true
  bcp_excel import my_excel.xlsx into mydb..mytbl -U sa -S MSSQL1 -P pw123 --create false --truncate true

  --forcecreate [false|true]     Create table. Default is false.
  --trucate [false|true]         Delete (truncate) existing data in table. Default is false.
  -f, --format                   Column definitions. Format: [datatype]<size>,[datatype]<size> ...
  -F, --firstrow                 First row of import - counting from line 1
  -L, --lastrow                  Last row of import - counting from line 1
  -E, --errors                   Number of import errors to allow before stopping
  -b, --batchsize                Number of rows to commit in a batch
  --colwidth [512]               Default varchar column size (defaults to 512)
  --sheet [1]                    Sheet number in excel. First sheet has number 1
  -S, --server                   SQL Server
  -U, --user                     SQL user login name
  -P, --password                 SQL user login password
  --help                         Display this help screen.
  --version                      Display version information.
  Excel file (pos. 0)            Excel filename
  Action (pos. 1)                into
  Database and Table (pos. 2)    [database]..[tablename]

-------Setting Format----------
-f i,i,i    -  means 3 integer columns
-f s50,u20  -  means 1 string 50 varchar column, 1 unicode 20 nvarchar
-f i,s,u30  -  means 1 int, 1 string 512 varchar and 1 unicode 30 nvarchar

-------Setting Log level ----------
Use one of the following options to set log level (--warn is default):
    --debug
    --verbose
    --info
    --warn
    --error
    --fatal
";
    // run as 
    // dotnet run -- sbnwork..test in test.xlsx -S localhost/sbnms1 -U sa -P sbntests
    static int Main(string[] args)
    {
        args = ScanForLogLevelAndSetLogger(args);
        // Log.Logger = new LoggerConfiguration()
        //     // .MinimumLevel.Verbose() // Change to .MinimumLevel.Debug() if less info is needed
        //     .MinimumLevel.Debug() // Change to .MinimumLevel.Verbose() if more info is needed
        //     .WriteTo.Console()
        //     .CreateLogger();
        // return
        //     Parser.Default
        //     .ParseArguments<ImportOptions, ExportOptions>(args)
        //     .MapResult(
        //         (ImportOptions opts) => RunImportAndReturnExitCode(opts),
        //         (ExportOptions opts) => RunExportAndReturnExitCode(opts),
        //         errors => DisplayHelp(parserResult, errors)
        //     );
        var parserResult = new Parser(c => c.HelpWriter = null).ParseArguments<ImportOptions, ExportOptions>(args);
        return parserResult.MapResult(
                (ImportOptions opts) => RunImportAndReturnExitCode(opts),
                (ExportOptions opts) => RunExportAndReturnExitCode(opts),
                errors => DisplayHelp(parserResult, errors)
            );

    }

    static string[] ScanForLogLevelAndSetLogger(string[] args)
    {
        string[] logLevels = { "--debug", "--verbose", "--info", "--warn", "--error", "--fatal" };
        int logLevelIx = -1;
        var newArgs = new List<string>();
        foreach (var arg in args)
        {
            var ix = Array.FindIndex(logLevels, t => t.Equals(arg, StringComparison.InvariantCultureIgnoreCase));
            if (ix >= 0)
            {
                logLevelIx = ix;
                continue;
            }
            newArgs.Add(arg);
        }

        switch (logLevelIx)
        {
            case 0:
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Debug()
                    .WriteTo.Console()
                    .CreateLogger();
                break;
            case 1:
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Verbose()
                    .WriteTo.Console()
                    .CreateLogger();
                break;
            case 2:
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Information()
                    .WriteTo.Console()
                    .CreateLogger();
                break;
            case 3:
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Warning()
                    .WriteTo.Console()
                    .CreateLogger();
                break;
            case 4:
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Error()
                    .WriteTo.Console()
                    .CreateLogger();
                break;
            case 5:
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Fatal()
                    .WriteTo.Console()
                    .CreateLogger();
                break;
            default:
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Warning()
                    .WriteTo.Console()
                    .CreateLogger();
                break;
        }
        return newArgs.ToArray();
    }

    static int DisplayHelp(ParserResult<object> result, IEnumerable<Error> errs)
    {
        Console.WriteLine(ProgramHelpText);
        return 1;
    }

    static int RunImportAndReturnExitCode(ImportOptions opt)
    {
        Log.Information("Args: {args}", opt.ToJsonString());
        var (ok, msg) = ValidateImportOptions(opt);
        if (!ok)
        {
            Log.Error(msg);
            return 1;
        }

        var sw = new Stopwatch();
        sw.Start();
        using var excel = new ExcelReader(opt);

        var sheet_no = opt.SheetNo ?? 1;
        // fmt contains a string with column definitions
        // it may look like "i,i,i" or "s50,u20" or "i,s,u30"
        var fmt = opt.Format ?? "";
        var ctrl = excel.GetBcpController(sheet_no - 1, fmt);
        foreach (var r in ctrl.Fields)
        {
            if (opt.ColWidth.HasValue)
                r.DbSize = opt.ColWidth.Value;
            Log.Verbose("Got Field Ix: {ix}, Nm: {nm}, Def: {def}", r.ColIx, r.Name, r.Definition);
        }


        using var conn = new MssqlConnection(opt.Server!, opt.User!, opt.Password!, opt.Database!);
        conn.Open();
        if (opt.ForceCreateTable ?? ImportOptions.Default_CreateTable)
        {
            conn.DropTable(opt.Tablename!);
            conn.CreateTable(opt.Tablename!, ctrl);
        }

        if (opt.Truncate ?? ImportOptions.Default_Truncate)
            conn.TruncateTable(opt.Tablename!);

        int goodRows = 0;
        int badRows = 0;
        int errCnt = 0;
        using (var bulkInsert = new SqlBulkCopy(conn.DSN))
        {
            var maxErrors = opt.ImportErrors;
            bulkInsert.DestinationTableName = opt.Tablename;
            var bchSize = opt.BatchSize ?? ImportOptions.Default_BatchSize;
            foreach (var tbl in ctrl.GetData(batchSize: bchSize, firstRow: opt.FirstRowNum, lastRow: opt.LastRowNum))
            {
                try
                {
                    bulkInsert.WriteToServer(tbl);
                    goodRows += tbl.Rows.Count;
                    Log.Information("Inserted {r} rows", tbl.Rows.Count);
                }
                catch (SqlException ex)
                {
                    Log.Error("SQL Error: {msg}", ex.Message);
                    errCnt++;
                    badRows += tbl.Rows.Count;
                    maxErrors--;
                    if (maxErrors <= 0)
                    {
                        Log.Error("Too many errors. Aborting");
                        break;
                    }
                }
            }
        }
        Log.Information("Inserted {good} rows in {x} ms", goodRows, sw.ElapsedMilliseconds);
        if (errCnt > 0)
            Log.Error("Total {err} errors, affecting {rows} rows", errCnt, badRows);
        return 0;
    }

    static int RunExportAndReturnExitCode(ExportOptions opt)
    {
        Log.Verbose("Args: {args}", opt.ToJsonString());

        var (ok, msg) = ValidateExportOptions(opt);
        if (!ok)
        {
            Log.Error(msg);
            return 1;
        }
        var queryText = QueryTextFromTablename(opt.Tablename!);
        var sw = new Stopwatch();
        sw.Start();

        using var builder = new ExcelBuilder(opt);
        using var conn = new MssqlConnection(opt.Server!, opt.User!, opt.Password!, opt.Database!);
        conn.Open();

        using (var command = conn.CreateCommand())
        {
            command.CommandText = queryText;
            using (var reader = command.ExecuteReader())
            {
                int tableIx = 0;
                do
                {
                    Log.Information("==================================================");
                    Log.Information($"READING TABLE #{tableIx++}");
                    Log.Information("==================================================");

                    var tbl = builder.GetTable(reader.GetSchemaTable());
                    tbl.MakeTable();
                    tbl.AddData(reader);
                    tbl.AddIndexes();
                }
                while (reader.NextResult());
            }
        }

        Log.Information("DONE in {x} ms", sw.ElapsedMilliseconds);
        return 0;
    }

    static string QueryTextFromTablename(string tablename)
    {
        // if tablename starts with FILE: then it is a file and function should return content of file
        // if tablename starts with SQL: then it is a sql statement and function should return sql statement
        // otherwise it is a table name and function should return select * from tablename
        if (tablename.StartsWith("FILE:"))
        {
            var filename = tablename.Substring(5);
            if (!File.Exists(filename))
                throw new Exception($"File: {filename} does not exist");
            return File.ReadAllText(filename);
        }
        if (tablename.StartsWith("SQL:"))
        {
            var sql = tablename.Substring(4);
            return sql;
        }
        return $"select * from [{tablename}]";
    }

    static (bool, string) ValidateBaseOptions(BaseOptions opts)
    {
        if (opts.MetaAction != "into")
            return (false, "Only 'into' action allowed");
        if (string.IsNullOrWhiteSpace(opts.Server))
            return (false, "Missing server");
        if (string.IsNullOrWhiteSpace(opts.User))
            return (false, "Missing user");
        if (string.IsNullOrWhiteSpace(opts.Password))
            return (false, "Missing password");
        if (string.IsNullOrWhiteSpace(opts.Tablename))
            return (false, "Missing table name");
        if (string.IsNullOrWhiteSpace(opts.Filename))
            return (false, "Missing filename");

        return (true, "");
    }

    static (bool, string) ValidateImportOptions(ImportOptions opt)
    {
        var (ok, msg) = ValidateBaseOptions(opt);
        if (!File.Exists(opt.Filename!))
            return (false, $"File: {opt.Filename} does not exist");

        return (ok, msg);
    }
    static (bool, string) ValidateExportOptions(ExportOptions opt)
    {
        var (ok, msg) = ValidateBaseOptions(opt);
        if (!opt.Force && File.Exists(opt.Filename!))
            return (false, $"File: {opt.Filename} already exist");
        return (ok, msg);
    }
}
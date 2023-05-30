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
bcp_excel 0.2.0.0
Copyright (c) 2023 AlarmPeople A/S
USAGE:
Basic export:
  bcp_excel export mydb..mytbl into myXls.xlsx -U sbn0 -S SBN1 -P pw123

  -S, --server                   SQL Server
  -U, --user                     SQL user login name
  -P, --password                 SQL user login password
  --help                         Display this help screen.
  --version                      Display version information.
  Database and Table (pos. 0)    [database]..[tablename]
  Action (pos. 1)                into
  Excel file (pos. 2)            Excel filename
  
Basic import:
  bcp_import import myXls.xlsx into mydb..mytbl

  --nocreate                     Don't create table
  --keep                         Keep existing data in table
  -f, --format                   Column definitions. Format: [datatype]<size>,[datatype]<size> ...
  -F, --firstrow                 First row of import - counting from line 1
  -L, --lastrow                  Last row of import - counting from line 1
  -b, --batchsize                Number of rows to commit in a batch
  --sheet                        Sheet number in excel. First sheet has number 1
  -S, --server                   SQL Server
  -U, --user                     SQL user login name
  -P, --password                 SQL user login password
  --help                         Display this help screen.
  --version                      Display version information.
  Excel file (pos. 0)            Excel filename
  Action (pos. 1)                into
  Database and Table (pos. 2)    [database]..[tablename]

-------Setting Format----------
-fi,i,i    -  means 3 integer columns
-fs50,u20  -  means 1 string 50 varchar column, 1 unicode 20 nvarchar
-fi,s,u30  -  means 1 int, 1 string 512 varchar and 1 unicode 30 nvarchar
";
    // run as 
    // dotnet run -- sbnwork..test in test.xlsx -S localhost/sbnms1 -U sa -P sbntests
    static int Main(string[] args)
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug() // Change to .MinimumLevel.Verbose() if more info is needed
            .WriteTo.Console()
            .CreateLogger();
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

    static int DisplayHelp(ParserResult<object> result, IEnumerable<Error> errs)
    {
        // HelpText helpText;
        // if (errs.IsVersion())  //check if error is version request
        //     helpText = HelpText.AutoBuild(result);
        // else
        // {
        //     helpText = HelpText.AutoBuild(result, h =>
        //     {
        //         //configure help
        //         h.AdditionalNewLineAfterOption = false;
        //         h.Heading = "bcp_excel 0.2.0.0"; //change header
        //         h.Copyright = "Copyright (c) 2023 AlarmPeople A/S"; //change copyright text
        //         h.AddPostOptionsLine(PostOptionsHelpText());
        //         return h; //HelpText.DefaultParsingErrorsHandler(result, h);
        //     }, e => e);
        // }
        Console.WriteLine(ProgramHelpText);
        return 1;
    }

    // static void PrintHelp(ParserResult<ImportOptions> parserResult)
    // {
    //     Console.WriteLine("SPORSGAARD!!!2");
    //     var helpText = GetHelp<ImportOptions>(parserResult);
    //     Console.WriteLine(helpText);
    // }

    // static string PostOptionsHelpText()
    // {
    //     var header = "-------Setting Format----------";

    //     return $"{header}\n" + 
    //             "-fiii    -  means 3 integer columns\n" +
    //             "-fs50u20 -  means 1 string 50 char column, 1 unicode 20 char\n" +
    //             "-fisu30  -  means 1 int, 1 string 512 char and 1 unicode 30 char";
    // }

    //Generate Help text
    // static string GetHelp<T>(ParserResult<T> result)
    // {
    //     Console.WriteLine("SPORSGAARD!!!3");
    //     // use default configuration
    //     // you can customize HelpText and pass different configurations
    //     //see wiki
    //     // https://github.com/commandlineparser/commandline/wiki/How-To#q1
    //     // https://github.com/commandlineparser/commandline/wiki/HelpText-Configuration
    //     return HelpText
    //         .AutoBuild(result,
    //             h =>
    //             {
    //                 h.AddPostOptionsLine(PostOptionsHelpText());
    //                 return h;
    //             }, e => e);
    // }
    static int RunImportAndReturnExitCode(ImportOptions opt)
    {
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
        var fmt = opt.Format ?? "";
        var ctrl = excel.GetBcpController(sheet_no - 1, fmt);
        foreach (var r in ctrl.Fields)
        {
            Log.Verbose("Got Field Ix: {ix}, Nm: {nm}, Def: {def}", r.ColIx, r.Name, r.Definition);
        }

        using var conn = new MssqlConnection(opt.Server!, opt.User!, opt.Password!, opt.Database!);
        conn.Open();
        if (opt.CreateTable ?? ImportOptions.Default_CreateTable)
        {
            conn.DropTable(opt.Tablename!);
            conn.CreateTable(opt.Tablename!, ctrl);
        }

        if (opt.Truncate ?? ImportOptions.Default_Truncate)
            conn.TruncateTable(opt.Tablename!);

        using (var bulkInsert = new SqlBulkCopy(conn.DSN))
        {
            bulkInsert.DestinationTableName = opt.Tablename;
            var bchSize = opt.BatchSize ?? ImportOptions.Default_BatchSize;
            foreach (var tbl in ctrl.GetData(batchSize: bchSize))
            {
                bulkInsert.WriteToServer(tbl);
                Log.Warning("Inserted {r} rows", tbl.Rows.Count);
            }
        }
        Log.Information("DONE in {x} ms", sw.ElapsedMilliseconds);
        return 0;
    }

    static int RunExportAndReturnExitCode(ExportOptions opt)
    {
        var (ok, msg) = ValidateExportOptions(opt);
        if (!ok)
        {
            Log.Error(msg);
            return 1;
        }
        var queryText = $"select * from [{opt.Tablename}]";
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
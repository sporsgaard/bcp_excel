using CommandLine;
using CommandLine.Text;
using System.Text.Json;

namespace AlarmPeople.Bcp;

public abstract class BaseOptions
{
    public void SetDatabaseAndTable(string value)
    {
        var v1 = value.Split("..");
        if (v1.Length != 2)
            throw new ArgumentException("Argument must be [database]..[tablename]");
        Database = v1[0];
        Tablename = v1[1];
    }
    public string GetDatabaseAndTable() => $"{Database}..{Tablename}";

    public abstract string DbTable { get; set; }
    public abstract string MetaAction { get; set; }
    public abstract string? Filename { get; set; }

    [Option('S', "server", HelpText = "SQL Server")]
    public string? Server { get; set; }

    [Option('U', "user", HelpText = "SQL user login name")]
    public string? User { get; set; }

    [Option('P', "password", HelpText = "SQL user login password")]
    public string? Password { get; set; }

    public string? Tablename { get; set; }
    public string? Database { get; set; }


    // ===========================================================================
    //    IMPORT EXCEL INTO TABLE OPTIONS
    // ===========================================================================

    [Option("forcecreate", HelpText = "Force Create table")]
    public bool? ForceCreateTable { get; set; } = false;

    [Option("truncate", HelpText = "Delete (truncate) existing data in table")]
    public bool? Truncate { get; set; } = false;

    [Option("colwidth", HelpText = "Default varchar column size")]
    public int? ColWidth { get; set; } = null;

    [Option('f', "format", HelpText = "Column definitions. Format: [datatype]<size> [datatype]<size> ...")]
    public string? Format { get; set; }

    [Option('F', "firstrow", HelpText = "First row of import - counting from line 1")]
    public int FirstRowNum { get; set; } = -1;

    [Option('L', "lastrow", HelpText = "Last row of import - counting from line 1")]
    public int LastRowNum { get; set; } = -1;

    [Option('b', "batchsize", HelpText = "Number of rows to commit in a batch")]
    public int? BatchSize { get; set; }

    [Option("sheet", HelpText = "Sheet number in excel. First sheet has number 1")]
    public int? SheetNo { get; set; }

    [Option('E', "errors", HelpText = "Number of import errors to allow before stopping")]
    public int ImportErrors { get; set; } = 3;


    // ===========================================================================
    //    EXPORT TABLE INTO EXCEL OPTIONS
    // ===========================================================================
    [Option('f', "force", HelpText = "Force overwrite existing Excel file")]
    public bool Force { get; set; }


    // ===========================================================================
    //    Easy debugging
    // ===========================================================================
    public string ToJsonString() => JsonSerializer.Serialize(this);
}


[Verb("import", isDefault: true)]
public class ImportOptions : BaseOptions
{
    public const bool Default_CreateTable = false;
    public const bool Default_Truncate = false;
    public const int Default_BatchSize = 1000;
    // public readonly string Default_DbDataType = "varchar(512) null";
    public string Default_DbDataType { get { return $"varchar({ColWidth ?? 0}) null"; } }


    [Value(2, MetaName = "Database and Table", HelpText = "[database]..[tablename]")]
    public override string DbTable { get => GetDatabaseAndTable(); set { SetDatabaseAndTable(value); } }

    [Value(1, MetaName = "Action", HelpText = "into")]
    public override string MetaAction { get; set; } = ""; // must be "into"

    [Value(0, MetaName = "Excel file", HelpText = "Excel filename")]
    public override string? Filename { get; set; }

    [Usage(ApplicationAlias = "bcp_export")]
    public static IEnumerable<Example> Examples
    {
        get
        {
            return new List<Example>()
            {
                new Example("Basic import", new ImportOptions { DbTable = "mydb..mytbl", MetaAction = "into", Filename = "myXls.xlsx" })
            };
        }
    }
}

[Verb("export")]
public class ExportOptions : BaseOptions
{
    [Value(0, Required = true, MetaName = "Database and Table", HelpText = "[database]..[tablename]")]
    public override string DbTable { get => GetDatabaseAndTable(); set { SetDatabaseAndTable(value); } }

    [Value(1, Required = true, MetaName = "Action", HelpText = "into")]
    public override string MetaAction { get; set; } = ""; // must be "into"

    [Value(2, Required = true, MetaName = "Excel file", HelpText = "Excel filename")]
    public override string? Filename { get; set; } = "";

}
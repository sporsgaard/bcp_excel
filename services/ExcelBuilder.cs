using Microsoft.Extensions.Logging;
using System.Data;
using System.Data.Common;
using Serilog;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace AlarmPeople.Bcp;

public class ExcelBuilder: IDisposable
{
    private ExportOptions _options;
    private XSSFWorkbook _book;
    private int _tableIndex = 1;
    // private SqliteBuilderTable? lastTable;
    public string? NextTableName { get; private set; } = null;
    private int debugLevel = 0;
    public ExcelBuilder(ExportOptions options, int debug = 0)
    {
        _options = options;
        debugLevel = debug;
        try
        {
            _book = new XSSFWorkbook(XSSFWorkbookType.XLSX);
        }
        catch (Exception e)
        {
            throw new Exception($"{e.Message}\nProbably open in Excel");
        }
    }

    public ISheet MakeSheet(string name) => _book.CreateSheet(name);

    public ExcelBuilderTable GetTable(DataTable schema)
    {
        var firstColName = schema.Rows[0][SchemaTableColumn.BaseColumnName].ToString() ?? "";
        // Log.Debug($"First column name: {firstColName}");

        ExcelBuilderTable tbl;
        if (firstColName.ToLower() == "__meta__")
        {
            tbl = new ExcelBuilderTable($"TABLE_{_tableIndex++}", this);
            tbl.OnColumnValue += MetaCallback;
            tbl.ToSheet = debugLevel > 2;
        }
        else
        if (firstColName.StartsWith("_"))
        {
            tbl = new ExcelBuilderTable($"_HIDDEN_TABLE_{_tableIndex++}", this);
            tbl.ToSheet = debugLevel > 2;
        }
        else
        {
            tbl = new ExcelBuilderTable(NextTableName ?? $"TABLE_{_tableIndex++}", this);
            tbl.ToSheet = true;
        }
        tbl.SetColumns(schema);
        return tbl;
    }

    private void MetaCallback(int columnIndex, string colName, string colValue)
    {
        // Log.Information("Callback: {ix} - {nm} - {val}", columnIndex, colName, colValue);
        var cn = colName.ToLower().Trim();
        if (cn == "name")
            NextTableName = colValue;
    }

    private void SaveBook()
    {
        if (!string.IsNullOrWhiteSpace(_options.Filename))
        {
            if (File.Exists(_options.Filename))
            {
                File.Delete(_options.Filename);
            }
            // if _options.Filename does not have an extension, add .xlsx
            if (Path.GetExtension(_options.Filename) == "")
                _options.Filename += ".xlsx";
            using var strm = new FileStream(_options.Filename, FileMode.CreateNew );
            _book.Write(strm, false);
        }
    }

    private bool disposedValue;
    protected virtual void Dispose(bool disposing)
    {
        // check if already disposed
        if (!disposedValue)
        {
            if (disposing)
            {
                SaveBook();
                _book.Close();
            }
            disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

}

public class ExcelBuilderTable
{
    public string Name { get; set; }
    public event Action<int, string, string>? OnColumnValue;
    public bool ToSheet { get; set; }
    private List<IBuilderColumn> _columns = new List<IBuilderColumn>();
    private ExcelBuilder _builder;

    private ISheet? _currentSheet;
    public ExcelBuilderTable(string name, ExcelBuilder builder)
    {
        Name = name;
        _builder = builder;
    }

    public ExcelBuilderTable SetColumns(DataTable schema)
    {
        // Log.Information("Table {nm} - Key: {key} - Unique: {uq}", Name, _builder.NextKey, _builder.NextUnique);
        var ix = 0;
        List<string> cols = new List<string>();
        foreach (DataRow column in schema.Rows)
        {
            ix++;
            var c = SqliteType(column, ix);
            if (c == null)
            {
                Log.Error("NOT SUPPORTED: Table: {tbl}, ColIx: {colIx}, ColName: {col}, Datatype: {dt}",
                    Name, ix, column[SchemaTableColumn.BaseColumnName].ToString(), column[SchemaTableColumn.DataType].ToString());
            }
            else
            {
                if (c.ColumnName.Trim() == "")
                    c.ColumnName = $"col_{ix}";
                while (cols.Contains(c.ColumnName.ToLower()))
                    c.ColumnName += 'x';

                _columns.Add(c);
                cols.Add(c.ColumnName.ToLower());
            }
        }
        return this;
    }

    public ExcelBuilderTable MakeTable()
    {
        if (!ToSheet)
            return this;

        _currentSheet = _builder.MakeSheet(Name);
        IRow headerRow = _currentSheet.CreateRow(0);
        
        foreach (var c in _columns.Where(c => c != null))
        {
            var cell = headerRow.CreateCell(c.ColumnIndex-1);
            cell.SetCellType(CellType.String);
            cell.SetCellValue(c.ColumnName);
        }
        return this;
    }

    public ExcelBuilderTable AddIndexes()
    {
        return this;
    }

    private void DoOnColumnValue(int index, string colName, string colValue)
    {
        OnColumnValue?.Invoke(index, colName, colValue);
    }

    public ExcelBuilderTable AddData(DbDataReader reader) => ToSheet ? AddDataToSqlite(reader) : AddDataNoSqlite(reader);

    public ExcelBuilderTable AddDataToSqlite(DbDataReader reader)
    {
        if (_currentSheet == null)
            return AddDataNoSqlite(reader);

        int rowIx = 1;

        while (reader.Read())
        {
            //var cmd = _builder.GetSqliteCommand(stmt, tran);
            IRow row = _currentSheet.CreateRow(rowIx++);
            foreach (var c in _columns.Where(c => c != null))
            {
                var cell = row.CreateCell(c.ColumnIndex - 1);
                c!.SetCellData(cell, reader);
                if (OnColumnValue != null)
                    DoOnColumnValue(c.ColumnIndex, c.ColumnName, c!.GetDataAsString(reader));
            }
        }
        return this;
    }

    public ExcelBuilderTable AddDataNoSqlite(DbDataReader reader)
    {
        while (reader.Read())
        {
            foreach (var c in _columns.Where(c => c != null))
            {
                if (OnColumnValue != null)
                    DoOnColumnValue(c!.ColumnIndex, c!.ColumnName, c!.GetDataAsString(reader));
            }
        }
        return this;
    }

    static IBuilderColumn? SqliteType(DataRow column, int columnIndex)
        => column[SchemaTableColumn.DataType].ToString()
        switch
        {
            "System.String" => new BuilderColumnText(column, columnIndex),
            "System.Int32" => new BuilderColumnInt32(column, columnIndex),
            "System.Int16" => new BuilderColumnInt16(column, columnIndex),
            _ => null
        };
}

public interface IBuilderColumn
{
    int ColumnIndex { get; }
    string OriginalColumnName { get; }
    string ColumnName { get; set; }
    string AliasName { get; }
    bool IsKey { get; set; }
    bool IsUnique { get; set; }
    string DbType { get; }
    void SetCellData(ICell cell, DbDataReader reader);
    string GetDataAsString(DbDataReader reader);
}

public abstract class BuilderColumnBase : IBuilderColumn
{
    public int ColumnIndex { get; }
    public string OriginalColumnName { get; }
    public string ColumnName { get; set; }
    public bool IsKey { get; set; } = false;
    public bool IsUnique { get; set; } = false;
    public string DbType { get; protected set; }
    public bool AllowNull { get; }
    public string AllowNullText { get => AllowNull ? "NULL" : "NOT NULL"; }
    public string AliasName { get => $"@v{ColumnIndex}"; }

    public abstract void SetCellData(ICell cell, DbDataReader reader);


    public BuilderColumnBase(DataRow column, int columnIndex, string dbType)
    {
        var colName = column[SchemaTableColumn.BaseColumnName].ToString();
        OriginalColumnName = colName ?? "";
        ColumnName = String.IsNullOrWhiteSpace(colName) ? $"COLUMN_{columnIndex}" : colName;
        AllowNull = (column[SchemaTableColumn.AllowDBNull] as bool?) ?? true;
        DbType = dbType;
        ColumnIndex = columnIndex;
    }

    public override string ToString()
    {
        string constraint = IsKey ? "PRIMARY KEY" : IsUnique ? "UNIQUE" : "";
        return $"[{ColumnName}] {DbType} {AllowNullText} {constraint}".TrimEnd();
    }

    public void SetData(DbDataReader reader)
    {
        if (reader.IsDBNull(ColumnIndex - 1))
        {
            // cmd.Parameters.AddWithValue(AliasName, DBNull.Value);
        }
        else
        {
            //SetDataValue(cmd, reader);
        }
    }
    public string GetDataAsString(DbDataReader reader) => reader.IsDBNull(ColumnIndex - 1) ? "" : GetDataValueAsString(reader);

    // protected abstract void SetDataValue(SqliteCommand cmd, DbDataReader reader);
    protected abstract string GetDataValueAsString(DbDataReader reader);
}

public class BuilderColumnText : BuilderColumnBase, IBuilderColumn
{
    public BuilderColumnText(DataRow column, int columnIndex)
        : base(column, columnIndex, "TEXT") { }
    // protected override void SetDataValue(SqliteCommand cmd, DbDataReader reader) => cmd.Parameters.AddWithValue(AliasName, GetDataValueAsString(reader));
    protected override string GetDataValueAsString(DbDataReader reader) => reader.GetString(ColumnIndex - 1).TrimEnd();

    public override void SetCellData(ICell cell, DbDataReader reader)
    {
        cell.SetCellType(CellType.String);
        if (!reader.IsDBNull(ColumnIndex - 1))
            cell.SetCellValue(GetDataAsString(reader));
    }
}

public class BuilderColumnInt32 : BuilderColumnBase, IBuilderColumn
{
    public BuilderColumnInt32(DataRow column, int columnIndex)
        : base(column, columnIndex, "INT") { }

    //protected override void SetDataValue(SqliteCommand cmd, DbDataReader reader) => cmd.Parameters.AddWithValue(AliasName, reader.GetInt32(ColumnIndex - 1));
    protected override string GetDataValueAsString(DbDataReader reader) => reader.GetInt32(ColumnIndex - 1).ToString();
    public override void SetCellData(ICell cell, DbDataReader reader)
    {
        cell.SetCellType(CellType.Numeric);
        if (!reader.IsDBNull(ColumnIndex - 1))
            cell.SetCellValue(reader.GetInt32(ColumnIndex - 1));
    }

}

public class BuilderColumnInt16 : BuilderColumnBase, IBuilderColumn
{
    public BuilderColumnInt16(DataRow column, int columnIndex)
        : base(column, columnIndex, "INT") { }
    //protected override void SetDataValue(SqliteCommand cmd, DbDataReader reader) => cmd.Parameters.AddWithValue(AliasName, reader.GetInt16(ColumnIndex - 1));
    protected override string GetDataValueAsString(DbDataReader reader) => reader.GetInt16(ColumnIndex - 1).ToString();
    public override void SetCellData(ICell cell, DbDataReader reader)
    {
        cell.SetCellType(CellType.Numeric);
        if (!reader.IsDBNull(ColumnIndex - 1))
            cell.SetCellValue(reader.GetInt16(ColumnIndex - 1));
    }

}

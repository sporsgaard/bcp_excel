using System.Data;
using NPOI.SS.UserModel;
using Serilog;

namespace AlarmPeople.Bcp;

public class BcpController
{

    public BcpController(ISheet sheet, int controlCol, int nameRow, int definitionRow, List<(string, int)> fmts, string defaultDbDataType)
    {
        _sheet = sheet;
        ControlCol = controlCol;
        NameRow = nameRow < 0 ? 0 : nameRow;
        DefinitionRow = definitionRow;
        _fmts = fmts;
        _defaultDbDataType = defaultDbDataType;
        Fields = GetFields().ToList();
    }

    public delegate Field AcceptDefinition(int columnIndex, string name, string def);
    private static List<AcceptDefinition> FieldFactory = new List<AcceptDefinition>();
    public static void AddFieldFactory(AcceptDefinition factory) =>
        FieldFactory.Add(factory);
    public static Field GetField(int columnIndex, string name, string def)
    {
        foreach (var fact in FieldFactory)
        {
            var v = fact(columnIndex, name, def);
            if (v != null)
                return v;
        }
        throw new Exception("No matching field types");
    }

    ISheet _sheet;
    public int ControlCol { get; set; }
    public int NameRow { get; set; }
    public int DefinitionRow { get; set; }
    public List<Field> Fields { get; }
    private List<(string, int)> _fmts;
    private String _defaultDbDataType;

    public IEnumerable<string> SqlDef()
    {
        foreach (var r in Fields)
            yield return r.SqlDef;
    }

    private static string? FmtToColDef( (string, int) fmt)
    {
        var sz = fmt.Item2 > 0 ? fmt.Item2 : 512;
        switch (fmt.Item1)
        {
            case "int":
                return "int";
            case "string":
                return $"varchar({sz})";
            case "unicode":
                return $"nvarchar({sz})"; 
            default:
                return null;
        }

    }

    private IEnumerable<Field> GetFields()
    {
        Log.Debug("Getting fields");
        var nm_row = _sheet.GetRow(NameRow);
        var maxIx = nm_row.LastCellNum;
        IRow? def_row = DefinitionRow < 0 ? null : _sheet.GetRow(DefinitionRow);
        (string, int) empty = ("", -1);

        int sqlColIx = 0;
        for (int i = 0; i < maxIx; i++)
        {
            if (i == ControlCol)
                continue;
            var fmt = _fmts.Count > i ? _fmts[i] : empty;
            var nm = nm_row.GetCell(i)?.ToString()?.Trim();
            var def = def_row?.GetCell(i)?.ToString()?.Trim() ?? FmtToColDef(fmt) ?? _defaultDbDataType;

            if (!string.IsNullOrEmpty(nm))
            {
                var f = Field.MakeField(i, sqlColIx, nm, def);
                sqlColIx++;
                Log.Debug($"[{sqlColIx,3}] [{i+1,3}] {nm,-18} {f.Definition()}");
                yield return f;
            }
            else
            {
                Log.Debug($"[   ] [{i+1,3}] <Empty>");
            }
        }
    }

    private DataTable GetDataTable()
    {
        DataTable result = new DataTable();
        foreach (var r in Fields)
        {
            result.Columns.Add(r.GetDataColumn());
        }
        return result;
    }
    public IEnumerable<DataTable> GetData(int batchSize, int firstRow = -1, int lastRow = -1)
    {
        if (firstRow <= 0)
            firstRow = 1 + (NameRow > DefinitionRow ? NameRow : DefinitionRow);
        if (lastRow <= 0 || lastRow > _sheet.LastRowNum)
            lastRow = _sheet.LastRowNum + 1;
        // var firstRow = 1 + (NameRow > DefinitionRow ? NameRow : DefinitionRow);
        DataTable result = GetDataTable();
        int rows_left = batchSize;
        for (int i = firstRow; i < lastRow; i++)
        {
            // Log.Information("Getting data from row {r}", i);
            DataRow dataRow = result.NewRow();
            var excel_row = _sheet.GetRow(i);
            try
            {
                // Log.Information("Min col: {c1} max col {c2}", excel_row.FirstCellNum, excel_row.LastCellNum);
                bool inserted = false;
                foreach (var r in Fields)
                    inserted |= r.SetData(dataRow, _sheet, excel_row);
                if (inserted)
                {
                    result.Rows.Add(dataRow);
                    rows_left--;
                }
                else
                {
                    Log.Information("No data in row {row}", i + 1);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error in data row: {r}", i + 1);
            }

            if (rows_left == 0)
            {
                yield return result;
                result = GetDataTable();
                rows_left = batchSize;
            }
        }
        if (result.Rows.Count > 0)
            yield return result;
    }
}

using System.Data;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using Serilog;

namespace AlarmPeople.Bcp;

public abstract class Field
{
    public string Name { get; set; } = "";
    public string Definition 
        => DbDataType + (DbSize>0? $"({DbSize})" : "") + (DbNullable? " null" : " not null");
    public int ColIx { get; set; }
    public int SqlIx { get; set; } = 0;
    public string SqlDef => $"[{Name}] {Definition}";

    public string DbDataType { get; set; } = "";
    public int DbSize { get; set; } = -1;
    public bool DbNullable { get; set; } = true;

    public DataColumn GetDataColumn()
    {
        DataColumn res = new DataColumn();
        res.ColumnName = Name;
        res.DataType = InternGetDataType();
        res.AllowDBNull = DbNullable;
        return res;
    }

    protected abstract Type? InternGetDataType();
    protected abstract void InternSetEmptyField(DataRow dataRow);
    protected abstract void InternSetData(DataRow dataRow, ICell cell);
    

    public bool SetData(DataRow dataRow, ISheet sheet, IRow row)
    {
        // Log.Information("Get data for col nm: {n} colix: {c}, sqlix: {s}", Name, ColIx, SqlIx);
        ICell? cell = row.GetCell(ColIx);
        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString()))
        {
            // Log.Information("Empty cell");
            if (DbNullable)
            {
                dataRow[SqlIx] = null;
            }
            else
            {
                InternSetEmptyField(dataRow);
            }
            return false;
        }
        InternSetData(dataRow, cell);
        return true;
    }

    private static Regex dataTypeRegEx = 
        new Regex(@"^(?<type>\w+)\s*(\(\s*(?<size>\d+)\s*\))?\s*(?<not>not\s+)?(?<null>null)?", RegexOptions.IgnoreCase);

    private static Dictionary<string, Func<Field>> FieldTypesMap = new Dictionary<string, Func<Field>>
    {
        {"i", () => new IntegerField() },
        {"int", () => new IntegerField() },
        {"s", () => new StringField() },
        {"string", () => new StringField() },
        {"char", () => new StringField("char")},
        {"varchar", () => new StringField()},
        {"nvarchar", () => new StringField("nvarchar")}
    };

    public static Field MakeField(int colIx, int sqlColIx, string name, string definition)
    {
        // definition is either "S", "I" or generic syntax
        // generic syntax is
        // <type>[(size)] [[not] null]

        var d = definition.ToLower();
        Match m = dataTypeRegEx.Match(d);
        if (!m.Success)
            throw new ArgumentException($"ColIx {colIx}, Name: {name}, Def: {definition} is not a valid SQL type definition");
        
        string tp = m.Groups["type"].Value;
        if (FieldTypesMap.TryGetValue(tp, out var factory))
        {
            Field res = factory();
            res.ColIx = colIx;
            res.Name = name;
            res.SqlIx = sqlColIx;
            string sz = m.Groups["size"].Value;
            res.DbSize = string.IsNullOrWhiteSpace(sz) ? res.DbSize : Int32.Parse(sz);

            res.DbNullable = string.IsNullOrWhiteSpace(m.Groups["not"].Value);
            return res;
        }
        else
        {
            throw new ArgumentException($"ColIx {colIx}, Name: {name}, Def: {definition}, datatype: {tp} is not supported");
        }
    }

}

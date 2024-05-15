using System.Data;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using Serilog;

namespace AlarmPeople.Bcp;

public class StringField : Field
{
    static Type? dataType = System.Type.GetType("System.String");

    public override string Definition()
        => DbDataType + (DbSize>0? $"({DbSize})" : "(512)") + (DbNullable? " null" : " not null");

    protected override Type? InternGetDataType() => dataType;

    protected override void InternSetEmptyField(DataRow dataRow)
        => dataRow[SqlIx] = "";
    protected override void InternSetData(DataRow dataRow, ICell cell)
    {
        string? val = cell.ToString()?.Trim();
        // Log.Information("S.Value: {v}", val);
        if (!string.IsNullOrWhiteSpace(val) && DbSize>0 && val.Length > DbSize)
        {
            var cellRow = cell.RowIndex + 1;
            var cellCol = cell.ColumnIndex + 1;
            Serilog.Log.Warning("Truncating string in Excel (row,col): ({r},{c}) from {x} to {s}", cellRow, cellCol, val.Length, DbSize);
            val = val[..DbSize];
        }
        dataRow[SqlIx] = val;        
    }

    internal StringField(string? dbDataType = null)
    {
        DbDataType = dbDataType ?? "varchar";
    }
}
using System.Data;
using NPOI.SS.UserModel;

namespace AlarmPeople.Bcp;

public class StringField : Field
{
    static Type? dataType = System.Type.GetType("System.String");
 
    protected override Type? InternGetDataType() => dataType;

    protected override void InternSetEmptyField(DataRow dataRow)
        => dataRow[SqlIx] = "";
    protected override void InternSetData(DataRow dataRow, ICell cell)
    {
        string? val = cell.ToString()?.Trim();
        // Log.Information("S.Value: {v}", val);
        dataRow[SqlIx] = val;
    }

    internal StringField(string? dbDataType = null)
    {
        this.DbDataType = dbDataType ?? "varchar";
        this.DbSize = 512;
    }
}
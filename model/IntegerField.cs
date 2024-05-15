using System.Data;
using NPOI.SS.UserModel;

namespace AlarmPeople.Bcp;

public class IntegerField: Field
{
    static Type? dataType = System.Type.GetType("System.Int32");
    public override string Definition()
        => DbDataType + (DbNullable ? " null" : " not null");

    protected override Type? InternGetDataType() => dataType;

    protected override void InternSetEmptyField(DataRow dataRow)
        => dataRow[SqlIx] = 0;

    protected override void InternSetData(DataRow dataRow, ICell cell)
        => dataRow[SqlIx] = (int)cell.NumericCellValue;

    internal IntegerField()
    {
        this.DbDataType = "int";
    }
}
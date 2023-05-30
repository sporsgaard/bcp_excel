using System.Data;
using System.Text.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace AlarmPeople.Bcp;


public class ExcelReader : IDisposable
{
    private XSSFWorkbook book;
    private ImportOptions _options;
    public ExcelReader(ImportOptions options)
    {
        _options = options;
        try
        {
            book = new XSSFWorkbook(_options.Filename);
        }
        catch (Exception e)
        {
            throw new Exception($"{e.Message}\nProbably open in Excel");
        }
    }


    public BcpController GetBcpController(int sheetNo, string fmt)
    {
        var sheet = book.GetSheetAt(sheetNo);
        var fmts = ParseFormatOptions.Parse(fmt);
        // search for _DEF column
        int ctrl_col = FindCellInRow(sheet, 0, "_CONTROL");
        if (ctrl_col<0)
            return new BcpController(sheet, -1, -1, -1, fmts);
        int nm_row_ix = FindCellInCol(sheet, ctrl_col, "COLNAME");
        int def_row_ix = FindCellInCol(sheet, ctrl_col, "DEF");
        return new BcpController(sheet, ctrl_col, nm_row_ix, def_row_ix, fmts);
    }


    private int FindCellInRow(ISheet sheet, int rowNo, string search_str)
    {
        IRow headerRow = sheet.GetRow(rowNo);
        for (int j = 0; j < headerRow.LastCellNum; j++)
        {
            ICell? cell = headerRow.GetCell(j);
            if (cell?.ToString()?.Trim() == search_str)
            {
                return j;
            }
        }
        return -1;
    }


    private int FindCellInCol(ISheet sheet, int colNo, string search_str)
    {
        //IRow headerRow = sheet.Get(rowNo);
        for (int i = sheet.FirstRowNum; i < sheet.LastRowNum; i++)
        {
            var r = sheet.GetRow(i);
            if (r==null)
                continue;

            ICell? cell = r.GetCell(colNo);
            if (cell?.ToString()?.Trim() == search_str)
            {
                return i;
            }
        }
        return -1;
    }




    static void ReadExcel(string input_filename, Stream output_stream, int? sheet_no = null)
    {
        DataTable dtTable = new DataTable();
        List<string> rowList = new List<string>();
        ISheet sheet;
        HashSet<int> ignored_columns = new HashSet<int>();
        using (var stream = new FileStream(input_filename, FileMode.Open))
        {
            stream.Position = 0;
            XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
            sheet = xssWorkbook.GetSheetAt(sheet_no ?? 0);
            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString()))
                {
                    ignored_columns.Add(j);
                }
                else
                {
                    dtTable.Columns.Add(cell.ToString());
                }
            }
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                if (row.Cells.All(d => d.CellType == CellType.Blank))
                {
                    continue;
                }
                for (int j = 0; j < cellCount; j++)
                {
                    if (ignored_columns.Contains(j))
                    {
                        continue; // dont get values for ignored columns.
                    }
                    var c = row.GetCell(j)?.ToString();
                    rowList.Add(string.IsNullOrEmpty(c) ? "NULL" : c.Trim());
                }
                if (rowList.Count > 0)
                {
                    dtTable.Rows.Add(rowList.ToArray());
                }
                rowList.Clear();
            }
        }
        var res = DataTable_to_Json(dtTable);
        var buf = System.Text.Encoding.UTF8.GetBytes(res);
        if (buf != null)
        {
            output_stream.Write(buf, 0, buf.Length);
        }
    }

    public static string DataTable_to_Json(DataTable dataTable)
    {
        if (dataTable == null)
        {
            return string.Empty;
        }

        var data = dataTable.Rows.OfType<DataRow>()
                    .Select(row => dataTable.Columns.OfType<DataColumn>()
                        .ToDictionary(col => col.ColumnName, c => row[c]));

        return JsonSerializer.Serialize(data);
    }



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
                book.Close();
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
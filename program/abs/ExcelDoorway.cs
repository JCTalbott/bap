using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

// class to communicate effectively with Excel
// inherits IDisposable in order to "use" the file properly (not needed, but it's good for cleanup)
public class ExcelDoorway : IDisposable {
    public bool IsDisposed { get; private set; } = false;
    public string FilePath { get; }
    public SpreadsheetDocument spreadsheetDoc { get; }
    public Stylesheet Stylesheet { get; }
    public readonly WorkbookPart WorkbookPart;
    public Workbook Workbook;
    public Worksheet Worksheet;

    public ExcelDoorway(string filePath, bool isEditable) {
        this.FilePath = filePath;
        spreadsheetDoc = SpreadsheetDocument.Open(filePath, isEditable);
        WorkbookPart = spreadsheetDoc.WorkbookPart;
        Stylesheet = WorkbookPart.WorkbookStylesPart.Stylesheet
            ?? throw new InvalidOperationException("WorkbookPart could not be initialized.");
        Workbook = WorkbookPart.Workbook;
    }

    public Sheets getSheets() { return Workbook.Sheets; }
    public (WorksheetPart, SheetData) getSheetData(StringValue id) { 
        WorksheetPart part = (WorksheetPart) WorkbookPart.GetPartById(id);
        return (part, part.Worksheet.Elements<SheetData>().FirstOrDefault());
    }
    
    // Enhanced methods for better data extraction
    public Dictionary<string, string> ExtractSheetData(string sheetName) {
        var result = new Dictionary<string, string>();
        var sheet = Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
        if (sheet == null) return result;
        
        var (worksheetPart, sheetData) = getSheetData(sheet.Id);
        if (sheetData == null) return result;
        
        foreach (var row in sheetData.Elements<Row>()) {
            foreach (var cell in row.Elements<Cell>()) {
                if (cell.CellReference != null && !string.IsNullOrEmpty(cell.InnerText)) {
                    result[cell.CellReference] = GetCellValue(cell: cell);
                }
            }
        }
        return result;
    }
    
    public List<Dictionary<string, string>> ExtractTableData(string sheetName, int headerRow = 1) {
        var result = new List<Dictionary<string, string>>();
        var sheet = Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
        if (sheet == null) return result;
        
        var (worksheetPart, sheetData) = getSheetData(sheet.Id);
        if (sheetData == null) return result;
        
        var rows = sheetData.Elements<Row>().ToList();
        if (rows.Count <= headerRow) return result;
        
        var headerRowData = rows[headerRow - 1];
        var headers = new List<string>();
        
        foreach (var cell in headerRowData.Elements<Cell>()) {
            if (cell.CellReference != null) {
                headers.Add(GetCellValue(cell: cell));
            }
        }
        
        for (int i = headerRow; i < rows.Count; i++) {
            var rowData = new Dictionary<string, string>();
            var cells = rows[i].Elements<Cell>().ToList();
            
            for (int j = 0; j < Math.Min(headers.Count, cells.Count); j++) {
                if (j < headers.Count && cells[j] != null) {
                    rowData[headers[j]] = GetCellValue(cell: cells[j]);
                }
            }
            
            if (rowData.Values.Any(v => !string.IsNullOrEmpty(v))) {
                result.Add(rowData);
            }
        }
        
        return result;
    }


    // convenience methods 
    // gets value at cell location
    public string GetCellValue(string colLetter = "", int row = 0, Cell cell = null) {
        if (cell == null) {
            cell = Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"{colLetter}{row}"); }
        string value = "";
        if (cell != null) { value = cell.InnerText; }
        
        if (cell != null && cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
            SharedStringTablePart stringTablePart = WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTablePart != null) value = stringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
        }
        return value;
    }

    // gets the in-place header (pre-defined)
    // public void GetHdrRow(SheetData sheetData, char salesOrIncome) {
    //     Row targetRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == 1);
    //     if (targetRow != null) {
    //         foreach (Cell cell in targetRow.Elements<Cell>()) {
    //             string cellValue = GetCellValue("", 0, null, cell);
    //             if (salesOrIncome == 's') { hdrRowS[$"{cell.CellReference}"] = $"{cellValue}"; }
    //             else { hdrRowR[$"{cell.CellReference}"] = $"{cellValue}"; }
    //         }
    //     }
    // }


    // IDisposable's needed methods, save and dispose .. 
    public void Save()
    {
        if (IsDisposed) throw new ObjectDisposedException("ExcelFile");
        spreadsheetDoc.Save();
    }

    public void Dispose()
    {
        if (!IsDisposed)
        {
            spreadsheetDoc.Dispose();
            IsDisposed = true;
        }
    }
}
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

public class WordTableModel
{
    public int? TableWidth { get; init; }
    public TableWidthUnitValues? WidthType { get; init; }
    public List<int?> ColumnWidths { get; init; } = new();
    public string TableStyleId { get; init; }
    public int ColumnCount => ColumnWidths.Count;
}
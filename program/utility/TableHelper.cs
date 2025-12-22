using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class TableFinder
{
    private readonly Body _body;
    private readonly Styles _styles;

    public TableFinder(MainDocumentPart mainPart)
    {
        _body = mainPart.Document.Body;
        // Corrected property name here
        _styles = mainPart.StyleDefinitionsPart?.Styles;
    }

public Table GetTableAfterHeading(string headingText, int startAfterXmlChar = 1000000)
{
    int currentXmlOffset = 0;
    OpenXmlElement targetContainer = null;

    // We iterate through direct children of the body (Paragraphs, Tables, SdtBlocks)
    foreach (var element in _body.Elements())
    {
        // 1. Update the offset based on the length of this element's XML
        currentXmlOffset += element.OuterXml.Length;

        // 2. Only start "listening" for the heading after we pass the XML threshold
        if (currentXmlOffset > startAfterXmlChar)
        {
            // We search inside the element for the heading text
            if (element.InnerText.Contains(headingText, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine(currentXmlOffset);
                targetContainer = element;
                break;
            }
        }
    }

    if (targetContainer == null)
    {
        Console.WriteLine($"DEBUG: Could not find '{headingText}' after XML offset {startAfterXmlChar}.");
        return null;
    }

    // 3. Find the first table that appears after the container where we found the text
    return targetContainer.ElementsAfter().OfType<Table>().FirstOrDefault();
}

    public TableInfo GetTableInfo(Table table)
    {
        if (table == null) return null;

        var props = table.GetFirstChild<TableProperties>();

        return new TableInfo
        {
            StyleId = props?.TableStyle?.Val?.Value,
            Borders = props?.TableBorders,
            Rows = GetRowInfos(table)
        };
    }

    private List<RowInfo> GetRowInfos(Table table)
    {
        return table.Elements<TableRow>().Select(r => 
        {
            var rowProps = r.GetFirstChild<TableRowProperties>();
            bool isHeaderRow = rowProps?.GetFirstChild<TableHeader>() != null;

            return new RowInfo
            {
                IsHeader = isHeaderRow,
                Cells = r.Elements<TableCell>().Select(c => 
                {
                    var cellProps = c.GetFirstChild<TableCellProperties>();
                    
                    // 1. Get the value as a string first
                    int? rawSpan = cellProps?.GridSpan?.Val?.Value;
                    
                    // 2. Coalesce to "1" and parse
                    int span = rawSpan ?? 1;

                    return new CellInfo
                    {
                        Text = c.InnerText?.Trim(), // Added Trim() for cleaner data
                        Width = cellProps?.TableCellWidth?.Width?.Value,
                        ColSpan = span
                    };
                }).ToList()
            };
        }).ToList();
    }
}

public class TableInfo
{
    public string StyleId { get; set; }
    public TableBorders Borders { get; set; }
    public List<RowInfo> Rows { get; set; }
}
public class RowInfo
{
    public bool IsHeader { get; set; }
    public List<CellInfo> Cells { get; set; }
}
public class CellInfo
{
    public string Text { get; set; }
    public string Width { get; set; }
    public int ColSpan { get; set; }
}
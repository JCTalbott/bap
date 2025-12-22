using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

public class WordTableEngine
{
    private readonly Body _body;
    private readonly MainDocumentPart _mainPart;

    public WordTableEngine(MainDocumentPart mainPart)
    {
        _mainPart = mainPart;
        _body = mainPart.Document.Body;
    }

    public Table FindTableAfterHeading(string headingText, int startAfterXmlChar = 0)
    {
        int currentXmlOffset = 0;
        OpenXmlElement target = null;

        foreach (var element in _body.Elements())
        {
            currentXmlOffset += element.OuterXml.Length;

            if (currentXmlOffset < startAfterXmlChar)
                continue;

            if (element.InnerText.Contains(headingText, StringComparison.OrdinalIgnoreCase))
            {
                target = element;
                break;
            }
        }

        return target?
            .ElementsAfter()
            .OfType<Table>()
            .FirstOrDefault();
    }

    public WordTableModel ExtractModel(Table table)
    {
        var tblPr = table.GetFirstChild<TableProperties>();
        var tblGrid = table.GetFirstChild<TableGrid>();

        return new WordTableModel
        {
            TableWidth = int.TryParse(tblPr?.TableWidth?.Width?.Value, out int w) ? w : null,
            WidthType = tblPr?.TableWidth?.Type?.Value,
            TableStyleId = tblPr?.TableStyle?.Val?.Value,
            ColumnWidths = tblGrid?
                .Elements<GridColumn>()
                .Select(c => int.TryParse(c.Width?.Value, out int cw) ? (int?)cw : null)
                .ToList()
                ?? new List<int?>()
        };
    }

    public Table BuildTableFromModel(WordTableModel model, List<Dictionary<string, string>> columnData)
    {
        var table = new Table();
        var tblPr = new TableProperties();

        if (!string.IsNullOrEmpty(model.TableStyleId))
            tblPr.AppendChild(new TableStyle { Val = model.TableStyleId });

        if (model.TableWidth != null)
        {
            tblPr.AppendChild(new TableWidth
            {
                Width = model.TableWidth.ToString(),
                Type = model.WidthType ?? TableWidthUnitValues.Dxa
            });
        }

        tblPr.AppendChild(new TableLayout { Type = TableLayoutValues.Fixed });
        table.AppendChild(tblPr);

        var grid = new TableGrid();
        foreach (var width in model.ColumnWidths)
            grid.AppendChild(new GridColumn { Width = width?.ToString() });
        table.AppendChild(grid);

        foreach (var rowData in columnData)
        {
            var row = new TableRow();
            
            var trPr = new TableRowProperties();
            trPr.AppendChild(new CantSplit());
            trPr.AppendChild(new TableRowHeight { Val = 360, HeightType = HeightRuleValues.AtLeast });
            row.AppendChild(trPr);

            foreach (var kv in rowData)
            {
                row.AppendChild(CreateCell(kv.Key, kv.Value));
            }

            table.AppendChild(row);
        }

        return table;
    }

    public void InsertAfter(Table referenceTable, Table newTable, bool pageBreakFirst = false)
    {
        OpenXmlElement insertAfter = referenceTable;

        // if (pageBreakFirst)
        // {
        //     var breakPara = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
        //     _body.InsertAfter(breakPara, referenceTable);
        //     insertAfter = breakPara;
        // }

        _body.InsertAfter(newTable, insertAfter);
    }

    private TableCell CreateCell(string key, string value)
    {
        var run = new Run();
        run.AppendChild(new Text(key));
        run.AppendChild(new Break());
        run.AppendChild(new Text(value));

        return new TableCell(
            new Paragraph(run),
            new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top })
        );
    }
}
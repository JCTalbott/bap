using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;


class TableBoy {
  public Table table;
  public TableBoy() {

  }
  public void CreateTable()
  {
          Table table = new Table();

          TableProperties tblProp = new TableProperties(
              new TableBorders(
                  new TopBorder()
                  {
                      Val =
                      new EnumValue<BorderValues>(BorderValues.Dashed),
                      Size = 24
                  },
                  new BottomBorder()
                  {
                      Val =
                      new EnumValue<BorderValues>(BorderValues.Dashed),
                      Size = 24
                  },
                  new LeftBorder()
                  {
                      Val =
                      new EnumValue<BorderValues>(BorderValues.Dashed),
                      Size = 24
                  },
                  new RightBorder()
                  {
                      Val =
                      new EnumValue<BorderValues>(BorderValues.Dashed),
                      Size = 24
                  },
                  new InsideHorizontalBorder()
                  {
                      Val =
                      new EnumValue<BorderValues>(BorderValues.Dashed),
                      Size = 24
                  },
                  new InsideVerticalBorder()
                  {
                      Val =
                      new EnumValue<BorderValues>(BorderValues.Dashed),
                      Size = 24
                  }
              )
          );

          table.AppendChild<TableProperties>(tblProp);
          TableRow tr = new TableRow();

          TableCell tc1 = new TableCell();

          tc1.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

          tc1.Append(new Paragraph(new Run(new Text("some text"))));
          tr.Append(tc1);
          TableCell tc2 = new TableCell(tc1.OuterXml);
          tr.Append(tc2);

          table.Append(tr);

          //body.Append(table);
      }
}
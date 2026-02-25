// using DocumentFormat.OpenXml;
// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Wordprocessing;
// using A = DocumentFormat.OpenXml.Drawing;
// using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
// using System;
// using System.IO;
// using System.Linq;
// using System.Collections.Generic;

// public class WordImageEngine
// {
//     private readonly MainDocumentPart _mainPart;

//     public WordImageEngine(MainDocumentPart mainPart)
//     {
//         _mainPart = mainPart;
//     }

//     public string RegisterImage(Stream imageStream, string extension)
//     {
//         PartTypeInfo type = extension.ToLower() switch
//         {
//             ".jpg" or ".jpeg" => ImagePartType.Jpeg,
//             ".png" => ImagePartType.Png,
//             _ => ImagePartType.Png
//         };

//         ImagePart imagePart = _mainPart.AddImagePart(type);
//         imagePart.FeedData(imageStream);
//         return _mainPart.GetIdOfPart(imagePart);
//     }

//     public Drawing CloneAndReplace(Drawing templateDrawing, string newRelationshipId)
//     {
//         Drawing newDrawing = (Drawing)templateDrawing.CloneNode(true);
        
//         var blip = newDrawing.Descendants<A.Blip>().FirstOrDefault();
//         if (blip != null) blip.Embed = newRelationshipId;

//         var docProp = newDrawing.Descendants<DW.DocProperties>().FirstOrDefault();
//         if (docProp != null)
//         {
//             docProp.Id = (uint)Random.Shared.Next(1, int.MaxValue);
//             docProp.Name = "Photo_" + docProp.Id;
//         }

//         return newDrawing;
//     }

//     public Drawing FindImageByAltText(string altText)
//     {
//         return _mainPart.Document.Body.Descendants<Drawing>()
//             .FirstOrDefault(d => d.Descendants<DW.DocProperties>()
//             .Any(prop => prop.Name == altText || prop.Description == altText));
//     }

//     public void GeneratePhotoPages(Drawing template, List<PhotoInput> photos)
//     {
//         var anchorParagraph = template.Ancestors<Paragraph>().FirstOrDefault();
//         if (anchorParagraph == null) return;

//         OpenXmlElement currentAnchor = anchorParagraph;

//         for (int i = 0; i < photos.Count; i++)
//         {
//             if (!File.Exists(photos[i].Path)) continue;

//             using (FileStream fs = new FileStream(photos[i].Path, FileMode.Open, FileAccess.Read))
//             {
//                 string rId = RegisterImage(fs, Path.GetExtension(photos[i].Path));
//                 Drawing drawing = CloneAndReplace(template, rId);

//                 var imgPara = new Paragraph(
//                     new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
//                     new Run(drawing)
//                 );

//                 var descPara = new Paragraph(
//                     new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
//                     new Run(new RunProperties(new Bold()), new Text(photos[i].Description))
//                 );

//                 currentAnchor.InsertAfterSelf(imgPara);
//                 imgPara.InsertAfterSelf(descPara);
//                 currentAnchor = descPara;

//                 bool isSecondImageOnPage = (i + 1) % 2 == 0;
//                 bool isLastImageOverall = (i == photos.Count - 1);

//                 if (!isLastImageOverall)
//                 {
//                     if (isSecondImageOnPage)
//                     {
//                         var pageBreakPara = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
//                         currentAnchor.InsertAfterSelf(pageBreakPara);
//                         currentAnchor = pageBreakPara;
//                     }
//                     else
//                     {
//                         var spacer = new Paragraph(new Run(new Break()));
//                         currentAnchor.InsertAfterSelf(spacer);
//                         currentAnchor = spacer;
//                     }
//                 }
//             }
//         }

//         anchorParagraph.Remove();
//     }
// }



using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

public class WordImageEngine
{
    private readonly MainDocumentPart _mainPart;

    public WordImageEngine(MainDocumentPart mainPart) => _mainPart = mainPart;
    public Drawing FindImageByAltText(string altText)
    {
        return _mainPart.Document.Body.Descendants<Drawing>()
            .FirstOrDefault(d => d.Descendants<DW.DocProperties>()
            .Any(prop => prop.Name == altText || prop.Description == altText));
    }
    public void GeneratePhotoPages(Drawing template, List<PhotoInput> photos)
    {
        var anchorParagraph = template.Ancestors<Paragraph>().FirstOrDefault();
        if (anchorParagraph == null) return;

        // Get dimensions from template (EMUs)
        var extent = template.Descendants<DW.Extent>().FirstOrDefault();
        long width = extent?.Cx ?? 5486400L; 
        long height = extent?.Cy ?? 3657600L;

        // Page logic: 2 per page
        for (int i = 0; i < photos.Count; i += 2)
        {
            Table table = CreatePhotoTable(photos.Skip(i).Take(2).ToList(), width, height);
            anchorParagraph.InsertBeforeSelf(table);
            
            // Add page break after every table except the last one
            if (i + 2 < photos.Count)
            {
                anchorParagraph.InsertBeforeSelf(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
            }
        }

        anchorParagraph.Remove();
    }

    private Table CreatePhotoTable(List<PhotoInput> pagePhotos, long width, long height)
    {
        Table table = new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableJustification { Val = TableRowAlignmentValues.Center },
                new TableBorders(
                    new TopBorder { Val = BorderValues.None },
                    new BottomBorder { Val = BorderValues.None },
                    new LeftBorder { Val = BorderValues.None },
                    new RightBorder { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.None },
                    new InsideVerticalBorder { Val = BorderValues.None }
                )
            )
        );

        foreach (var photo in pagePhotos)
        {
            using var fs = new FileStream(photo.Path, FileMode.Open, FileAccess.Read);
            string rId = RegisterImage(fs, Path.GetExtension(photo.Path));
            
            // Row 1: The Image
            table.Append(new TableRow(
                new TableRowProperties(new TableRowHeight { Val = (uint)(height / 635), HeightType = HeightRuleValues.AtLeast }),
                new TableCell(new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Center }, new SpacingBetweenLines { After = "0" }),
                    new Run(CreateInlineDrawing(rId, width, height))
                ))
            ));

            // Row 2: The Caption (snug to image)
            table.Append(new TableRow(
                new TableCell(new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines { Before = "0", After = "120" }
                    ),
                    new Run(
                        new RunProperties(
                            // 1. Remove Bold (implicitly done by not adding new Bold())
                            // 2. Set Font Size to 10pt (Value = 20)
                            new FontSize { Val = "20" }, 
                            // 3. Set Font Family to Calibri
                            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }
                        ),
                        new Text(photo.Description)
                    )
                ))
            ));
        }

        return table;
    }

    public string RegisterImage(Stream imageStream, string extension)
    {
        PartTypeInfo type = extension.ToLower().Contains("png") ? ImagePartType.Png : ImagePartType.Jpeg;
        ImagePart imagePart = _mainPart.AddImagePart(type);
        imagePart.FeedData(imageStream);
        return _mainPart.GetIdOfPart(imagePart);
    }

    private Drawing CreateInlineDrawing(string relationshipId, long width, long height)
    {
        var inline = new DW.Inline(
            new DW.Extent { Cx = width, Cy = height },
            new DW.DocProperties { Id = (uint)Random.Shared.Next(1, int.MaxValue), Name = "Img" },
            new A.Graphic(new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties { Id = (uint)Random.Shared.Next(1, int.MaxValue), Name = "Img" },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(new A.Blip { Embed = relationshipId }),
                    new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }))
            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
        );
        return new Drawing(inline);
    }
}
public class PhotoInput
{
    public string Path { get; set; }
    public string Description { get; set; }
}

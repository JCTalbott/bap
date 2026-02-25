using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;


public class ImageFinder {
public static void PrintImageInfo(MainDocumentPart mainPart)
{
    // Images live inside 'Drawing' elements in the document body
    var drawings = mainPart.Document.Body.Descendants<Drawing>().ToList();

    Console.WriteLine($"--- Image Inspection ({drawings.Count} found) ---");

    foreach (var drawing in drawings)
    {
        // 1. Get the Extents (Size)
        // Word stores sizes in EMUs (English Metric Units)
        // 1 inch = 914,400 EMUs
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
        
        // 2. Get the Reference (Relationship ID)
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        string relId = blip?.Embed?.Value;

        // 3. Get the Filename/Alt Text
        var docProp = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        string name = docProp?.Name ?? "Unnamed";
        string desc = docProp?.Description ?? "No Alt Text";

        // 4. Get Style/Transform (Rotation, Flip, etc.)
        var xfrm = drawing.Descendants<A.Transform2D>().FirstOrDefault();

        Console.WriteLine($"Name:      {name}");
        Console.WriteLine($"Alt Text:  {desc}");
        Console.WriteLine($"Rel ID:    {relId}");

        if (extent != null)
        {
            double widthInches = extent.Cx / 914400.0;
            double heightInches = extent.Cy / 914400.0;
            Console.WriteLine($"Size:      {widthInches:F2}\" x {heightInches:F2}\" ({extent.Cx} x {extent.Cy} EMUs)");
        }

        if (xfrm?.Rotation != null)
        {
            // Rotation is stored in 60,000ths of a degree
            Console.WriteLine($"Rotation:  {xfrm.Rotation.Value / 60000.0}°");
        }

        // 5. Locate the actual part to see the file size/type
        if (!string.IsNullOrEmpty(relId))
        {
            var imagePart = mainPart.GetPartById(relId);
            Console.WriteLine($"Type:      {imagePart.ContentType}");
        }
        
        Console.WriteLine(new string('-', 30));
    }
}
}
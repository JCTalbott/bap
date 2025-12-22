using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

public class ProperExcelToWord {
    
    public void ProcessDocuments(DataHolder dh) {
        using (WordprocessingDocument report = WordprocessingDocument.Open("doc-files/pony.docx", true)) {
            WordprocessingDocument newPlace = report.Clone("doc-files/pony_processed_proper.docx");
            MainDocumentPart mainPart = newPlace.MainDocumentPart;
            Body body = mainPart.Document.Body;
            
            Console.WriteLine("Processing Word document with actual placeholders...");
            Console.WriteLine($"DataHolder has {dh.subjectValues.Count} items to process");
            
            // Process all the actual placeholders found in the Word document
            ProcessPlaceholders(body, dh);
            
            newPlace.MainDocumentPart.Document.Body = body;
            newPlace.Save();
            
            Console.WriteLine("Successfully processed document with proper placeholders!");
        }
    }
    
    private void ProcessPlaceholders(Body body, DataHolder dh) {
        // Replace all the actual placeholders found in the Word document
        var replacements = new Dictionary<string, string>();
        
        // Add all the data from DataHolder
        foreach (var kvp in dh.subjectValues) {
            replacements[kvp.Key] = kvp.Value;
        }
        
        // Process paragraphs
        var paragraphs = body.Descendants<Paragraph>().ToList();
        foreach (var paragraph in paragraphs) {
            var text = paragraph.InnerText;
            if (text.Contains("~@")) {
                foreach (var replacement in replacements) {
                    if (text.Contains(replacement.Key)) {
                        paragraph.InnerXml = paragraph.InnerXml.Replace(replacement.Key, replacement.Value);
                        Console.WriteLine($"Replaced {replacement.Key} with {replacement.Value}");
                    }
                }
            }
        }
        
        // Process table cells
        var cells = body.Descendants<TableCell>().ToList();
        foreach (var cell in cells) {
            var text = cell.InnerText;
            if (text.Contains("~@")) {
                foreach (var replacement in replacements) {
                    if (text.Contains(replacement.Key)) {
                        cell.InnerXml = cell.InnerXml.Replace(replacement.Key, replacement.Value);
                        Console.WriteLine($"Replaced {replacement.Key} with {replacement.Value} in table cell");
                    }
                }
            }
        }
        
        // Process runs (for inline text)
        var runs = body.Descendants<Run>().ToList();
        foreach (var run in runs) {
            var text = run.InnerText;
            if (text.Contains("~@")) {
                foreach (var replacement in replacements) {
                    if (text.Contains(replacement.Key)) {
                        run.InnerXml = run.InnerXml.Replace(replacement.Key, replacement.Value);
                        Console.WriteLine($"Replaced {replacement.Key} with {replacement.Value} in run");
                    }
                }
            }
        }
    }
}

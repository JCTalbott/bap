using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

public class WordBeholder : IDisposable {
    public bool IsDisposed { get; private set; } = false;
    public string FilePath { get; }
    public WordprocessingDocument? WordDoc { get; private set; }
    public Body? DocumentBody => WordDoc?.MainDocumentPart?.Document?.Body;
    
    public WordBeholder(string filePath) {
        FilePath = filePath;
    }

    public void Open(bool isEditable = true) {
        if (WordDoc != null) return;
        WordDoc = WordprocessingDocument.Open(FilePath, isEditable);
    }

    public WordprocessingDocument Clone(string newFilePath) {
        if (WordDoc == null) throw new InvalidOperationException("Document must be opened before cloning");
        return (WordprocessingDocument)WordDoc.Clone(newFilePath);
    }

    public void Save()
    {
        if (IsDisposed) throw new ObjectDisposedException("WordFile");
        WordDoc?.Save();
    }

    // Enhanced methods for document manipulation
    public List<string> FindPlaceholders(string pattern = @"~@\w+") {
        var placeholders = new List<string>();
        if (DocumentBody == null) return placeholders;
        
        var regex = new Regex(pattern);
        var matches = regex.Matches(DocumentBody.InnerText);
        
        foreach (Match match in matches) {
            if (!placeholders.Contains(match.Value)) {
                placeholders.Add(match.Value);
            }
        }
        
        return placeholders;
    }

    public bool ReplaceText(string oldText, string newText) {
        if (DocumentBody == null) return false;
        
        var paragraphs = DocumentBody.Descendants<Paragraph>().ToList();
        bool replaced = false;
        
        foreach (var paragraph in paragraphs) {
            if (paragraph.InnerText.Contains(oldText)) {
                paragraph.InnerXml = paragraph.InnerXml.Replace(oldText, newText);
                replaced = true;
            }
        }
        
        return replaced;
    }

    public List<Table> GetTables() {
        if (DocumentBody == null) return new List<Table>();
        return DocumentBody.Descendants<Table>().ToList();
    }

    public List<Table> FindTablesWithPlaceholder(string placeholder) {
        var tables = GetTables();
        var matchingTables = new List<Table>();
        
        foreach (var table in tables) {
            if (table.InnerText.Contains(placeholder)) {
                matchingTables.Add(table);
            }
        }
        
        return matchingTables;
    }

    public void Dispose()
    {
        if (!IsDisposed)
        {
            WordDoc?.Dispose();
            IsDisposed = true;
        }
    }
}
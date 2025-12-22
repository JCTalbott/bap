using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

public class EnhancedExcelToWord {
    private DataMapper _dataMapper;
    private ExcelDataWorker _dataWorker;
    
    public EnhancedExcelToWord() {
        _dataMapper = new DataMapper();
        _dataWorker = new ExcelDataWorker();
    }
    
    public void ProcessDocuments(string excelFilePath, string wordTemplatePath, string outputPath) {
        using (var excelDoorway = new ExcelDoorway(excelFilePath, false))
        using (var wordBeholder = new WordBeholder(wordTemplatePath)) {
            
            wordBeholder.Open(true);
            var clonedDoc = wordBeholder.Clone(outputPath);
            
            try {
                // Extract all data from Excel
                var excelData = ExtractAllExcelData(excelDoorway);
                
                // Process the Word document
                ProcessWordDocument(clonedDoc, excelData);
                
                clonedDoc.Save();
                Console.WriteLine($"Successfully processed document: {outputPath}");
                
            } catch (Exception ex) {
                Console.WriteLine($"Error processing documents: {ex.Message}");
                throw;
            } finally {
                clonedDoc.Dispose();
            }
        }
    }
    
    private Dictionary<string, Dictionary<string, string>> ExtractAllExcelData(ExcelDoorway excelDoorway) {
        var allData = new Dictionary<string, Dictionary<string, string>>();
        
        var sheets = excelDoorway.getSheets();
        foreach (var sheet in sheets.Elements<Sheet>()) {
            var sheetName = sheet.Name;
            Console.WriteLine($"Processing sheet: {sheetName}");
            
            try {
                // Extract structured table data
                var tableData = excelDoorway.ExtractTableData(sheetName, 1);
                if (tableData.Any()) {
                    var sheetData = new Dictionary<string, string>();
                    
                    // Flatten the first row of data for simple placeholders
                    if (tableData.Count > 0) {
                        var firstRow = tableData[0];
                        foreach (var kvp in firstRow) {
                            sheetData[kvp.Key] = kvp.Value;
                        }
                    }
                    
                    allData[sheetName] = sheetData;
                    
                    // Also store full table data for complex mappings
                    allData[$"{sheetName}_table"] = new Dictionary<string, string>();
                }
                
                // Extract raw cell data as well
                var rawData = excelDoorway.ExtractSheetData(sheetName);
                if (rawData.Any()) {
                    allData[$"{sheetName}_raw"] = rawData;
                }
                
            } catch (Exception ex) {
                Console.WriteLine($"Error processing sheet {sheetName}: {ex.Message}");
            }
        }
        
        return allData;
    }
    
    private void ProcessWordDocument(WordprocessingDocument wordDoc, Dictionary<string, Dictionary<string, string>> excelData) {
        var mainPart = wordDoc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;
        
        var body = mainPart.Document.Body;
        
        // Find all placeholders in the document
        var placeholders = FindAllPlaceholders(body);
        Console.WriteLine($"Found {placeholders.Count} placeholders in document");
        
        // Map Excel data to placeholders
        var mappedData = _dataMapper.MapDataToPlaceholders(excelData);
        
        // Replace simple text placeholders
        ReplaceTextPlaceholders(body, mappedData);
        
        // Process table placeholders and complex structures
        ProcessTablePlaceholders(body, excelData);
        
        // Process any remaining complex structures
        ProcessComplexStructures(body, excelData);
    }
    
    private List<string> FindAllPlaceholders(Body body) {
        var placeholders = new List<string>();
        var regex = new Regex(@"~@\w+");
        var matches = regex.Matches(body.InnerText);
        
        foreach (Match match in matches) {
            if (!placeholders.Contains(match.Value)) {
                placeholders.Add(match.Value);
            }
        }
        
        return placeholders;
    }
    
    private void ReplaceTextPlaceholders(Body body, Dictionary<string, string> mappedData) {
        foreach (var kvp in mappedData) {
            var placeholder = kvp.Key;
            var value = kvp.Value;
            
            // Apply appropriate transformation based on placeholder name
            value = ApplyTransformation(placeholder, value);
            
            // Replace in paragraphs
            var paragraphs = body.Descendants<Paragraph>().ToList();
            foreach (var paragraph in paragraphs) {
                if (paragraph.InnerText.Contains(placeholder)) {
                    paragraph.InnerXml = paragraph.InnerXml.Replace(placeholder, value);
                }
            }
            
            // Replace in table cells
            var cells = body.Descendants<TableCell>().ToList();
            foreach (var cell in cells) {
                if (cell.InnerText.Contains(placeholder)) {
                    cell.InnerXml = cell.InnerXml.Replace(placeholder, value);
                }
            }
        }
    }
    
    private void ProcessTablePlaceholders(Body body, Dictionary<string, Dictionary<string, string>> excelData) {
        var tables = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
        
        foreach (var table in tables) {
            // Check if this table has placeholders that need special processing
            var tableText = table.InnerText;
            
            if (tableText.Contains("~@table_")) {
                ProcessDynamicTable(table, excelData);
            } else if (tableText.Contains("~@")) {
                // Process regular table placeholders
                ProcessRegularTablePlaceholders(table, excelData);
            }
        }
    }
    
    private void ProcessDynamicTable(DocumentFormat.OpenXml.Wordprocessing.Table table, Dictionary<string, Dictionary<string, string>> excelData) {
        // This is a placeholder for more complex table processing
        // Could expand rows, modify structure, etc.
        Console.WriteLine("Processing dynamic table structure");
    }
    
    private void ProcessRegularTablePlaceholders(DocumentFormat.OpenXml.Wordprocessing.Table table, Dictionary<string, Dictionary<string, string>> excelData) {
        var cells = table.Descendants<TableCell>().ToList();
        
        foreach (var cell in cells) {
            var cellText = cell.InnerText;
            var regex = new Regex(@"~@\w+");
            var matches = regex.Matches(cellText);
            
            foreach (Match match in matches) {
                var placeholder = match.Value;
                var value = FindValueForPlaceholder(placeholder, excelData);
                
                if (!string.IsNullOrEmpty(value)) {
                    value = ApplyTransformation(placeholder, value);
                    cell.InnerXml = cell.InnerXml.Replace(placeholder, value);
                }
            }
        }
    }
    
    private void ProcessComplexStructures(Body body, Dictionary<string, Dictionary<string, string>> excelData) {
        // Process any complex document structures that need special handling
        // This could include headers, footers, text boxes, etc.
        
        // Look for section placeholders that might need special processing
        var paragraphs = body.Descendants<Paragraph>().ToList();
        foreach (var paragraph in paragraphs) {
            var text = paragraph.InnerText;
            
            // Handle section headers that might need dynamic content
            if (text.Contains("~@section_")) {
                ProcessSectionPlaceholder(paragraph, excelData);
            }
        }
    }
    
    private void ProcessSectionPlaceholder(Paragraph paragraph, Dictionary<string, Dictionary<string, string>> excelData) {
        // Extract section type from placeholder
        var regex = new Regex(@"~@section_(\w+)");
        var match = regex.Match(paragraph.InnerText);
        
        if (match.Success) {
            var sectionType = match.Groups[1].Value;
            Console.WriteLine($"Processing section: {sectionType}");
            
            // This is where you could add specific logic for different section types
            // For now, just replace the placeholder with the section type
            paragraph.InnerXml = paragraph.InnerXml.Replace(match.Value, sectionType);
        }
    }
    
    private string FindValueForPlaceholder(string placeholder, Dictionary<string, Dictionary<string, string>> excelData) {
        // Try to find the value for a placeholder in the Excel data
        foreach (var sheetData in excelData.Values) {
            if (sheetData.ContainsKey(placeholder)) {
                return sheetData[placeholder];
            }
        }
        
        // Try with modified placeholder names
        var cleanPlaceholder = placeholder.Replace("~@", "").ToLower();
        foreach (var sheetData in excelData.Values) {
            foreach (var kvp in sheetData) {
                if (kvp.Key.ToLower().Contains(cleanPlaceholder)) {
                    return kvp.Value;
                }
            }
        }
        
        return "";
    }
    
    private string ApplyTransformation(string placeholder, string value) {
        // Apply appropriate transformation based on placeholder name
        if (placeholder.Contains("price") || placeholder.Contains("cost") || placeholder.Contains("amount")) {
            return _dataWorker.currencyWorker(value);
        } else if (placeholder.Contains("date")) {
            return _dataWorker.dateWorker(value);
        } else if (placeholder.Contains("percent") || placeholder.Contains("ratio")) {
            return _dataWorker.percentageWorker(value);
        }
        
        return value;
    }
    
    public void ConfigureMapping(string placeholder, string excelColumn, string transformer = null) {
        _dataMapper.AddMappingRule(placeholder, excelColumn);
        
        if (!string.IsNullOrEmpty(transformer)) {
            // Add custom transformer if needed
            Console.WriteLine($"Added mapping: {placeholder} -> {excelColumn} with transformer: {transformer}");
        }
    }
    
    public void ListAvailablePlaceholders() {
        var placeholders = _dataMapper.GetAvailablePlaceholders();
        Console.WriteLine("Available placeholders:");
        foreach (var placeholder in placeholders) {
            Console.WriteLine($"  {placeholder}");
        }
    }
}

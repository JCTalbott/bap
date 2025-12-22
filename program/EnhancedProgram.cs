using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

public class EnhancedProgram {
    private ConfigurationManager _configManager;
    private ValidationEngine _validationEngine;
    private EnhancedExcelToWord _excelToWord;
    
    public EnhancedProgram() {
        _configManager = new ConfigurationManager();
        _validationEngine = new ValidationEngine(_configManager.GetConfiguration().Validation);
        _excelToWord = new EnhancedExcelToWord();
    }
    
    public void Run(string excelFilePath = "excel-files/pony.xlsx", string wordTemplatePath = "doc-files/pony.docx") {
        Console.WriteLine("=== Enhanced Excel to Word Data Injection System ===");
        
        try {
            // Validate input files
            if (!File.Exists(excelFilePath)) {
                Console.WriteLine($"Excel file not found: {excelFilePath}");
                return;
            }
            
            if (!File.Exists(wordTemplatePath)) {
                Console.WriteLine($"Word template not found: {wordTemplatePath}");
                return;
            }
            
            Console.WriteLine($"Processing Excel file: {excelFilePath}");
            Console.WriteLine($"Using Word template: {wordTemplatePath}");
            
            // Extract and validate Excel data
            var excelData = ExtractAndValidateExcelData(excelFilePath);
            
            // Validate Word template
            ValidateWordTemplate(wordTemplatePath);
            
            // Process documents
            var outputPath = GenerateOutputPath(wordTemplatePath);
            _excelToWord.ProcessDocuments(excelFilePath, wordTemplatePath, outputPath);
            
            Console.WriteLine($"\nProcessing completed successfully!");
            Console.WriteLine($"Output file: {outputPath}");
            
        } catch (Exception ex) {
            Console.WriteLine($"Error during processing: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }
    
    private Dictionary<string, Dictionary<string, string>> ExtractAndValidateExcelData(string excelFilePath) {
        Console.WriteLine("\n--- Extracting Excel Data ---");
        
        using (var excelDoorway = new ExcelDoorway(excelFilePath, false)) {
            var allData = new Dictionary<string, Dictionary<string, string>>();
            
            var sheets = excelDoorway.getSheets();
            foreach (var sheet in sheets.Elements<Sheet>()) {
                var sheetName = sheet.Name;
                Console.WriteLine($"Processing sheet: {sheetName}");
                
                try {
                    // Extract table data
                    var tableData = excelDoorway.ExtractTableData(sheetName, 1);
                    if (tableData.Any()) {
                        var sheetData = new Dictionary<string, string>();
                        
                        // Use the first row of data for simple mappings
                        if (tableData.Count > 0) {
                            var firstRow = tableData[0];
                            foreach (var kvp in firstRow) {
                                sheetData[kvp.Key] = kvp.Value;
                                Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
                            }
                        }
                        
                        allData[sheetName] = sheetData;
                    }
                    
                    // Extract raw cell data
                    var rawData = excelDoorway.ExtractSheetData(sheetName);
                    if (rawData.Any()) {
                        allData[$"{sheetName}_raw"] = rawData;
                        Console.WriteLine($"  Extracted {rawData.Count} raw cell values");
                    }
                    
                } catch (Exception ex) {
                    Console.WriteLine($"  Error processing sheet {sheetName}: {ex.Message}");
                }
            }
            
            // Validate the extracted data
            Console.WriteLine("\n--- Validating Excel Data ---");
            var validationResult = _validationEngine.ValidateData(allData, _configManager.GetConfiguration());
            validationResult.PrintResults();
            
            return allData;
        }
    }
    
    private void ValidateWordTemplate(string wordTemplatePath) {
        Console.WriteLine("\n--- Validating Word Template ---");
        
        using (var wordDoc = WordprocessingDocument.Open(wordTemplatePath, false)) {
            var validationResult = _validationEngine.ValidateWordDocument(wordDoc, _configManager.GetConfiguration());
            validationResult.PrintResults();
        }
    }
    
    private string GenerateOutputPath(string templatePath) {
        var fileName = Path.GetFileNameWithoutExtension(templatePath);
        var directory = Path.GetDirectoryName(templatePath);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(directory ?? "", $"{fileName}_processed_{timestamp}.docx");
    }
    
    public void ConfigureMappings() {
        Console.WriteLine("\n--- Configuring Data Mappings ---");
        
        // Add custom mappings based on your specific needs
        _excelToWord.ConfigureMapping("~@property_name", "Property Name");
        _excelToWord.ConfigureMapping("~@property_address", "Address");
        _excelToWord.ConfigureMapping("~@sale_price", "Sale Price", "currency");
        _excelToWord.ConfigureMapping("~@date_of_sale", "Date of Sale", "date");
        
        Console.WriteLine("Custom mappings configured.");
    }
    
    public void ListAvailablePlaceholders() {
        Console.WriteLine("\n--- Available Placeholders ---");
        _excelToWord.ListAvailablePlaceholders();
    }
    
    public void ShowConfiguration() {
        Console.WriteLine("\n--- Current Configuration ---");
        var config = _configManager.GetConfiguration();
        
        Console.WriteLine("Excel Sheets:");
        foreach (var sheet in config.ExcelSheets) {
            Console.WriteLine($"  {sheet.Key}: {sheet.Value.Name}");
            foreach (var mapping in sheet.Value.Mappings) {
                Console.WriteLine($"    {mapping.Key} -> {mapping.Value}");
            }
        }
        
        Console.WriteLine("\nWord Templates:");
        foreach (var template in config.WordTemplates) {
            Console.WriteLine($"  {template.Key}: {template.Value.Name}");
            Console.WriteLine($"    Template: {template.Value.TemplatePath}");
            Console.WriteLine($"    Output: {template.Value.OutputPath}");
        }
    }
    
    public void InteractiveMode() {
        Console.WriteLine("\n=== Interactive Mode ===");
        Console.WriteLine("Commands:");
        Console.WriteLine("  run - Process documents");
        Console.WriteLine("  config - Show configuration");
        Console.WriteLine("  mappings - Configure mappings");
        Console.WriteLine("  placeholders - List available placeholders");
        Console.WriteLine("  validate - Validate files only");
        Console.WriteLine("  exit - Exit program");
        
        while (true) {
            Console.Write("\n> ");
            var command = Console.ReadLine()?.ToLower().Trim();
            
            switch (command) {
                case "run":
                    Run();
                    break;
                case "config":
                    ShowConfiguration();
                    break;
                case "mappings":
                    ConfigureMappings();
                    break;
                case "placeholders":
                    ListAvailablePlaceholders();
                    break;
                case "validate":
                    ValidateFilesOnly();
                    break;
                case "exit":
                    return;
                default:
                    Console.WriteLine("Unknown command. Type 'exit' to quit.");
                    break;
            }
        }
    }
    
    private void ValidateFilesOnly() {
        Console.WriteLine("\n--- File Validation Only ---");
        
        var excelFilePath = "excel-files/pony.xlsx";
        var wordTemplatePath = "doc-files/pony.docx";
        
        if (File.Exists(excelFilePath) && File.Exists(wordTemplatePath)) {
            ExtractAndValidateExcelData(excelFilePath);
            ValidateWordTemplate(wordTemplatePath);
        } else {
            Console.WriteLine("Required files not found for validation.");
        }
    }
}

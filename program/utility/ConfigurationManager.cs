using System.Text.Json;

public class ConfigurationManager {
    private string _configFilePath;
    private DocumentMappingConfig _config;
    
    public ConfigurationManager(string configFilePath = "mapping-config.json") {
        _configFilePath = configFilePath;
        _config = LoadConfiguration();
    }
    
    public DocumentMappingConfig GetConfiguration() {
        return _config;
    }
    
    public void SaveConfiguration() {
        try {
            var jsonString = JsonSerializer.Serialize(_config, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_configFilePath, jsonString);
            Console.WriteLine($"Configuration saved to {_configFilePath}");
        } catch (Exception ex) {
            Console.WriteLine($"Error saving configuration: {ex.Message}");
        }
    }
    
    private DocumentMappingConfig LoadConfiguration() {
        if (File.Exists(_configFilePath)) {
            try {
                var jsonString = File.ReadAllText(_configFilePath);
                return JsonSerializer.Deserialize<DocumentMappingConfig>(jsonString) ?? CreateDefaultConfiguration();
            } catch (Exception ex) {
                Console.WriteLine($"Error loading configuration: {ex.Message}");
                return CreateDefaultConfiguration();
            }
        }
        
        var defaultConfig = CreateDefaultConfiguration();
        SaveConfiguration(); // Save default config
        return defaultConfig;
    }
    
    private DocumentMappingConfig CreateDefaultConfiguration() {
        return new DocumentMappingConfig {
            ExcelSheets = new Dictionary<string, SheetMapping> {
                ["Subj"] = new SheetMapping {
                    Name = "Subject Property",
                    DataType = "Property",
                    Mappings = new Dictionary<string, string> {
                        ["~@property_name"] = "Property Name",
                        ["~@property_address"] = "Address",
                        ["~@property_owner"] = "Owner",
                        ["~@building_area"] = "Building Area",
                        ["~@site_area"] = "Site Area"
                    }
                },
                ["SSum"] = new SheetMapping {
                    Name = "Sales Summary",
                    DataType = "Sales",
                    Mappings = new Dictionary<string, string> {
                        ["~@sale_price"] = "Sale Price",
                        ["~@date_of_sale"] = "Date of Sale",
                        ["~@price_per_sf"] = "Price/SF"
                    }
                },
                ["RSum"] = new SheetMapping {
                    Name = "Rental Summary",
                    DataType = "Rental",
                    Mappings = new Dictionary<string, string> {
                        ["~@rental_income"] = "Rental Income",
                        ["~@expenses"] = "Expenses",
                        ["~@net_income"] = "Net Income"
                    }
                }
            },
            Transformers = new Dictionary<string, TransformerConfig> {
                ["currency"] = new TransformerConfig {
                    Name = "Currency",
                    Function = "FormatCurrency",
                    Parameters = new Dictionary<string, string> {
                        ["format"] = "C"
                    }
                },
                ["date"] = new TransformerConfig {
                    Name = "Date",
                    Function = "FormatDate",
                    Parameters = new Dictionary<string, string> {
                        ["format"] = "MM/dd/yyyy"
                    }
                },
                ["percentage"] = new TransformerConfig {
                    Name = "Percentage",
                    Function = "FormatPercentage",
                    Parameters = new Dictionary<string, string> {
                        ["format"] = "P1"
                    }
                }
            },
            WordTemplates = new Dictionary<string, WordTemplateConfig> {
                ["report"] = new WordTemplateConfig {
                    Name = "Standard Report",
                    TemplatePath = "doc-files/report.docx",
                    OutputPath = "doc-files/output-report.docx",
                    PlaceholderPattern = @"~@\w+",
                    SpecialSections = new List<string> {
                        "Executive Summary",
                        "Property Analysis",
                        "Market Analysis",
                        "Financial Analysis"
                    }
                }
            },
            Validation = new ValidationConfig {
                RequiredFields = new List<string> {
                    "~@property_name",
                    "~@property_address",
                    "~@sale_price"
                },
                MaxFieldLength = 1000,
                ValidateNumbers = true,
                ValidateDates = true
            }
        };
    }
    
    public void AddSheetMapping(string sheetName, SheetMapping mapping) {
        _config.ExcelSheets[sheetName] = mapping;
    }
    
    public void AddWordTemplate(string templateName, WordTemplateConfig template) {
        _config.WordTemplates[templateName] = template;
    }
    
    public void UpdateMapping(string sheetName, string placeholder, string excelColumn) {
        if (_config.ExcelSheets.ContainsKey(sheetName)) {
            _config.ExcelSheets[sheetName].Mappings[placeholder] = excelColumn;
        }
    }
}

public class DocumentMappingConfig {
    public Dictionary<string, SheetMapping> ExcelSheets { get; set; } = new();
    public Dictionary<string, TransformerConfig> Transformers { get; set; } = new();
    public Dictionary<string, WordTemplateConfig> WordTemplates { get; set; } = new();
    public ValidationConfig Validation { get; set; } = new();
}

public class SheetMapping {
    public string Name { get; set; } = "";
    public string DataType { get; set; } = "";
    public Dictionary<string, string> Mappings { get; set; } = new();
    public List<string> RequiredFields { get; set; } = new();
    public int HeaderRow { get; set; } = 1;
}

public class TransformerConfig {
    public string Name { get; set; } = "";
    public string Function { get; set; } = "";
    public Dictionary<string, string> Parameters { get; set; } = new();
}

public class WordTemplateConfig {
    public string Name { get; set; } = "";
    public string TemplatePath { get; set; } = "";
    public string OutputPath { get; set; } = "";
    public string PlaceholderPattern { get; set; } = @"~@\w+";
    public List<string> SpecialSections { get; set; } = new();
}

public class ValidationConfig {
    public List<string> RequiredFields { get; set; } = new();
    public int MaxFieldLength { get; set; } = 1000;
    public bool ValidateNumbers { get; set; } = true;
    public bool ValidateDates { get; set; } = true;
    public List<string> AllowedDateFormats { get; set; } = new() { "MM/dd/yyyy", "yyyy-MM-dd" };
}

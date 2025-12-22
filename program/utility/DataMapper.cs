using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

public class DataMapper {
    private Dictionary<string, string> _mappingRules;
    private Dictionary<string, Func<string, string>> _transformers;
    
    public DataMapper() {
        _mappingRules = new Dictionary<string, string>();
        _transformers = new Dictionary<string, Func<string, string>>();
        InitializeDefaultMappings();
    }
    
    private void InitializeDefaultMappings() {
        // Default mapping rules for common placeholders
        _mappingRules["~@property_name"] = "Property Name";
        _mappingRules["~@property_address"] = "Address";
        _mappingRules["~@property_owner"] = "Owner";
        _mappingRules["~@building_area"] = "Building Area";
        _mappingRules["~@site_area"] = "Site Area";
        _mappingRules["~@sale_price"] = "Sale Price";
        _mappingRules["~@date_of_sale"] = "Date of Sale";
        _mappingRules["~@price_per_sf"] = "Price/SF";
        
        // Default transformers for data formatting
        _transformers["currency"] = FormatCurrency;
        _transformers["date"] = FormatDate;
        _transformers["percentage"] = FormatPercentage;
        _transformers["number"] = FormatNumber;
    }
    
    public void AddMappingRule(string placeholder, string excelColumn) {
        _mappingRules[placeholder] = excelColumn;
    }
    
    public void AddTransformer(string name, Func<string, string> transformer) {
        _transformers[name] = transformer;
    }
    
    public Dictionary<string, string> MapDataToPlaceholders(Dictionary<string, Dictionary<string, string>> excelData) {
        var result = new Dictionary<string, string>();
        
        // Process each sheet's data
        foreach (var sheetData in excelData) {
            var sheetName = sheetData.Key;
            var data = sheetData.Value;
            
            // Map data based on rules
            foreach (var rule in _mappingRules) {
                var placeholder = rule.Key;
                var columnName = rule.Value;
                
                if (data.ContainsKey(columnName)) {
                    result[placeholder] = data[columnName];
                }
            }
        }
        
        return result;
    }
    
    public Dictionary<string, string> MapTableDataToPlaceholders(string sheetName, List<Dictionary<string, string>> tableData) {
        var result = new Dictionary<string, string>();
        
        if (tableData == null || !tableData.Any()) return result;
        
        // Get all unique column names
        var allColumns = new HashSet<string>();
        foreach (var row in tableData) {
            foreach (var key in row.Keys) {
                allColumns.Add(key);
            }
        }
        
        // Create mappings for table data
        foreach (var column in allColumns) {
            var placeholder = $"~@{sheetName.ToLower()}_{column.ToLower().Replace(" ", "_").Replace("/", "_")}";
            
            // Get the first non-empty value for this column
            var value = tableData.FirstOrDefault(row => !string.IsNullOrEmpty(row.GetValueOrDefault(column, "")))?.GetValueOrDefault(column, "");
            if (!string.IsNullOrEmpty(value)) {
                result[placeholder] = value;
            }
        }
        
        return result;
    }
    
    public List<TableMapping> CreateTableMappings(List<Dictionary<string, string>> excelTableData, List<Table> wordTables) {
        var mappings = new List<TableMapping>();
        
        foreach (var table in wordTables) {
            var mapping = new TableMapping {
                WordTable = table,
                ExcelData = new List<Dictionary<string, string>>()
            };
            
            // Try to match table structure with Excel data
            var tableHeaders = ExtractTableHeaders(table);
            
            if (tableHeaders.Any()) {
                // Find matching Excel data based on headers
                var matchingData = excelTableData.Where(row => 
                    tableHeaders.Any(header => row.ContainsKey(header))).ToList();
                
                mapping.ExcelData = matchingData;
            }
            
            mappings.Add(mapping);
        }
        
        return mappings;
    }
    
    private List<string> ExtractTableHeaders(Table table) {
        var headers = new List<string>();
        
        var firstRow = table.Descendants<TableRow>().FirstOrDefault();
        if (firstRow != null) {
            var cells = firstRow.Descendants<TableCell>().ToList();
            foreach (var cell in cells) {
                var headerText = cell.InnerText.Trim();
                if (!string.IsNullOrEmpty(headerText)) {
                    headers.Add(headerText);
                }
            }
        }
        
        return headers;
    }
    
    public string TransformValue(string value, string transformerName) {
        if (_transformers.ContainsKey(transformerName)) {
            return _transformers[transformerName](value);
        }
        return value;
    }
    
    // Default transformers
    private string FormatCurrency(string value) {
        if (double.TryParse(value, out double amount)) {
            return amount.ToString("C");
        }
        return value;
    }
    
    private string FormatDate(string value) {
        if (double.TryParse(value, out double daysFrom1900)) {
            var date = DateTime.FromOADate(daysFrom1900);
            return date.ToString("MM/dd/yyyy");
        }
        return value;
    }
    
    private string FormatPercentage(string value) {
        if (double.TryParse(value, out double percentage)) {
            return (percentage / 100).ToString("P1");
        }
        return value;
    }
    
    private string FormatNumber(string value) {
        if (double.TryParse(value, out double number)) {
            return number.ToString("N0");
        }
        return value;
    }
    
    public List<string> GetAvailablePlaceholders() {
        return _mappingRules.Keys.ToList();
    }
    
    public List<string> GetAvailableTransformers() {
        return _transformers.Keys.ToList();
    }
}

public class TableMapping {
    public Table WordTable { get; set; }
    public List<Dictionary<string, string>> ExcelData { get; set; }
    
    public TableMapping() {
        ExcelData = new List<Dictionary<string, string>>();
    }
}

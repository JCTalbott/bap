using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class ValidationEngine {
    private ValidationConfig _config;
    
    public ValidationEngine(ValidationConfig config) {
        _config = config;
    }
    
    public ValidationResult ValidateData(Dictionary<string, Dictionary<string, string>> excelData, DocumentMappingConfig mappingConfig) {
        var result = new ValidationResult();
        
        // Validate required fields
        ValidateRequiredFields(excelData, mappingConfig, result);
        
        // Validate field lengths
        ValidateFieldLengths(excelData, result);
        
        // Validate numbers if enabled
        if (_config.ValidateNumbers) {
            ValidateNumbers(excelData, result);
        }
        
        // Validate dates if enabled
        if (_config.ValidateDates) {
            ValidateDates(excelData, result);
        }
        
        return result;
    }
    
    private void ValidateRequiredFields(Dictionary<string, Dictionary<string, string>> excelData, DocumentMappingConfig mappingConfig, ValidationResult result) {
        foreach (var requiredField in _config.RequiredFields) {
            bool found = false;
            
            // Check if the required field exists in any sheet
            foreach (var sheetData in excelData.Values) {
                if (sheetData.ContainsKey(requiredField) && !string.IsNullOrEmpty(sheetData[requiredField])) {
                    found = true;
                    break;
                }
            }
            
            if (!found) {
                result.AddError($"Required field '{requiredField}' not found or is empty");
            }
        }
    }
    
    private void ValidateFieldLengths(Dictionary<string, Dictionary<string, string>> excelData, ValidationResult result) {
        foreach (var sheetData in excelData) {
            foreach (var field in sheetData.Value) {
                if (field.Value.Length > _config.MaxFieldLength) {
                    result.AddWarning($"Field '{field.Key}' in sheet '{sheetData.Key}' exceeds maximum length of {_config.MaxFieldLength}");
                }
            }
        }
    }
    
    private void ValidateNumbers(Dictionary<string, Dictionary<string, string>> excelData, ValidationResult result) {
        var numericFields = new[] { "price", "cost", "amount", "area", "ratio", "percentage" };
        
        foreach (var sheetData in excelData) {
            foreach (var field in sheetData.Value) {
                var fieldName = field.Key.ToLower();
                
                if (numericFields.Any(nf => fieldName.Contains(nf))) {
                    if (!IsValidNumber(field.Value)) {
                        result.AddWarning($"Field '{field.Key}' in sheet '{sheetData.Key}' should contain a numeric value but contains: '{field.Value}'");
                    }
                }
            }
        }
    }
    
    private void ValidateDates(Dictionary<string, Dictionary<string, string>> excelData, ValidationResult result) {
        var dateFields = new[] { "date", "sale_date", "purchase_date", "effective_date" };
        
        foreach (var sheetData in excelData) {
            foreach (var field in sheetData.Value) {
                var fieldName = field.Key.ToLower();
                
                if (dateFields.Any(df => fieldName.Contains(df))) {
                    if (!IsValidDate(field.Value)) {
                        result.AddWarning($"Field '{field.Key}' in sheet '{sheetData.Key}' should contain a valid date but contains: '{field.Value}'");
                    }
                }
            }
        }
    }
    
    private bool IsValidNumber(string value) {
        if (string.IsNullOrEmpty(value)) return false;
        
        // Try parsing as double
        if (double.TryParse(value, out _)) return true;
        
        // Try parsing as currency (remove currency symbols)
        var cleanValue = value.Replace("$", "").Replace(",", "").Trim();
        return double.TryParse(cleanValue, out _);
    }
    
    private bool IsValidDate(string value) {
        if (string.IsNullOrEmpty(value)) return false;
        
        // Try parsing as DateTime
        if (DateTime.TryParse(value, out _)) return true;
        
        // Try parsing as Excel date (number of days since 1900)
        if (double.TryParse(value, out double daysFrom1900)) {
            try {
                DateTime.FromOADate(daysFrom1900);
                return true;
            } catch {
                return false;
            }
        }
        
        // Try parsing with specific formats
        foreach (var format in _config.AllowedDateFormats) {
            if (DateTime.TryParseExact(value, format, null, System.Globalization.DateTimeStyles.None, out _)) {
                return true;
            }
        }
        
        return false;
    }
    
    public ValidationResult ValidateWordDocument(WordprocessingDocument wordDoc, DocumentMappingConfig mappingConfig) {
        var result = new ValidationResult();
        
        if (wordDoc?.MainDocumentPart?.Document?.Body == null) {
            result.AddError("Word document body is null or invalid");
            return result;
        }
        
        var body = wordDoc.MainDocumentPart.Document.Body;
        
        // Check for required placeholders
        ValidateRequiredPlaceholders(body, mappingConfig, result);
        
        // Check for malformed placeholders
        ValidatePlaceholderFormat(body, result);
        
        return result;
    }
    
    private void ValidateRequiredPlaceholders(Body body, DocumentMappingConfig mappingConfig, ValidationResult result) {
        var documentText = body.InnerText;
        
        foreach (var requiredField in _config.RequiredFields) {
            if (!documentText.Contains(requiredField)) {
                result.AddWarning($"Required placeholder '{requiredField}' not found in Word document");
            }
        }
    }
    
    private void ValidatePlaceholderFormat(Body body, ValidationResult result) {
        var documentText = body.InnerText;
        var regex = new Regex(@"~@\w+");
        var matches = regex.Matches(documentText);
        
        foreach (Match match in matches) {
            var placeholder = match.Value;
            
            // Check for common placeholder format issues
            if (placeholder.Contains(" ")) {
                result.AddError($"Placeholder '{placeholder}' contains spaces - this may cause issues");
            }
            
            if (placeholder.Length > 50) {
                result.AddWarning($"Placeholder '{placeholder}' is unusually long");
            }
        }
    }
}

public class ValidationResult {
    public List<string> Errors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    
    public bool IsValid => !Errors.Any();
    
    public void AddError(string message) {
        Errors.Add(message);
    }
    
    public void AddWarning(string message) {
        Warnings.Add(message);
    }
    
    public void PrintResults() {
        if (Errors.Any()) {
            Console.WriteLine("Validation Errors:");
            foreach (var error in Errors) {
                Console.WriteLine($"  ERROR: {error}");
            }
        }
        
        if (Warnings.Any()) {
            Console.WriteLine("Validation Warnings:");
            foreach (var warning in Warnings) {
                Console.WriteLine($"  WARNING: {warning}");
            }
        }
        
        if (IsValid && !Warnings.Any()) {
            Console.WriteLine("Validation passed with no issues.");
        }
    }
}

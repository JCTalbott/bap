# Enhanced Excel to Word Data Injection System

A comprehensive C# application that intelligently extracts data from Excel spreadsheets and injects it into Word documents at relevant locations using placeholder-based mapping.

## Features

### 🔧 Core Functionality
- **Smart Data Extraction**: Automatically extracts structured data from Excel sheets
- **Intelligent Mapping**: Maps Excel data to Word document placeholders using configurable rules
- **Data Validation**: Validates both Excel data and Word templates before processing
- **Flexible Configuration**: JSON-based configuration system for custom mapping rules
- **Error Handling**: Comprehensive error handling and validation with detailed reporting

### 📊 Data Processing
- **Multi-Sheet Support**: Processes multiple Excel sheets simultaneously
- **Table Structure Recognition**: Automatically detects and processes table structures
- **Data Transformation**: Built-in transformers for currency, dates, percentages, and numbers
- **Placeholder Detection**: Automatically finds and processes placeholders in Word documents

### 🎯 Advanced Features
- **Dynamic Table Processing**: Handles complex table structures and dynamic content
- **Section Processing**: Supports section-based document organization
- **Interactive Mode**: Command-line interactive mode for configuration and testing
- **Validation Engine**: Comprehensive validation of data integrity and format

## Quick Start

### Basic Usage
```bash
# Process documents with default settings
dotnet run

# Run in interactive mode
dotnet run interactive
```

### Configuration
The system uses a JSON configuration file (`mapping-config.json`) that is automatically created on first run:

```json
{
  "ExcelSheets": {
    "Subj": {
      "Name": "Subject Property",
      "DataType": "Property",
      "Mappings": {
        "~@property_name": "Property Name",
        "~@property_address": "Address",
        "~@property_owner": "Owner"
      }
    }
  },
  "Transformers": {
    "currency": {
      "Name": "Currency",
      "Function": "FormatCurrency",
      "Parameters": { "format": "C" }
    }
  }
}
```

## Architecture

### Core Components

#### 1. ExcelDoorway (`program/abs/ExcelDoorway.cs`)
Enhanced Excel file reader with advanced data extraction capabilities:
- `ExtractSheetData()`: Extracts all cell data from a sheet
- `ExtractTableData()`: Extracts structured table data with headers
- `GetCellValue()`: Gets formatted cell values with proper type handling

#### 2. WordBeholder (`program/abs/WordBeholder.cs`)
Comprehensive Word document handler:
- `FindPlaceholders()`: Discovers placeholders using regex patterns
- `ReplaceText()`: Replaces text in paragraphs and table cells
- `GetTables()`: Extracts all tables from the document
- `Clone()`: Creates document copies for processing

#### 3. DataMapper (`program/utility/DataMapper.cs`)
Intelligent data mapping system:
- `MapDataToPlaceholders()`: Maps Excel data to Word placeholders
- `MapTableDataToPlaceholders()`: Handles table-specific mappings
- `CreateTableMappings()`: Creates mappings between Excel tables and Word tables
- Built-in transformers for common data types

#### 4. EnhancedExcelToWord (`program/EnhancedExcelToWord.cs`)
Main processing engine:
- `ProcessDocuments()`: Orchestrates the entire injection process
- `ProcessWordDocument()`: Handles Word document processing
- `ReplaceTextPlaceholders()`: Replaces simple text placeholders
- `ProcessTablePlaceholders()`: Handles complex table processing

#### 5. ConfigurationManager (`program/utility/ConfigurationManager.cs`)
Configuration management system:
- JSON-based configuration storage
- Default configuration generation
- Runtime configuration updates
- Sheet and template management

#### 6. ValidationEngine (`program/utility/ValidationEngine.cs`)
Comprehensive validation system:
- Excel data validation
- Word template validation
- Data type validation (numbers, dates, currency)
- Required field validation

### Data Flow

```
Excel File → ExcelDoorway → Data Extraction → DataMapper → 
Word Template → WordBeholder → Placeholder Processing → 
EnhancedExcelToWord → Output Document
```

## Usage Examples

### Basic Document Processing
```csharp
var enhancedProgram = new EnhancedProgram();
enhancedProgram.Run("excel-files/data.xlsx", "doc-files/template.docx");
```

### Interactive Configuration
```csharp
var enhancedProgram = new EnhancedProgram();
enhancedProgram.InteractiveMode();
```

### Custom Mapping Configuration
```csharp
var excelToWord = new EnhancedExcelToWord();
excelToWord.ConfigureMapping("~@property_name", "Property Name");
excelToWord.ConfigureMapping("~@sale_price", "Sale Price", "currency");
```

## Placeholder System

### Placeholder Format
Placeholders follow the format: `~@field_name`

Examples:
- `~@property_name` - Property name
- `~@sale_price` - Sale price (automatically formatted as currency)
- `~@date_of_sale` - Date of sale (automatically formatted as date)
- `~@building_area` - Building area (automatically formatted as number)

### Special Placeholders
- `~@table_*` - Dynamic table processing
- `~@section_*` - Section-based content processing

## Data Transformers

The system includes built-in transformers:

### Currency Transformer
```csharp
"1234.56" → "$1,234.56"
```

### Date Transformer
```csharp
"45000" (Excel date) → "03/15/2023"
```

### Percentage Transformer
```csharp
"0.125" → "12.5%"
```

### Number Transformer
```csharp
"1234567" → "1,234,567"
```

## Validation System

### Excel Data Validation
- Required fields validation
- Data type validation (numbers, dates, currency)
- Field length validation
- Format validation

### Word Template Validation
- Placeholder format validation
- Required placeholder detection
- Document structure validation

## Configuration Options

### Excel Sheet Configuration
```json
{
  "ExcelSheets": {
    "SheetName": {
      "Name": "Display Name",
      "DataType": "Property|Sales|Rental",
      "Mappings": {
        "~@placeholder": "Excel Column Name"
      },
      "RequiredFields": ["~@field1", "~@field2"],
      "HeaderRow": 1
    }
  }
}
```

### Word Template Configuration
```json
{
  "WordTemplates": {
    "template_name": {
      "Name": "Template Display Name",
      "TemplatePath": "path/to/template.docx",
      "OutputPath": "path/to/output.docx",
      "PlaceholderPattern": "~@\\w+",
      "SpecialSections": ["Executive Summary", "Analysis"]
    }
  }
}
```

### Validation Configuration
```json
{
  "Validation": {
    "RequiredFields": ["~@property_name", "~@sale_price"],
    "MaxFieldLength": 1000,
    "ValidateNumbers": true,
    "ValidateDates": true,
    "AllowedDateFormats": ["MM/dd/yyyy", "yyyy-MM-dd"]
  }
}
```

## Error Handling

The system provides comprehensive error handling:

### Validation Errors
- Missing required fields
- Invalid data formats
- Malformed placeholders

### Processing Errors
- File access issues
- Document corruption
- Mapping conflicts

### Warning System
- Data format warnings
- Missing optional fields
- Placeholder format issues

## Performance Considerations

- **Memory Management**: Uses IDisposable pattern for proper resource cleanup
- **Efficient Processing**: Processes sheets and documents in parallel where possible
- **Validation Caching**: Caches validation results to avoid redundant processing
- **Stream Processing**: Uses streaming for large file processing

## Extending the System

### Adding Custom Transformers
```csharp
var dataMapper = new DataMapper();
dataMapper.AddTransformer("custom", value => {
    // Custom transformation logic
    return transformedValue;
});
```

### Adding Custom Validation Rules
```csharp
var validationEngine = new ValidationEngine(config);
// Add custom validation logic in ValidationEngine class
```

### Adding New Data Sources
Extend the `ExcelDoorway` class to support additional data sources:
```csharp
public class CustomDataDoorway : ExcelDoorway {
    // Custom data extraction logic
}
```

## Troubleshooting

### Common Issues

1. **Missing Placeholders**: Ensure Word templates contain the required placeholders
2. **Data Format Issues**: Check Excel data formats and transformer configuration
3. **File Access Errors**: Ensure files are not locked by other applications
4. **Configuration Errors**: Validate JSON configuration file syntax

### Debug Mode
Run with detailed logging:
```bash
dotnet run --verbosity detailed
```

## Dependencies

- .NET 8.0
- DocumentFormat.OpenXml
- System.Text.Json

## License

This project is part of the open-sdk toolkit for document processing and data injection.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review validation output for specific errors
3. Check configuration file syntax
4. Verify file paths and permissions# bap

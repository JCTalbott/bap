# Quick Usage Guide

## Getting Started

### 1. Basic Processing
```bash
# Process your Excel and Word files
dotnet run
```

This will:
- Process `excel-files/pony.xlsx`
- Use `doc-files/pony.docx` as template
- Create output file with timestamp

### 2. Interactive Mode
```bash
# Run in interactive mode for configuration
dotnet run interactive
```

Available commands:
- `run` - Process documents
- `config` - Show current configuration
- `mappings` - Configure data mappings
- `placeholders` - List available placeholders
- `validate` - Validate files only
- `exit` - Exit program

## Excel File Structure

Your Excel file should have:
- **Sheet26**: Main property data
- **SSum**: Sales summary data
- **RSum**: Rental summary data
- Other sheets as needed

## Word Template Setup

Add placeholders to your Word document:
- `~@property_name` - Property name
- `~@property_address` - Property address
- `~@sale_price` - Sale price (formatted as currency)
- `~@date_of_sale` - Date of sale (formatted as date)
- `~@building_area` - Building area (formatted as number)

## Configuration

The system creates `mapping-config.json` automatically. You can edit this file to:
- Add new mappings
- Configure transformers
- Set validation rules
- Define custom templates

## Output

Processed documents are saved with timestamps:
- `doc-files/pony_processed_YYYYMMDD_HHMMSS.docx`

## Validation

The system validates:
- Required fields in Excel data
- Data formats (numbers, dates, currency)
- Word template placeholders
- File accessibility

Check console output for validation results and warnings.

## Troubleshooting

### Common Issues:
1. **Missing placeholders**: Add `~@field_name` placeholders to Word template
2. **Data format errors**: Check Excel data formats
3. **File access errors**: Close Excel/Word files before processing
4. **Configuration errors**: Check JSON syntax in `mapping-config.json`

### Getting Help:
1. Run `dotnet run interactive` and use `validate` command
2. Check console output for specific error messages
3. Review the generated `mapping-config.json` file

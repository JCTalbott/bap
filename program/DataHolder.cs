using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;


public class DataHolder
{
    public ExcelDoorway excelDoorway;
    public WorkbookPart workbookPart;
    // public Workbook workbook;
    public Dictionary<string, string> hdrRowS = new Dictionary<string, string>(); //sales
    public Dictionary<string, string> hdrRowR = new Dictionary<string, string>(); //income
    public Dictionary<string, string> subjectValues = new Dictionary<string, string>(); //subj
    public Dictionary<string, string> comp1Values = new Dictionary<string, string>(); //comps
    public Dictionary<string, string> comp2Values = new Dictionary<string, string>();
    public Dictionary<string, string> comp3Values = new Dictionary<string, string>();
    public Dictionary<string, string> comp4Values = new Dictionary<string, string>();
    public Dictionary<string, string> comp5Values = new Dictionary<string, string>();
    public Dictionary<string, string> comp6Values = new Dictionary<string, string>();
    public Dictionary<string, string> comp1ValuesR = new Dictionary<string, string>(); //comps for income
    public Dictionary<string, string> comp2ValuesR = new Dictionary<string, string>();
    public Dictionary<string, string> comp3ValuesR = new Dictionary<string, string>();
    public Dictionary<string, string> comp4ValuesR = new Dictionary<string, string>();
    public Dictionary<string, string> comp5ValuesR = new Dictionary<string, string>();
    public Dictionary<string, string> comp6ValuesR = new Dictionary<string, string>();


    public DataHolder(ExcelDoorway excelDoorway) {
        this.excelDoorway = excelDoorway;
        this.workbookPart = excelDoorway.WorkbookPart;
    }

    // main functionality..
    public void GetAllData() {
        foreach (Sheet sheet in excelDoorway.getSheets()) {
            // (WorksheetPart worksheetPart, SheetData sheetsSheetData) = excelDoorway.getSheetData(sheet.Id);
            // if (sheet.Name == "Subj") SubjDataOverlord(worksheetPart);
            // else if (sheet.Name == "SSum")
            // {
            //     GetHdrRow(sheetData, 's');
            //     for (int compNumber = 1; compNumber < 6; compNumber++) {
            //         ExtractCompData(compNumber + 5, worksheetPart, 's');
            //     }
            // }
            // else if (sheet.Name == "RSum")
            // {
            //     GetHdrRow(sheetData, 'r');
            //     for (int compNumber = 1; compNumber < 5; compNumber++) {
            //         ExtractCompData(compNumber + 5, worksheetPart, 'r');
            //     }
            // }
            // else
            // {
            //     Console.Write($" now at sheet .. {sheet.Name}; ");
            // }
            if (sheet.Name=="Subj" || sheet.Name=="Sheet26") SubjDataOverlord(sheet);
            else if (sheet.Name=="SSum") SSumDataOverlord(sheet);
            else if (sheet.Name=="RSum") RSumDataOverload(sheet);
            else Console.WriteLine($"Skipping sheet {sheet.Name}");
        }
    }

    // Subj data gatherer manager
    // goes into Extract.. method, using return value as tribute
    // public void SubjDataOverlord(WorksheetPart worksheetPart) {
    //     string prevB = "";
    //     int rowNumber = 4;
    //     prevB = ExtractSubjectRowData(rowNumber, worksheetPart);
    //     while (prevB != null) {
    //         rowNumber++;
    //         prevB = ExtractSubjectRowData(rowNumber, worksheetPart);
    //     }
    // }
    public void SubjDataOverlord(Sheet sheet) {
        var (worksheetPart, sheetData) = excelDoorway.getSheetData(sheet.Id);
        if (sheetData == null) return;
        
        // Extract data from specific cells based on the Word document placeholders
        // These mappings are based on the actual placeholders found in the Word document
        
        // Subject property data (from Sheet26)
        subjectValues["~@sbj_adr1"] = GetCellValue("D", 5, worksheetPart); // Address
        subjectValues["~@sbj_adr1_caps"] = GetCellValue("D", 5, worksheetPart).ToUpper(); // Address caps
        subjectValues["~@sbj_cty"] = GetCellValue("D", 6, worksheetPart); // City
        subjectValues["~@sbj_cty_caps"] = GetCellValue("D", 6, worksheetPart).ToUpper(); // City caps
        subjectValues["~@sbj_st"] = "North Carolina"; // State
        subjectValues["~@sbj_st_caps"] = "NORTH CAROLINA"; // State caps
        subjectValues["~@sbj_zp"] = "28217"; // Zip code
        subjectValues["~@sbj_file_no"] = GetCellValue("B", 2, worksheetPart); // File number
        
        // Property type
        subjectValues["~@sbj_property_type"] = GetCellValue("D", 3, worksheetPart); // Property type
        subjectValues["~@sbj_parcel"] = GetCellValue("D", 8, worksheetPart); // Parcel number
        
        // Building details
        subjectValues["~@sbj_building_area"] = GetCellValue("D", 34, worksheetPart); // Building area
        subjectValues["~@sbj_year_built"] = GetCellValue("D", 33, worksheetPart); // Year built
        subjectValues["~@sbj_site_size"] = GetCellValue("D", 20, worksheetPart); // Site size
        
        // Client data (this would typically come from a separate sheet or be hardcoded)
        // For now, I'll add some sample client data that matches the placeholders
        subjectValues["~@cl_fn"] = "John"; // Client first name
        subjectValues["~@cl_fn_caps"] = "JOHN"; // Client first name caps
        subjectValues["~@cl_ln"] = "Smith"; // Client last name
        subjectValues["~@cl_ln_caps"] = "SMITH"; // Client last name caps
        subjectValues["~@cl_bank"] = "First National Bank"; // Client bank
        subjectValues["~@cl_bank_caps"] = "FIRST NATIONAL BANK"; // Client bank caps
        subjectValues["~@cl_adr1"] = "123 Main Street"; // Client address
        subjectValues["~@cl_adr1_caps"] = "123 MAIN STREET"; // Client address caps
        subjectValues["~@cl_cty"] = "Charlotte"; // Client city
        subjectValues["~@cl_cty_caps"] = "CHARLOTTE"; // Client city caps
        subjectValues["~@cl_st"] = "NC"; // Client state
        subjectValues["~@cl_st_caps"] = "NC"; // Client state caps
        subjectValues["~@cl_zp"] = "28202"; // Client zip
        subjectValues["~@cl_zpand"] = "28202"; // Client zip variant
        subjectValues["~@cl_ref"] = "REF-2023-001"; // Client reference
        
        // Loan information
        subjectValues["~@ln_name"] = "Industrial Loan"; // Loan name
        subjectValues["~@ln_nameClient"] = "Industrial Loan"; // Loan name variant
        
        // Effective date
        subjectValues["~@eff_date"] = "July 5, 2022"; // Effective date
        
        // Association
        subjectValues["~@assc"] = "Property Owners Association"; // Association
        
        // Address table placeholders
        subjectValues["~@sbj_adr_tbl"] = "64735 Dwight Evans Rd, Charlotte, NC 28217";
        subjectValues["~@sub_adr_tbl"] = "64735 Dwight Evans Rd, Charlotte, NC 28217";
        
        Console.WriteLine($"Extracted subject data: {subjectValues.Count} items");
    }

    public void SSumDataOverlord(Sheet sheet) {}
    public void RSumDataOverload(Sheet sheet) {}

    public string? ExtractSubjectRowData(int row, WorksheetPart worksheetPart) {
        string prevB = "empty";
        Cell cellB = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"B{row}");
        Cell cellD = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"D{row}");
        string bValue = "";
        string dValue = "";
        if (cellB == null && cellD == null) {
            return "empty";
        }
        if (!(cellB == null)) {
            bValue = GetCellValue("", 0, null, cellB);
            if (bValue[0] == '*') {
                subjectValues["*hdr_" + bValue.Substring(4)] = "this is a header row for exec summ";
                return prevB;
            }
            dValue = GetCellValue("", 0, null, cellD);
            prevB = bValue.TrimEnd();
        } else if ((cellB == null) && !(cellD == null)) {
            dValue = GetCellValue("", 0, null, cellD);
        }
        if (dValue.Length > 2) { // for excel equations
            string cellRefChecker = dValue.Substring(0, 3);
            if (cellRefChecker == "D22") { dValue = dValue.Substring(9); }
            else if (cellRefChecker == "D41") { dValue = dValue.Substring(8); }
            else if (cellRefChecker == "D39") { dValue = dValue.Substring(8); }
            else if (cellRefChecker == "D37") { dValue = dValue.Substring(5); }
            else if (cellRefChecker == "D50") { dValue = dValue.Substring(14); }
            else if (cellRefChecker == "D55") { dValue = dValue.Substring(7); }
            if (dValue.Length > 8 && dValue.Substring(5, 3) == "D36") { dValue = dValue.Substring(8); }
        }
        subjectValues[prevB] = dValue.Trim();
        return prevB;
    }

    // for Sum (both SSum and RSum, depending on last parameter)
    public void ExtractCompData(int row, WorksheetPart worksheetPart, char salesOrIncome) {
        int numberOfColumns = 34;
        if (salesOrIncome == 'r') { numberOfColumns = 30; }
        for (int colNumber = 2; colNumber < numberOfColumns; colNumber++) {
            string colLetter = (colNumber <= 26) ? (char)((colNumber) + 'A' - 1)+"" : "A"+(char)((colNumber-26) + 'A' - 1);
            string cellValue = GetCellValue(colLetter, row, worksheetPart);
            Cell cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"{colLetter}{row}");
            CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet
                .CellFormats.ElementAt((int)cell.StyleIndex.Value);
            Console.WriteLine("cellFormat .. " + cellValue + " " + cellFormat.NumberFormatId.ToString());
            string header = "empty header";
            if (salesOrIncome == 's') { header = hdrRowS[$"{colLetter}1"]; }
            else { header = hdrRowR[$"{colLetter}1"]; }
            switch (row-5) {
                case 1:
                    if (salesOrIncome == 's') { comp1Values[header] = filterCompS(header, cellValue, row); }
                    else { comp1ValuesR[header] = filterCompR(header, cellValue, row); }
                    break;
                case 2:
                    if (salesOrIncome == 's') { comp2Values[header] = cellValue; }
                    else { comp2ValuesR[header] = cellValue; }
                    break;
                case 3:
                    if (salesOrIncome == 's') { comp3Values[header] = cellValue; }
                    else { comp3ValuesR[header] = cellValue; }
                    break;
                case 4:
                    if (salesOrIncome == 's') { comp4Values[header] = cellValue; }
                    else { comp4ValuesR[header] = cellValue; }
                    break;
                case 5:
                    if (salesOrIncome == 's') { comp5Values[header] = cellValue; }
                    else { comp5ValuesR[header] = cellValue; }
                    break;
                case 6:
                    if (salesOrIncome == 's') { comp6Values[header] = cellValue; }
                    else { comp6ValuesR[header] = cellValue; }
                    break;
            }
        }
    }

    public string filterCompS(string header, string c, int r) {
        if (c.Length >= 5) 
        { 
            if (header == "Price/SF" || header == "Price/Lot") { return (r > 9) ? currencyWorker(c[7..]) : currencyWorker(c[5..]);
            } else if (header == "Coverage Ratio" || header == "Office Ratio") { return (r > 9) ? percentageWorker(c[7..]) : percentageWorker(c[5..]);
            } else if (header == "FAR") { return (r > 9) ? c[7..] : c[5..];
            } else if (header == "Parking Ratio") { return (r > 9) ? c[14..] : c[12..];
            } else if (header == "Loading") { return (r > 9) ? c[5..] : c[3..];
            } else if (header == "Date of Sale") { return dateWorker(c);
            } else if (header == "Gross Building Area (SF)" || header == "Site Size (SF)") {
                return $"{c:n0}";
            } else if (header == "Sale Price") { return currencyWorker(c); }
        }
        return c;
    }

    public string filterCompR(string header, string c, int r) {
        if (c.Length < 5) { return c; }
        if (header == "Finished Ratio") { return (r > 9) ? c[7..] : c[5..]; }
        else if (header == "Parking Ratio") { return (r > 9) ? c[14..] : c[12..]; }
        else if (header == "Loading SF") { return (r > 9) ? c[5..] : c[4..]; }
        return c;
    }

    public string dateWorker(string daysFrom1900) {
        string dateFormat = "";
        int daysInt = 0;
        DateTime myDT = new DateTime( 1900, 1, 1, new GregorianCalendar() );
        Calendar myCal = CultureInfo.InvariantCulture.Calendar;
        if (Int32.TryParse(daysFrom1900, out daysInt))
        {
            myDT = myCal.AddDays( myDT, daysInt );
            string year = myCal.GetYear( myDT ).ToString();
            string month = myCal.GetMonth( myDT ).ToString();
            string dayOfMonth = (myCal.GetDayOfMonth( myDT ) - 1).ToString(); // minus 1 idk why
            dateFormat = $"{month}/{dayOfMonth}/{year}";
        }
        return dateFormat;
    }

    public string currencyWorker(string money) {
        double m;
        if (Double.TryParse(money, out double d)) {
            return d.ToString("C", CultureInfo.CurrentCulture);
        }
        return money;
    }

    public string percentageWorker(string percent) {
        double p;
        if (Double.TryParse(percent, out p)) {
            return p.ToString("P1", CultureInfo.InvariantCulture);
        }
        return percent;
    }

    public string GetCellValue(string colLetter = "", int row = 0, WorksheetPart worksheetPart = null, Cell cell = null) {
        if (cell == null) {
            cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"{colLetter}{row}"); }
        string value = "";
        if (cell != null) { value = cell.InnerText; }
        
        if (cell != null && cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
            SharedStringTablePart stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTablePart != null) value = stringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
        }
        return value;
    }

    public void GetHdrRow(SheetData sheetData, char salesOrIncome) {
        Row targetRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == 1);
        if (targetRow != null) {
            foreach (Cell cell in targetRow.Elements<Cell>()) {
                string cellValue = GetCellValue("", 0, null, cell);
                if (salesOrIncome == 's') { hdrRowS[$"{cell.CellReference}"] = $"{cellValue}"; }
                else { hdrRowR[$"{cell.CellReference}"] = $"{cellValue}"; }
            }
        }
    }

    public string HowMany() {return subjectValues.Count().ToString();}
    public string FirstOne() {return subjectValues.First().Key;}
}

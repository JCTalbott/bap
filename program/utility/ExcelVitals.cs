using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// this is where printing vitals for Excel go
class ExcelVitals {
  // this will check the data within the dataholder (from Excel data)
  public static void printData(DataHolder dh) {
    foreach (var i in dh.subjectValues) {
        Console.Write($"{i.Key} EQUALS {i.Value} ... ");
    }
    Console.WriteLine("");
    foreach (var i in dh.comp3Values) {
        Console.Write($"{i.Key} EQUALS {i.Value} ... ");
    }
    Console.WriteLine("");
    foreach (var i in dh.comp1ValuesR) {
        Console.Write($"{i.Key} EQUALS {i.Value} ... ");
    }
  }


}
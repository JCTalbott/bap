// using System;

// class Program {
//     static void Main(string[] args) {
//         Console.WriteLine("Excel to Word Data Injection System");
//         Console.WriteLine("=====================================");
        
//         if (args.Length > 0 && args[0].ToLower() == "proper") {
//             // Use the proper system with actual placeholders
//             RunProperSystem();
//         } else {
//             var enhancedProgram = new EnhancedProgram();
            
//             if (args.Length > 0 && args[0].ToLower() == "interactive") {
//                 enhancedProgram.InteractiveMode();
//             } else {
//                 // Default processing
//                 enhancedProgram.Run();
//             }
//         }
//     }
    
//     static void RunProperSystem() {
//         Console.WriteLine("Running Proper Excel to Word System");
//         Console.WriteLine("===================================");
        
//         using (ExcelDoorway excelDoorway = new ExcelDoorway("excel-files/pony.xlsx", false)) {
//             try {
//                 DataHolder dh = new DataHolder(excelDoorway);
//                 dh.GetAllData();
                
//                 ProperExcelToWord properEtw = new ProperExcelToWord();
//                 properEtw.ProcessDocuments(dh);
                
//                 Console.WriteLine("\nProcessing completed successfully!");
//             } catch (Exception ex) {
//                 Console.WriteLine($"Error: {ex.Message}");
//                 throw new ApplicationException("Something wrong happened in the module:", ex);
//             }
//         }
//     }
// }

// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Wordprocessing;
// using System;
// using System.Linq;

// class Program
// {
//     static void Main(string[] args)
//     {
//         // Path to your Word document
//         string filePath = @"doc-files/good.docx";

//         try
//         {
//             // Open the document as Read-Only (false)
//             using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
//             {
//                 // 1. Initialize the TableFinder with the MainDocumentPart
//                 var finder = new TableFinder(doc.MainDocumentPart);

//                 // 2. Identify the table after a specific heading
//                 string searchHeading = "Comparable Improved Sales Adjustment Grid"; // Change this to your heading
//                 Table tableElement = finder.GetTableAfterHeading(searchHeading);

//                 if (tableElement != null)
//                 {
//                     // 3. Extract the structured data
//                     var tableInfo = finder.GetTableInfo(tableElement);

//                     // 4. Print the results to the Console
//                     Console.WriteLine($"--- Table Found after '{searchHeading}' ---");
//                     Console.WriteLine($"Style Used: {tableInfo.StyleId ?? "Default Table Style"}");
//                     Console.WriteLine(new string('-', 50));

//                     foreach (var row in tableInfo.Rows)
//                     {
//                         // Visually distinguish header rows
//                         string prefix = row.IsHeader ? "[HEADER] " : "         ";
                        
//                         // Join cell text with a pipe separator
//                         var cellTexts = row.Cells.Select(c => 
//                             c.ColSpan > 1 ? $"{c.Text} (Merged {c.ColSpan})" : c.Text);
                        
//                         Console.WriteLine($"{prefix} | {string.Join(" | ", cellTexts)} |");
//                     }
//                 }
//                 else
//                 {
//                     Console.WriteLine($"Could not find a table following the heading: '{searchHeading}'");
//                 }
//             }
//         }
//         catch (System.IO.FileNotFoundException)
//         {
//             Console.WriteLine("Error: The file path provided was not found.");
//         }
//         catch (Exception ex)
//         {
//             Console.WriteLine($"An error occurred: {ex.Message}");
//         }

//         Console.WriteLine("\nPress any key to exit...");
//         Console.ReadKey();
//     }
// }



using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;

// class Program
// {
//     static void Main(string[] args)
//     {
//         WordInjector wi = new WordInjector("doc-files/good.docx", "doc-files/dump/generated_report.docx");
//         wi.Run();
//     }
// }

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

class Program
{
    static void printNearby(string body, string finder, int back, int lngth, int instance) {
        int startingSpot = -1;
        for (int i = 0; i < instance; i++)
        {
            startingSpot = body.IndexOf(finder, startingSpot+1);
        }
        Console.WriteLine(body.Substring(startingSpot - back, lngth));
    }

    static void Main(string[] args)
    {
        string templatePath = "doc-files/test.docx";
        string outputPath = "appraisal_report.docx";

        byte[] templateBytes = File.ReadAllBytes(templatePath);
        using (MemoryStream ms = new MemoryStream())
        {
            ms.Write(templateBytes, 0, templateBytes.Length);
            ms.Position = 0;

            using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
            {
                var mainPart = doc.MainDocumentPart!;
                var bodyString = WordVitals.getDocString(doc);
                Console.WriteLine(bodyString.IndexOf("Report Criteria",311395));
                printNearby(bodyString, "Sales Comparison Approach", -4000, 80000, 5);
                // var imageEngine = new WordImageEngine(mainPart);
                // var imageTemplate = imageEngine.FindImageByAltText("TemplatePhoto1");
                
                // if (imageTemplate != null)
                // {
                //     var photos = new List<PhotoInput>
                //     {
                //         new PhotoInput { Path = "pics/image_rId97.png", Description = "Subject Front" },
                //         new PhotoInput { Path = "pics/image_rId100.png", Description = "Subject Rear" },
                //         new PhotoInput { Path = "pics/image_rId17.png", Description = "Kitchen" },
                //         new PhotoInput { Path = "pics/image_rId18.png", Description = "Street View" }
                //     };
                //     imageEngine.GeneratePhotoPages(imageTemplate, photos);
                // }
                
                // var tableEngine = new WordTableEngine(mainPart);
                // var firstTable = tableEngine.FindTableAfterHeading("Executive Summary");
                // if (firstTable != null)
                // {
                //     var model = tableEngine.ExtractModel(firstTable);
                //     PrintTableInfo(model);
                // }

                doc.Save();
            }

            File.WriteAllBytes(outputPath, ms.ToArray());
        }
    }

    static void PrintTableInfo(WordTableModel model)
    {
        Console.WriteLine("--- Table Metadata ---");
        Console.WriteLine($"Style ID:    {model.TableStyleId ?? "None"}");
        Console.WriteLine($"Total Width: {model.TableWidth?.ToString() ?? "Auto"}");
        Console.WriteLine($"Width Type:  {model.WidthType?.ToString() ?? "N/A"}");
        Console.WriteLine($"Columns:     {model.ColumnCount}");
        
        Console.WriteLine("\n--- Column Widths (Twips) ---");
        for (int i = 0; i < model.ColumnWidths.Count; i++)
        {
            Console.WriteLine($"Column {i + 1}: {model.ColumnWidths[i]?.ToString() ?? "Unknown"}");
        }
    }
}



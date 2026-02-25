using DocumentFormat.OpenXml.Packaging;

public class WordInjector
{
    private readonly String from;
    private readonly String to;

    public WordInjector(string from, string to)
    {
        this.from = from;
        this.to = to;
    }

    public void Run() {
        byte[] templateBytes = File.ReadAllBytes(from);
        using (MemoryStream ms = new MemoryStream())
        {
            ms.Write(templateBytes, 0, templateBytes.Length);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
            {
                var mainPart = doc.MainDocumentPart;
                var engine = new WordTableEngine(mainPart);
                var referenceTable = engine.FindTableAfterHeading("Comparable Improved Sales Adjustment Grid");
                if (referenceTable != null)
                {
                    var model = engine.ExtractModel(referenceTable);
                    var data = new List<Dictionary<string, string>>
                    {
                        new Dictionary<string, string> { 
                            { "Address", "123 Maple St" }, 
                            { "Price", "$450,000" } 
                        },
                        new Dictionary<string, string> { 
                            { "Address", "456 Oak Ave" }, 
                            { "Price", "$475,000" } 
                        }
                    };
                    var newTable = engine.BuildTableFromModel(model, data);
                    engine.InsertAfter(referenceTable, newTable, pageBreakFirst: true);
                }
                else
                {
                    Console.WriteLine("Could not find the heading or the table following it.");
                }
                doc.Save();
            }
            File.WriteAllBytes(to, ms.ToArray());
            Console.WriteLine($"File saved to: {to}");
        }
    }
}
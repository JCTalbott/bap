using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class NameInjector {

#pragma warning disable CS8602
  static void swapString(string n, string o) {
    // using WordprocessingDocument document = WordprocessingDocument.Open("report.docx", true);
    // var newDoc = (WordprocessingDocument)document.Clone();
    File.Copy("report.docx", "report-swap.docx");
    using WordprocessingDocument document = WordprocessingDocument.Open("report-swap.docx", isEditable: true); {
    var body = document.MainDocumentPart.Document.Body;
    var paragraphs = body.Elements<Paragraph>();

    foreach (var paragraph in paragraphs) {
        var runs = paragraph.Elements<Run>();
        foreach (var run in runs) {
            Console.WriteLine("hi");
            var text = run.Elements<Text>().FirstOrDefault(t => t.Text.Contains(o));
            if (text != null) text.Text = text.Text.Replace(o, n); }  }
    }
    // document.MainDocumentPart.Document.Save();
  }

  private static void findTagLocation(string tag) {

  }

  static void main() {
    swapString("Mr.", "Mr.");
  }

#pragma warning restore CS8602
}
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.Linq;
using System.Globalization;

class ExcelToWord {
  string bodyXml = "";
  Body? body;
  Dictionary<int, string> subjIndices = new Dictionary<int, string>();
  string[] testKeywords = { "Property", "Address", "Owner", "Building", "Area", "Price", "Sale", "Date" };

  private void printNearby(string body, string finder, int back, int lngth) {
    int i = body.IndexOf(finder);
    Console.WriteLine(body.Substring(i - back, lngth));
  }

  private void saveImages(MainDocumentPart m) {
        Body body = m.Document.Body;
        var imageParts = m.ImageParts;

        // Iterate through each image part
        foreach (var imagePart in imageParts)
        {
            string fileName = $"try_image_{m.GetIdOfPart(imagePart)}.png";
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                imagePart.GetStream().CopyTo(fs);
            }
        }
  }


  private String RegisterImage(MainDocumentPart m, String imagePath) {
    ImagePart imagePart = m.AddImagePart(ImagePartType.Png);
    using (FileStream stream = new FileStream(imagePath, FileMode.Open))
    {
        imagePart.FeedData(stream);
        Console.WriteLine(imagePart.Uri);
        return m.GetIdOfPart(imagePart);
    }
  }

  private void ChangeImage(Body b, String rId, String newId) {
    var blip = b.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault(b => b.Embed == rId);
    if (blip != null) {
      blip.Embed = newId;
    }
  }


  public void DoIt(DataHolder dh) {
    using (WordprocessingDocument report = WordprocessingDocument.Open("doc-files/report.docx", true))
    {
        WordprocessingDocument newPlace = report.Clone("doc-files/newPlace.docx");
        MainDocumentPart mainPart = newPlace.MainDocumentPart;
        body = mainPart.Document.Body;
        printNearby(body.InnerXml, "~@10", 250, 500);

        String rId = "";
        rId = RegisterImage(mainPart, "image_rId15.png");
        ChangeImage(body, "rId11", rId);
        rId = RegisterImage(mainPart, "image_rId16.png");
        ChangeImage(body, "rId12", rId);
        // foreach (var i in mainPart.ImageParts) { // we have image 41,2 in media..
        //   Console.WriteLine(i.Uri);
        // }
        
        body = ExecSummaryPage(body, dh);
        // body = ImprovedSalesGrid(body, dh);
        newPlace.MainDocumentPart.Document.Body = body;
        newPlace.Save();

        // bodyXml = body.OuterXml;
        //Paragraph? execSummTableHeading = body.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains("ESTitle"));
        //Run? execSummRowHeading = body.Descendants<Run>().FirstOrDefault(r => r.InnerText.Contains("esRow"));
        //Console.WriteLine(targetParagraph.InnerXml);
        //Console.WriteLine(targetRow.InnerXml);
        //printNearby(bodyXml, "ESTitle", 100, 2000);
        
        // int i = 0;
        // foreach (var d in body.Descendants()) {
        //   i++;
        //   if (i>100) {break;}
        //   Console.WriteLine(d.NamespaceUri);
        // }
        // int i = body.IndexOf("~@sbj_crnt_ownr");
        // int i1 = body.IndexOf("Northeast");
        // int i2 = body.IndexOf("Facing Southwest");
        // int e = body.IndexOf("Email");
        // int i4 = body.IndexOf("row1hdr");
        // Console.WriteLine(body.Substring(i4-700, 1500));
        // int i5 = body.IndexOf("prop1wtf");
        // Console.WriteLine(body.Substring(i5-800, 1200)); // idea : clone a row and header row and just sub text in as template for dyn # of rows .. 
        // Console.WriteLine(i1.ToString() + " ~~~ " + body.Substring(i1 - 700, 1000));
        // Console.WriteLine(i2.ToString() + "~~~ " + body.Substring(i2 - 800, 1200));
        // Console.WriteLine(e.ToString() + "~~~ " + body.Substring(e+200, 4000));
        // var imagePart = docClone.MainDocumentPart.ImageParts.FirstOrDefault(ip => docClone.MainDocumentPart.GetIdOfPart(ip) == "182"); //rId12
        // var imagePart2 = docClone.MainDocumentPart.ImageParts.FirstOrDefault(ip => docClone.MainDocumentPart.GetIdOfPart(ip) == "rId12"); //13
        // Console.WriteLine(imagePart2.Uri.ToString());

        // ImagePart part = imageClone.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
        // string imageId = imageClone.MainDocumentPart.GetIdOfPart(part);
        // Console.WriteLine("newly added image id : " + imageId);

        // using (System.IO.FileStream stream = new System.IO.FileStream("images/image3", System.IO.FileMode.Open)) {
        //     part.FeedData(stream);
        // }
        // mainPart.Document.Body = new Body(body.Replace("~@sbj_crnt_ownr", "Investicore Prop Co 9, LLC")); 
        // printNearby(body, "salesCompTitle1", 2000, 2600);
        // printNearby(body, "salesCompTitle2", 2000, 2600);
        // printNearby(body, "subjCol", 2000, 2600);
        // printNearby(body, "scPartitioner", 2000, 2600);
        //printNearby(body, "sbjtbl1hdr", 1500, 4500);
        // printNearby(body, "asIsWording", 1800, 2400);
        
        // subject();
        // foreach (var si in subjIndices) { // replaces certain strings in body
        //   string replacement = si.Value;
        //   int lengthToReplace = replacement.Length;
        //   int startIndex = si.Key;
        //   StringBuilder sb = new StringBuilder(body);
        //   sb.Remove(startIndex, lengthToReplace);
        //   sb.Insert(startIndex, replacement);
        //   body = sb.ToString();
        // }
        // string replace = body.Replace("[prc]", "Atypical");
        // string replace_2 = body.Replace("McKoy", "Talbott");
        // mainPart.Document.Body = new Body(replace_2);
        //docClone.Save();
    }}


  private Body ExecSummaryPage(Body body, DataHolder dh) {
    Paragraph firstTitle = body.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains("~@11"));
    Paragraph rowTitleSource = firstTitle;
    firstTitle = (Paragraph) SetExecSummSectionHeader(firstTitle, dh.subjectValues.First().Key);
    OpenXmlElement nodePointer = firstTitle.NextSibling();
    Paragraph emptyParagraph = (Paragraph) nodePointer;

    Table? tableSource = (Table?) body.Descendants<Table>().FirstOrDefault(t => t?.Descendants<TableRow>()?.FirstOrDefault()?.InnerText.Contains("Templateprop") == true);
    TableRow? tableRowSource = (TableRow?) tableSource?.Elements<TableRow>().FirstOrDefault();
    TableRow? emptyTableRow = (TableRow?) tableRowSource?.NextSibling();

    bool firstOverallPick = true;
    bool firstRowInTable = true;
    Table prevTable; // for the node pointer after we are done with adding rows
    foreach (var row in dh.subjectValues) {
      string prop = row.Key;
      string value = row.Value;
      if (prop[0] == '*') {
        firstRowInTable = true; // for next table
        if (!firstOverallPick) {
          Paragraph newHeader = (Paragraph) rowTitleSource.CloneNode(true);
          newHeader = (Paragraph) SetExecSummSectionHeader(newHeader, prop);
          body.InsertAfter(newHeader, nodePointer);
          body.InsertAfter(emptyParagraph.CloneNode(true), newHeader);
          nodePointer = emptyParagraph; // the space under the heading
        } else {
          Run run = new Run(new Break());
          firstOverallPick = false; continue; // cuz first one is done already
        }
      } else {
        if (firstRowInTable) {
          firstRowInTable = false;
          Table newTable = (Table) tableSource.CloneNode(true);
          TableRow newRow = (TableRow) newTable.Descendants<TableRow>().FirstOrDefault();
          newRow.InnerXml = SetNewRowXml(newRow, prop, value);
          body.InsertAfter(newTable, nodePointer);
          nodePointer = newRow.Parent;
          prevTable = newTable; // this is the current table, but will be the prev table
        } else {
          TableRow newRow = (TableRow) tableRowSource.CloneNode(true);
          newRow.InnerXml = SetNewRowXml(newRow, prop, value);
          nodePointer.AppendChild(newRow);
          nodePointer.AppendChild(emptyTableRow.CloneNode(true));
          //body.InsertAfter(newRow, nodePointer);
          //Console.WriteLine("new row : " + nodePointer.OuterXml);
        }
      }
    }
    //printNearby(body.InnerXml, "Property Info", 1000, 5000); (exec summary table area)
    return body;
  }

  private Body ImprovedSalesGrid(Body body, DataHolder dh) {
    printNearby(body.InnerXml, "~@12", 5, 2500);
    body = new Body(body.OuterXml.Replace("poopy", "now we gon make this shie loooooonggg"));
    return body;
  }

  private Body SalesAdjustmentGrid(Body body, DataHolder dh) {
    return body;
  }

  // helpers 
  private OpenXmlElement SetExecSummSectionHeader(OpenXmlElement node, string header) {
    header = header.Substring(5).Replace("_", " ");
    header = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(header.ToLower());
    node.InnerXml = node.InnerXml.Replace(node.InnerText, header);
    return node;
  }

  private string SetNewRowXml(TableRow r, string p, string v) {
    r.InnerXml = r.InnerXml.Replace("Templateprop", p);
    r.InnerXml = r.InnerXml.Replace("Propdata", v);
    return r.InnerXml;
  }

  private void subject() {
      // need indices for subject .. one time so -> instance variable
      FindIndicesOfSubject(bodyXml);
  }

  private void FindIndicesOfSubject(string body) {
      string wordBank = makeKeywords(); // testing right now (in this function)
      Regex regex = new Regex(wordBank, RegexOptions.IgnoreCase);
      MatchCollection matches = regex.Matches(body);
      foreach (Match m in matches) {
        subjIndices[m.Index] = m.Value;
        Console.WriteLine($"val: {m.Value} and ind.: {m.Index}");
      }
  }

  private string makeKeywords() {
    StringBuilder patternBuilder = new StringBuilder();
      foreach (string str in testKeywords) {
          patternBuilder.Append("\\b"); // for more than 1-worded phrases 
          patternBuilder.Append(Regex.Escape(str));
          patternBuilder.Append("\\b");
          patternBuilder.Append('|');
      }
      patternBuilder.Length--;
      return patternBuilder.ToString();
  }

}




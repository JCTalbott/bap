using DocumentFormat.OpenXml.Packaging;

class WordVitals {

  // returns the body of the Word doc in string format. Used "?" for null-safety, tho not important (maybe)
  public static string? getDocString(WordprocessingDocument doc) {
    return doc.MainDocumentPart?.Document.Body?.InnerXml;
  }


  // will print parts of the word doc around your index, with specified "range"..
  public static void printNearby(string body, string finder, int back, int lngth) {
    int i = body.IndexOf(finder);
    Console.WriteLine(body.Substring(i - back, lngth));
  }


}
using Spire.Doc;

namespace AIWorkAssistant.Services.HkOrder;

public static class DocReaderService
{
    public static string ReadText(string filePath)
    {
        var doc = new Document();
        doc.LoadFromFile(filePath);
        return doc.GetText();
    }
}

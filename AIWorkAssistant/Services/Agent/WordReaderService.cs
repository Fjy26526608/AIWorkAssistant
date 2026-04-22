using System.IO;
using System.Text;
using NPOI.XWPF.UserModel;

namespace AIWorkAssistant.Services.Agent;

public static class WordReaderService
{
    public static string ReadDocument(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".doc" => ReadDoc(filePath),
            ".docx" => ReadDocx(filePath),
            _ => throw new NotSupportedException($"不支持的文件格式: {ext}")
        };
    }

    private static string ReadDoc(string filePath)
    {
        var document = new Spire.Doc.Document();
        document.LoadFromFile(filePath);
        return document.GetText();
    }

    private static string ReadDocx(string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        var doc = new XWPFDocument(fs);
        var sb = new StringBuilder();

        foreach (var para in doc.Paragraphs)
            sb.AppendLine(para.Text);

        foreach (var table in doc.Tables)
        {
            foreach (var row in table.Rows)
            {
                var cells = row.GetTableCells().Select(c => c.GetText()).ToList();
                sb.AppendLine(string.Join("\t", cells));
            }
        }

        return sb.ToString();
    }
}

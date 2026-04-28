using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Spire.Doc;

namespace AIWorkAssistant.Services.HkOrder;

public static class OrderFileReaderService
{
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".doc", ".docx", ".xls", ".xlsx"
    };

    public static bool IsSupported(string filePath) => SupportedExtensions.Contains(Path.GetExtension(filePath));

    public static string[] GetOrderFiles(string folder) => Directory
        .EnumerateFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
        .Where(IsSupported)
        .OrderBy(f => f)
        .ToArray();

    public static string ReadText(string filePath)
    {
        return Path.GetExtension(filePath).ToLowerInvariant() switch
        {
            ".doc" or ".docx" => ReadWordText(filePath),
            ".xls" or ".xlsx" => ReadExcelText(filePath),
            var ext => throw new NotSupportedException($"不支持的订单文件类型：{ext}")
        };
    }

    private static string ReadWordText(string filePath)
    {
        var doc = new Document();
        doc.LoadFromFile(filePath);
        return doc.GetText();
    }

    private static string ReadExcelText(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        IWorkbook workbook = Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? new XSSFWorkbook(stream)
            : new HSSFWorkbook(stream);

        var sb = new StringBuilder();
        var formatter = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("zh-CN"));

        for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            var sheet = workbook.GetSheetAt(sheetIndex);
            if (sheet == null) continue;

            sb.AppendLine($"# 工作表：{sheet.SheetName}");
            for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                var cells = new List<string>();
                for (var cellIndex = row.FirstCellNum < 0 ? 0 : row.FirstCellNum; cellIndex < row.LastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    var value = cell == null ? string.Empty : formatter.FormatCellValue(cell).Trim();
                    cells.Add(value.Replace("\r", " ").Replace("\n", " "));
                }

                if (cells.Any(v => !string.IsNullOrWhiteSpace(v)))
                {
                    sb.AppendLine($"R{rowIndex + 1}: {string.Join(" | ", cells)}");
                }
            }

            sb.AppendLine();
        }

        return sb.ToString();
    }
}

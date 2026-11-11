using System.Text;
using ClosedXML.Excel;

const int ExcelMaxCellChars = 32767;
const int ExcelMaxRows = 1_048_576;

string? encodingName = null;
var positional = new List<string>();

foreach (var arg in args)
{
    if (arg.StartsWith("--encoding=", StringComparison.OrdinalIgnoreCase))
    {
        encodingName = arg["--encoding=".Length..];
    }
    else if (arg is "-h" or "--help" or "/?")
    {
        PrintUsage();
        return 0;
    }
    else if (arg.StartsWith("--"))
    {
        Console.Error.WriteLine($"Unknown option: {arg}");
        PrintUsage();
        return 1;
    }
    else
    {
        positional.Add(arg);
    }
}

if (positional.Count != 2)
{
    PrintUsage();
    return 1;
}

string fileA = positional[0];
string fileB = positional[1];

Encoding encoding;
if (string.IsNullOrWhiteSpace(encodingName))
{
    encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
}
else
{
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    try
    {
        encoding = int.TryParse(encodingName, out int codePage)
            ? Encoding.GetEncoding(codePage)
            : Encoding.GetEncoding(encodingName);
    }
    catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
    {
        Console.Error.WriteLine($"Unknown encoding: {encodingName}");
        return 1;
    }
}

List<string> linesA, linesB;
try
{
    linesA = ReadLines(fileA, encoding);
    linesB = ReadLines(fileB, encoding);
}
catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
{
    Console.Error.WriteLine($"Error reading input: {ex.Message}");
    return 1;
}

int rowCount = Math.Max(linesA.Count, linesB.Count);
if (rowCount + 1 > ExcelMaxRows)
{
    Console.Error.WriteLine($"Too many lines ({rowCount}): Excel supports at most {ExcelMaxRows - 1} data rows.");
    return 1;
}

string outputName = $"{Path.GetFileNameWithoutExtension(fileA)}_{Path.GetFileNameWithoutExtension(fileB)}.xlsx";
string outputPath = Path.Combine(Directory.GetCurrentDirectory(), outputName);

int truncated = 0;
using (var workbook = new XLWorkbook())
{
    var ws = workbook.Worksheets.Add("Compare");

    ws.Cell(1, 1).Value = "#";
    ws.Cell(1, 2).Value = Path.GetFileName(fileA);
    ws.Cell(1, 3).Value = Path.GetFileName(fileB);
    ws.Row(1).Style.Font.Bold = true;
    ws.SheetView.FreezeRows(1);

    for (int i = 0; i < rowCount; i++)
    {
        int row = i + 2;
        if (i < linesA.Count)
        {
            ws.Cell(row, 1).Value = i + 1;
            ws.Cell(row, 2).Value = Clip(linesA[i]);
        }
        if (i < linesB.Count)
        {
            ws.Cell(row, 3).Value = Clip(linesB[i]);
        }
    }

    ws.Columns(2, 3).Style.Font.FontName = "Consolas";
    ws.Columns(2, 3).Style.Alignment.WrapText = false;
    ws.Columns(2, 3).Width = 80;
    ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
    ws.Column(1).AdjustToContents();
    if (ws.Column(1).Width < 6) ws.Column(1).Width = 6;

    try
    {
        workbook.SaveAs(outputPath);
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
    {
        Console.Error.WriteLine($"Error writing output: {ex.Message}");
        return 1;
    }
}

if (truncated > 0)
    Console.Error.WriteLine($"Warning: {truncated} line(s) exceeded {ExcelMaxCellChars} characters and were truncated.");

Console.WriteLine($"Written {outputPath} ({linesA.Count} / {linesB.Count} lines)");
return 0;

string Clip(string s)
{
    if (s.Length <= ExcelMaxCellChars) return s;
    truncated++;
    return s[..ExcelMaxCellChars];
}

static List<string> ReadLines(string path, Encoding encoding)
{
    var lines = new List<string>();
    using var reader = new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true);
    string? line;
    while ((line = reader.ReadLine()) != null)
        lines.Add(line);
    return lines;
}

static void PrintUsage()
{
    Console.Error.WriteLine("Usage: SW059_ExcelComparer <filea> <fileb> [--encoding=<name>]");
    Console.Error.WriteLine("  Writes <filea>_<fileb>.xlsx to the current directory, with the lines of both files side by side.");
    Console.Error.WriteLine("  --encoding  Encoding of both input files (default: UTF-8, BOM detected). E.g. windows-1252, iso-8859-1, 1252.");
}

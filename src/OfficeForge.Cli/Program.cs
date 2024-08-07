using System.CommandLine;
using OfficeForge;
using OfficeForge.Export;
using OfficeForge.Models;
using OfficeForge.Templates;

var root = new RootCommand("Create, edit and convert Word/Excel/PowerPoint documents without Microsoft Office.");

var fileArg = new Argument<FileInfo>("file", "Path to the document");
var cellArg = new Argument<string>("cell", "Cell reference in A1 notation");
var sheetOption = new Option<string?>("--sheet", "Worksheet name (defaults to the first sheet)");

var readCell = new Command("read-cell", "Read a single cell from an .xlsx workbook");
readCell.AddArgument(fileArg);
readCell.AddArgument(cellArg);
readCell.AddOption(sheetOption);
readCell.SetHandler((FileInfo file, string cell, string? sheetName) =>
{
    var workbook = OfficeDocument.OpenWorkbook(file.FullName);
    var sheet = ResolveSheet(workbook, sheetName);
    Console.WriteLine(sheet[cell].ToString());
}, fileArg, cellArg, sheetOption);
root.AddCommand(readCell);

var valueArg = new Argument<string>("value", "Value to write");
var writeCell = new Command("write-cell", "Write a single cell in an .xlsx workbook");
writeCell.AddArgument(fileArg);
writeCell.AddArgument(cellArg);
writeCell.AddArgument(valueArg);
writeCell.AddOption(sheetOption);
writeCell.SetHandler((FileInfo file, string cell, string value, string? sheetName) =>
{
    var workbook = file.Exists ? OfficeDocument.OpenWorkbook(file.FullName) : new WorkbookModel();
    var sheet = workbook.Sheets.Count == 0
        ? workbook.AddSheet(sheetName ?? "Sheet1")
        : ResolveSheet(workbook, sheetName);
    sheet[cell] = ParseValue(value);
    OfficeDocument.SaveWorkbook(workbook, file.FullName);
}, fileArg, cellArg, valueArg, sheetOption);
root.AddCommand(writeCell);

var formatOption = new Option<ExportFormat>("--format", () => ExportFormat.Text, "Output format");
var outputOption = new Option<FileInfo?>("--output", "Output file (stdout when omitted)");
var convert = new Command("convert", "Convert a document to text, markdown or json");
convert.AddArgument(fileArg);
convert.AddOption(formatOption);
convert.AddOption(outputOption);
convert.SetHandler((FileInfo file, ExportFormat format, FileInfo? output) =>
{
    var result = OfficeDocument.Export(file.FullName, format);
    if (output is null) Console.Write(result);
    else File.WriteAllText(output.FullName, result);
}, fileArg, formatOption, outputOption);
root.AddCommand(convert);

var extractText = new Command("extract-text", "Extract plain text from a document");
extractText.AddArgument(fileArg);
extractText.SetHandler((FileInfo file) =>
    Console.Write(OfficeDocument.Export(file.FullName, ExportFormat.Text)), fileArg);
root.AddCommand(extractText);

var templateOutputArg = new Argument<FileInfo>("output", "Path for the filled copy");
var setOption = new Option<string[]>("--set", "Placeholder values as key=value") { AllowMultipleArgumentsPerToken = true };
var fillTemplate = new Command("fill-template", "Fill {{placeholders}} in a .docx or .xlsx template");
fillTemplate.AddArgument(fileArg);
fillTemplate.AddArgument(templateOutputArg);
fillTemplate.AddOption(setOption);
fillTemplate.SetHandler((FileInfo file, FileInfo output, string[] pairs) =>
{
    var filler = new TemplateFiller(TemplateFiller.ParsePairs(pairs));
    switch (OfficeDocument.DetectKind(file.FullName))
    {
        case DocumentKind.Workbook:
            var workbook = OfficeDocument.OpenWorkbook(file.FullName);
            filler.Fill(workbook);
            OfficeDocument.SaveWorkbook(workbook, output.FullName);
            break;
        case DocumentKind.Document:
            var document = OfficeDocument.OpenDocument(file.FullName);
            filler.Fill(document);
            OfficeDocument.SaveDocument(document, output.FullName);
            break;
        default:
            throw new NotSupportedException("Template filling supports .docx and .xlsx files.");
    }
}, fileArg, templateOutputArg, setOption);
root.AddCommand(fillTemplate);

return await root.InvokeAsync(args);

static SheetModel ResolveSheet(WorkbookModel workbook, string? name)
{
    if (name is null)
        return workbook.Sheets.FirstOrDefault() ?? throw new InvalidOperationException("Workbook has no sheets.");
    return workbook.FindSheet(name) ?? throw new InvalidOperationException($"Sheet '{name}' not found.");
}

static CellValue ParseValue(string value)
{
    if (value.StartsWith('=')) return CellValue.FromFormula(value[1..]);
    if (bool.TryParse(value, out var b)) return CellValue.FromBoolean(b);
    if (double.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var n)) return CellValue.FromNumber(n);
    return CellValue.FromText(value);
}

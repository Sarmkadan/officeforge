using System.Globalization;

namespace OfficeForge.Models;

public enum CellValueKind
{
    Empty,
    Text,
    Number,
    Boolean,
    DateTime,
    Formula
}

public readonly record struct CellValue
{
    public CellValueKind Kind { get; }
    public string? Text { get; }
    public double Number { get; }
    public bool Boolean { get; }
    public DateTime DateTime { get; }
    public string? Formula { get; }

    private CellValue(CellValueKind kind, string? text = null, double number = 0, bool boolean = false, DateTime dateTime = default, string? formula = null)
    {
        Kind = kind;
        Text = text;
        Number = number;
        Boolean = boolean;
        DateTime = dateTime;
        Formula = formula;
    }

    public static CellValue Empty { get; } = new(CellValueKind.Empty);
    public static CellValue FromText(string text) => new(CellValueKind.Text, text: text);
    public static CellValue FromNumber(double number) => new(CellValueKind.Number, number: number);
    public static CellValue FromBoolean(bool value) => new(CellValueKind.Boolean, boolean: value);
    public static CellValue FromDateTime(DateTime value) => new(CellValueKind.DateTime, dateTime: value);
    public static CellValue FromFormula(string formula) => new(CellValueKind.Formula, formula: formula);

    public bool IsEmpty => Kind == CellValueKind.Empty;

    public override string ToString() => Kind switch
    {
        CellValueKind.Text => Text ?? string.Empty,
        CellValueKind.Number => Number.ToString(CultureInfo.InvariantCulture),
        CellValueKind.Boolean => Boolean ? "TRUE" : "FALSE",
        CellValueKind.DateTime => DateTime.ToString("O", CultureInfo.InvariantCulture),
        CellValueKind.Formula => $"={Formula}",
        _ => string.Empty
    };
}

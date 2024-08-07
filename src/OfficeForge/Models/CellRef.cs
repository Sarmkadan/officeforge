namespace OfficeForge.Models;

public readonly record struct CellRef(int Row, int Column)
{
    public string ToA1() => $"{ColumnName(Column)}{Row}";

    public static CellRef Parse(string a1)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(a1);
        var i = 0;
        var col = 0;
        while (i < a1.Length && char.IsAsciiLetter(a1[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(a1[i]) - 'A' + 1);
            i++;
        }
        if (i == 0 || i == a1.Length || !int.TryParse(a1.AsSpan(i), out var row) || row < 1 || col < 1)
            throw new FormatException($"Invalid cell reference '{a1}'.");
        return new CellRef(row, col);
    }

    public static bool TryParse(string a1, out CellRef cell)
    {
        try
        {
            cell = Parse(a1);
            return true;
        }
        catch (FormatException)
        {
            cell = default;
            return false;
        }
    }

    public static string ColumnName(int column)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);
        var name = string.Empty;
        while (column > 0)
        {
            column--;
            name = (char)('A' + column % 26) + name;
            column /= 26;
        }
        return name;
    }

    public override string ToString() => ToA1();
}

namespace OfficeForge.Models;

public readonly record struct CellRef(int Row, int Column)
{
	/// <summary>
	/// Maximum number of columns in Excel (XFD = 16384).
	/// </summary>
	public const int MaxColumns = 16384;

	/// <summary>
	/// Maximum number of rows in Excel (1048576).
	/// </summary>
	public const int MaxRows = 1048576;

	public string ToA1() => $"{ColumnName(Column)}{Row}";

	/// <summary>
	/// Parses an A1-style cell reference (e.g., "A1", "B2", "AA100", "XFD1048576").
	/// </summary>
	/// <param name="a1">The A1-style cell reference to parse.</param>
	/// <returns>A <see cref="CellRef"/> representing the parsed cell reference.</returns>
	/// <exception cref="ArgumentNullException">Thrown when <paramref name="a1"/> is null.</exception>
	/// <exception cref="ArgumentException">Thrown when <paramref name="a1"/> is empty or whitespace.</exception>
	/// <exception cref="FormatException">Thrown when <paramref name="a1"/> is not a valid A1-style cell reference.
	/// This includes:
	/// - Invalid format (e.g., "1A", "A", "A:")
	/// - Row or column out of Excel range (row < 1 or row > 1048576, column < 1 or column > 16384)
	/// </exception>
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

		if (i == 0)
		{
			throw new FormatException($"Invalid cell reference '{a1}': missing column letters.");
		}

		if (i == a1.Length)
		{
			throw new FormatException($"Invalid cell reference '{a1}': missing row number.");
		}

		if (!int.TryParse(a1.AsSpan(i), out var row) || row < 1)
		{
			throw new FormatException($"Invalid cell reference '{a1}': row must be a positive integer.");
		}

		if (col < 1)
		{
			throw new FormatException($"Invalid cell reference '{a1}': column must be positive.");
		}

		if (row > MaxRows)
		{
			throw new FormatException($"Invalid cell reference '{a1}': row {row} exceeds maximum Excel row of {MaxRows}.");
		}

		if (col > MaxColumns)
		{
			throw new FormatException($"Invalid cell reference '{a1}': column {ColumnName(col)} ({col}) exceeds maximum Excel column XFD ({MaxColumns}).");
		}

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
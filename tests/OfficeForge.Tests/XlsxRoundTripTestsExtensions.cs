using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeForge.Tests;

/// <summary>
/// Provides extension methods for <see cref="XlsxRoundTripTests"/>.
/// </summary>
public static class XlsxRoundTripTestsExtensions
{
    /// <summary>
    /// Safely disposes the test instance.
    /// </summary>
    /// <param name="tests">The test instance.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="tests"/> is null.</exception>
    public static void DisposeSafely(this XlsxRoundTripTests tests)
    {
        ArgumentNullException.ThrowIfNull(tests);
        tests.Dispose();
    }

    /// <summary>
    /// Executes an action with the test instance.
    /// </summary>
    /// <param name="tests">The test instance.</param>
    /// <param name="action">The action to execute.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="tests"/> or <paramref name="action"/> is null.</exception>
    public static void ExecuteTest(this XlsxRoundTripTests tests, Action<XlsxRoundTripTests> action)
    {
        ArgumentNullException.ThrowIfNull(tests);
        ArgumentNullException.ThrowIfNull(action);
        action(tests);
    }

    /// <summary>
    /// Gets the names of the available test methods.
    /// </summary>
    /// <param name="_">The test instance.</param>
    /// <returns>A list of test method names.</returns>
    public static IReadOnlyList<string> GetTestNames(this XlsxRoundTripTests _)
    {
        return new List<string>
        {
            nameof(XlsxRoundTripTests.WriteRead_PreservesTypedCellValues),
            nameof(XlsxRoundTripTests.WriteRead_PreservesMultipleSheets),
            nameof(XlsxRoundTripTests.WriteRead_EmptyWorkbookGetsDefaultSheet),
            nameof(XlsxRoundTripTests.Export_FromXlsxPath_ProducesMarkdownTable),
            nameof(XlsxRoundTripTests.MissingCell_ReadsAsEmpty)
        };
    }

    /// <summary>
    /// Converts a double to an invariant string.
    /// </summary>
    /// <param name="_">The test instance.</param>
    /// <param name="value">The double value.</param>
    /// <returns>The invariant string representation.</returns>
    public static string ToInvariantString(this XlsxRoundTripTests _, double value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}

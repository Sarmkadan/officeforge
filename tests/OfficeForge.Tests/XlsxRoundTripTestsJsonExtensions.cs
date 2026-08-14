using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeForge.Tests;

/// <summary>
/// Provides System.Text.Json helper methods for <see cref="XlsxRoundTripTests"/>.
/// </summary>
public static class XlsxRoundTripTestsJsonExtensions
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    /// <summary>
    /// Serializes the specified <see cref="XlsxRoundTripTests"/> instance to a JSON string.
    /// </summary>
    /// <param name="value">The <see cref="XlsxRoundTripTests"/> instance to serialize.</param>
    /// <param name="indented">Whether to format the JSON string with indentation.</param>
    /// <returns>A JSON string representation of the <see cref="XlsxRoundTripTests"/> instance.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="value"/> is null.</exception>
    public static string ToJson(this XlsxRoundTripTests value, bool indented = false)
    {
        ArgumentNullException.ThrowIfNull(value);
        var options = indented ? new JsonSerializerOptions(Options) { WriteIndented = true } : Options;
        return JsonSerializer.Serialize(value, options);
    }

    /// <summary>
    /// Deserializes a <see cref="XlsxRoundTripTests"/> instance from a JSON string.
    /// </summary>
    /// <param name="json">The JSON string to deserialize.</param>
    /// <returns>The deserialized <see cref="XlsxRoundTripTests"/> instance.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="json"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="json"/> is empty or whitespace.</exception>
    /// <exception cref="JsonException">Thrown when the JSON is invalid.</exception>
    public static XlsxRoundTripTests? FromJson(string json)
    {
        ArgumentException.ThrowIfNullOrEmpty(json);
        return JsonSerializer.Deserialize<XlsxRoundTripTests>(json, Options);
    }

    /// <summary>
    /// Attempts to deserialize a <see cref="XlsxRoundTripTests"/> instance from a JSON string.
    /// </summary>
    /// <param name="json">The JSON string to deserialize.</param>
    /// <param name="value">The deserialized <see cref="XlsxRoundTripTests"/> instance, or null if deserialization failed.</param>
    /// <returns>True if deserialization succeeded; otherwise, false.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="json"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="json"/> is empty or whitespace.</exception>
    public static bool TryFromJson(string json, out XlsxRoundTripTests? value)
    {
        ArgumentException.ThrowIfNullOrEmpty(json);
        try
        {
            value = JsonSerializer.Deserialize<XlsxRoundTripTests>(json, Options);
            return true;
        }
        catch (JsonException)
        {
            value = null;
            return false;
        }
    }
}

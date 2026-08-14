using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeForge.Tests
{
    /// <summary>
    /// Validation helpers for <see cref="XlsxRoundTripTests"/>.
    /// </summary>
    public static class XlsxRoundTripTestsValidation
    {
        /// <summary>
        /// Validates the given <paramref name="value"/> and returns a list of human-readable problems.
        /// </summary>
        /// <param name="value">The value to validate.</param>
        /// <returns>A list of human-readable problems.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="value"/> is null.</exception>
        public static IReadOnlyList<string> Validate(this XlsxRoundTripTests value)
        {
            ArgumentNullException.ThrowIfNull(value);

            var problems = new List<string>();

            // No validation is possible for the given members, as they are all methods.

            return problems;
        }

        /// <summary>
        /// Checks if the given <paramref name="value"/> is valid.
        /// </summary>
        /// <param name="value">The value to check.</param>
        /// <returns>True if the value is valid; otherwise, false.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="value"/> is null.</exception>
        public static bool IsValid(this XlsxRoundTripTests value)
        {
            ArgumentNullException.ThrowIfNull(value);

            return Validate(value).Count == 0;
        }

        /// <summary>
        /// Ensures the given <paramref name="value"/> is valid, throwing an exception if it is not.
        /// </summary>
        /// <param name="value">The value to ensure is valid.</param>
        /// <exception cref="ArgumentException">Thrown if <paramref name="value"/> is not valid.</exception>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="value"/> is null.</exception>
        public static void EnsureValid(this XlsxRoundTripTests value)
        {
            ArgumentNullException.ThrowIfNull(value);

            var problems = Validate(value);

            if (problems.Count > 0)
            {
                throw new ArgumentException(string.Join(Environment.NewLine, problems), nameof(value));
            }
        }
    }
}

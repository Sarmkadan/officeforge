using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeForge.Tests
{
    /// <summary>
    /// Provides validation helpers for <see cref="TemplateFillerTests"/>.
    /// </summary>
    public static class TemplateFillerTestsValidation
    {
        /// <summary>
        /// Validates the state of a <see cref="TemplateFillerTests"/> instance and returns a list of human‑readable problems.
        /// </summary>
        /// <param name="value">The instance to validate.</param>
        /// <returns>An <see cref="IReadOnlyList{T}"/> containing validation error messages. The list is empty when the instance is valid.</returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="value"/> is <c>null</c>.</exception>
        public static IReadOnlyList<string> Validate(this TemplateFillerTests value)
        {
            ArgumentNullException.ThrowIfNull(value);

            var problems = new List<string>();

            // No public fields or properties are defined on TemplateFillerTests that require validation.
            // If future members are added (e.g., strings, numbers, dates), validation logic should be
            // introduced here, checking for null/empty strings, out‑of‑range numbers, or default dates.

            return problems;
        }

        /// <summary>
        /// Determines whether the specified <see cref="TemplateFillerTests"/> instance is valid.
        /// </summary>
        /// <param name="value">The instance to check.</param>
        /// <returns><c>true</c> if the instance has no validation problems; otherwise, <c>false</c>.</returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="value"/> is <c>null</c>.</exception>
        public static bool IsValid(this TemplateFillerTests value) =>
            value.Validate().Count == 0;

        /// <summary>
        /// Ensures that the specified <see cref="TemplateFillerTests"/> instance is valid.
        /// </summary>
        /// <param name="value">The instance to validate.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="value"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">Thrown when validation problems are found. The exception message lists all problems.</exception>
        public static void EnsureValid(this TemplateFillerTests value)
        {
            ArgumentNullException.ThrowIfNull(value);

            var problems = value.Validate();
            if (problems.Count > 0)
            {
                var message = $"TemplateFillerTests is invalid: {string.Join("; ", problems)}";
                throw new ArgumentException(message, nameof(value));
            }
        }
    }
}

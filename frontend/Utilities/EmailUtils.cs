using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelFlow.Utilities
{
    public static class EmailUtils
    {
        private static readonly Regex EmailRegex = new(@"([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?)", RegexOptions.IgnoreCase);

        public static List<string> ExtractEmails(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return new();

            return EmailRegex.Matches(input)
                             .Cast<Match>()
                             .Select(m => m.Groups[1].Value.Trim())
                             .Where(email => !string.IsNullOrWhiteSpace(email))
                             .Distinct(StringComparer.OrdinalIgnoreCase)
                             .ToList();
        }
    }
}

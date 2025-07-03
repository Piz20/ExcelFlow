namespace ExcelFlow.Utilities;
public static class StringUtils
{
    public static int ComputeLevenshteinDistance(string s, string t)
    {
        int n = s.Length;
        int m = t.Length;
        var d = new int[n + 1, m + 1];

        for (int i = 0; i <= n; i++) d[i, 0] = i;
        for (int j = 0; j <= m; j++) d[0, j] = j;

        for (int i = 1; i <= n; i++)
        {
            for (int j = 1; j <= m; j++)
            {
                int cost = (s[i - 1] == t[j - 1]) ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }
        return d[n, m];
    }

    public static string ShowStringDifferences(string original, string compared)
    {
        int maxLength = Math.Max(original.Length, compared.Length);
        var diffBuilder = new System.Text.StringBuilder();

        diffBuilder.AppendLine($"Original : \"{original}\"");
        diffBuilder.AppendLine($"Compared : \"{compared}\"");
        diffBuilder.AppendLine("Diff    : ");

        for (int i = 0; i < maxLength; i++)
        {
            char c1 = i < original.Length ? original[i] : '-';
            char c2 = i < compared.Length ? compared[i] : '-';

            diffBuilder.Append(c1 == c2 ? ' ' : '^');
        }

        return diffBuilder.ToString();
    }
}

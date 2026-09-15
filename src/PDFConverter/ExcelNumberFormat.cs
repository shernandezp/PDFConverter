using System.Globalization;
using System.Text;

namespace PDFConverter;

internal static class ExcelNumberFormat
{
    public static string Apply(string rawValue, string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode) || string.IsNullOrEmpty(rawValue)) return rawValue;
        if (!Units.TryParseDouble(rawValue, out var value)) return rawValue;

        var section = SelectSection(formatCode, value);
        if (section.Length == 0) return string.Empty;

        var (pattern, isDate) = Translate(section);
        if (pattern.Length == 0) return string.Empty;

        try
        {
            return isDate
                ? DateTime.FromOADate(value).ToString(pattern, CultureInfo.CurrentCulture)
                : Math.Abs(value).ToString(pattern, CultureInfo.CurrentCulture);
        }
        catch
        {
            return rawValue;
        }
    }

    // Excel formats read "positive;negative;zero;text" and bake the sign into the chosen section,
    // so once a section is picked the magnitude is what gets formatted.
    static string SelectSection(string formatCode, double value)
    {
        var sections = SplitSections(formatCode);
        if (sections.Count == 1) return value < 0 ? "-" + sections[0] : sections[0];
        if (value < 0) return sections[1];
        if (value == 0 && sections.Count > 2) return sections[2];
        return sections[0];
    }

    static List<string> SplitSections(string formatCode)
    {
        var sections = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;
        var inBracket = false;

        foreach (var c in formatCode)
        {
            if (c == '"') inQuotes = !inQuotes;
            else if (!inQuotes && c == '[') inBracket = true;
            else if (!inQuotes && c == ']') inBracket = false;

            if (c == ';' && !inQuotes && !inBracket)
            {
                sections.Add(current.ToString());
                current.Clear();
                continue;
            }
            current.Append(c);
        }
        sections.Add(current.ToString());
        return sections;
    }

    static (string Pattern, bool IsDate) Translate(string section)
    {
        var tokens = Tokenize(section);
        var isDate = tokens.Any(t => t.Kind == TokenKind.DateTime)
            && tokens.All(t => t.Kind != TokenKind.Numeric);
        var hasAmPm = tokens.Any(t => t.Kind == TokenKind.AmPm);

        var pattern = new StringBuilder();
        for (var i = 0; i < tokens.Count; i++)
        {
            pattern.Append(tokens[i].Kind switch
            {
                TokenKind.AmPm => "tt",
                TokenKind.DateTime => TranslateDateToken(tokens[i].Text, tokens, i, hasAmPm),
                TokenKind.Numeric or TokenKind.Separator => tokens[i].Text,
                _ => EscapeLiteral(tokens[i].Text),
            });
        }
        return (pattern.ToString(), isDate);
    }

    static string TranslateDateToken(string token, List<Token> tokens, int index, bool hasAmPm)
    {
        var letter = token[0];
        if (letter == 'h') return new string(hasAmPm ? 'h' : 'H', token.Length);
        if (letter == 's') return new string('s', token.Length);
        if (letter == 'y') return token.Length <= 2 ? "yy" : "yyyy";
        if (letter == 'd') return token;
        return IsMinuteContext(tokens, index) ? new string('m', token.Length) : token.ToUpperInvariant();
    }

    // "m" means minutes next to an hour or second field and months everywhere else.
    static bool IsMinuteContext(List<Token> tokens, int index)
    {
        return Neighbour(tokens, index, -1) || Neighbour(tokens, index, 1);

        static bool Neighbour(List<Token> tokens, int index, int step)
        {
            for (var i = index + step; i >= 0 && i < tokens.Count; i += step)
            {
                if (tokens[i].Kind is TokenKind.Separator or TokenKind.Literal) continue;
                return tokens[i].Kind == TokenKind.DateTime && tokens[i].Text[0] is 'h' or 's';
            }
            return false;
        }
    }

    static string EscapeLiteral(string text)
    {
        var escaped = new StringBuilder(text.Length * 2);
        foreach (var c in text) escaped.Append('\\').Append(c);
        return escaped.ToString();
    }

    enum TokenKind { Literal, Separator, DateTime, AmPm, Numeric }

    readonly record struct Token(TokenKind Kind, string Text);

    static List<Token> Tokenize(string section)
    {
        var tokens = new List<Token>();
        var i = 0;

        while (i < section.Length)
        {
            var c = section[i];

            if (c == '[')
            {
                var close = section.IndexOf(']', i);
                i = close < 0 ? section.Length : close + 1;
                continue;
            }
            if (c == '"')
            {
                var close = section.IndexOf('"', i + 1);
                var text = close < 0 ? section[(i + 1)..] : section[(i + 1)..close];
                if (text.Length > 0) tokens.Add(new Token(TokenKind.Literal, text));
                i = close < 0 ? section.Length : close + 1;
                continue;
            }
            if (c == '\\' && i + 1 < section.Length)
            {
                tokens.Add(new Token(TokenKind.Literal, section[i + 1].ToString()));
                i += 2;
                continue;
            }
            // "_x" reserves the width of x and "*x" pads with x; neither has a .NET equivalent.
            if (c is '_' or '*')
            {
                i += 2;
                continue;
            }
            if (MatchesAt(section, i, "AM/PM") || MatchesAt(section, i, "am/pm"))
            {
                tokens.Add(new Token(TokenKind.AmPm, "tt"));
                i += 5;
                continue;
            }
            if (char.ToLowerInvariant(c) is 'y' or 'm' or 'd' or 'h' or 's')
            {
                var start = i;
                while (i < section.Length
                    && char.ToLowerInvariant(section[i]) == char.ToLowerInvariant(c)) i++;
                tokens.Add(new Token(TokenKind.DateTime, section[start..i].ToLowerInvariant()));
                continue;
            }
            if (c is 'E' or 'e' && i + 1 < section.Length && section[i + 1] is '+' or '-')
            {
                tokens.Add(new Token(TokenKind.Numeric, section.Substring(i, 2)));
                i += 2;
                continue;
            }
            if (c is '0' or '#' or '%')
            {
                tokens.Add(new Token(TokenKind.Numeric, c.ToString()));
                i++;
                continue;
            }
            if (c == '?')
            {
                tokens.Add(new Token(TokenKind.Numeric, "#"));
                i++;
                continue;
            }
            if (c is '.' or ',' or '/' or ':')
            {
                tokens.Add(new Token(TokenKind.Separator, c.ToString()));
                i++;
                continue;
            }

            tokens.Add(new Token(TokenKind.Literal, c.ToString()));
            i++;
        }
        return tokens;
    }

    static bool MatchesAt(string text, int index, string value) =>
        index + value.Length <= text.Length && string.CompareOrdinal(text, index, value, 0, value.Length) == 0;
}

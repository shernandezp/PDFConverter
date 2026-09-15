using System.Globalization;
using System.Text;

namespace PDFConverter;

/// <summary>Tracks list counters across a document and renders the label Word would show.</summary>
internal sealed class WordNumbering(WordStyleCache styles)
{
    const int MaxLevels = 9;
    const string DefaultBullet = "•";

    readonly Dictionary<string, int[]> _countersByList = new(StringComparer.Ordinal);

    public string? NextLabel(string numId, int level)
    {
        if (level < 0 || level >= MaxLevels) return null;

        if (!_countersByList.TryGetValue(numId, out var counters))
            _countersByList[numId] = counters = new int[MaxLevels];

        counters[level]++;
        for (var i = level + 1; i < MaxLevels; i++) counters[i] = 0;

        var definition = styles.GetNumberingLevel(numId, level);
        if (string.Equals(definition.Format, "none", StringComparison.OrdinalIgnoreCase)) return null;
        if (string.Equals(definition.Format, "bullet", StringComparison.OrdinalIgnoreCase))
            return BulletGlyph(definition.LevelText) + " ";

        return Substitute(definition.LevelText, numId, counters) + " ";
    }

    string Substitute(string levelText, string numId, int[] counters)
    {
        var label = new StringBuilder();
        for (var i = 0; i < levelText.Length; i++)
        {
            if (levelText[i] == '%' && i + 1 < levelText.Length && char.IsDigit(levelText[i + 1]))
            {
                var referenced = levelText[i + 1] - '1';
                if (referenced >= 0 && referenced < MaxLevels)
                {
                    var definition = styles.GetNumberingLevel(numId, referenced);
                    var value = counters[referenced] + (definition.StartAt ?? 1) - 1;
                    label.Append(FormatNumber(value, definition.Format));
                    i++;
                    continue;
                }
            }
            label.Append(levelText[i]);
        }
        return label.ToString();
    }

    static string FormatNumber(int value, string format) => format switch
    {
        "decimalZero" => value.ToString("00", CultureInfo.InvariantCulture),
        "lowerLetter" => Alphabetic(value, 'a'),
        "upperLetter" => Alphabetic(value, 'A'),
        "lowerRoman" => ConverterExtensions.ToRoman(value).ToLowerInvariant(),
        "upperRoman" => ConverterExtensions.ToRoman(value),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    /// <summary>Word repeats the letter past z, so 27 becomes "aa".</summary>
    static string Alphabetic(int value, char start)
    {
        if (value < 1) return string.Empty;
        var index = value - 1;
        return new string((char)(start + index % 26), index / 26 + 1);
    }

    /// <summary>Bullet glyphs live in the private use area of Symbol and Wingdings.</summary>
    static string BulletGlyph(string levelText) => levelText.Length == 0 ? DefaultBullet : (int)levelText[0] switch
    {
        0xF0B7 or 0x00B7 => DefaultBullet,
        0xF0A7 or 0xF06E => "▪",
        0xF06F or 0xF0A1 or 'o' => "○",
        0xF0D8 or 0xF0E8 => "▸",
        '-' => "–",
        var glyph when glyph >= ' ' && glyph < 0xF000 => levelText,
        _ => DefaultBullet,
    };
}

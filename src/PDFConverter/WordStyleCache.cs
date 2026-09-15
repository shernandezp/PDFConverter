using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

/// <summary>
/// Per-document cache for style, theme and numbering lookups. Resolving these from the OpenXML tree
/// costs a linear scan each time, which is significant when it happens once per run or per list item.
/// </summary>
internal sealed class WordStyleCache
{
    static readonly ConditionalWeakTable<MainDocumentPart, WordStyleCache> s_caches = new();

    readonly MainDocumentPart _mainPart;
    readonly Dictionary<string, W.Style> _stylesById;
    readonly Dictionary<string, W.RunProperties?> _runPropertiesByStyle = new(StringComparer.Ordinal);
    readonly Dictionary<string, W.ParagraphProperties?> _paragraphPropertiesByStyle = new(StringComparer.Ordinal);
    readonly Dictionary<(string numId, int level), NumberingLevel> _numberingLevels = new();

    W.RunProperties? _docDefaultsRun;
    W.ParagraphProperties? _docDefaultsParagraph;
    string? _themeMinorFont;
    string? _themeMajorFont;
    bool _docDefaultsRead;
    bool _themeFontsRead;
    W.Numbering? _numbering;
    bool _numberingRead;

    WordStyleCache(MainDocumentPart mainPart)
    {
        _mainPart = mainPart;
        _stylesById = new Dictionary<string, W.Style>(StringComparer.Ordinal);
        var styles = mainPart.StyleDefinitionsPart?.Styles;
        if (styles == null) return;
        foreach (var style in styles.Elements<W.Style>())
        {
            var id = style.StyleId?.Value;
            if (!string.IsNullOrEmpty(id)) _stylesById.TryAdd(id, style);
        }
    }

    public static WordStyleCache For(MainDocumentPart mainPart) =>
        s_caches.GetValue(mainPart, part => new WordStyleCache(part));

    public W.Style? GetStyle(string? styleId) =>
        !string.IsNullOrEmpty(styleId) && _stylesById.TryGetValue(styleId, out var style) ? style : null;

    /// <summary>
    /// Walks a style and its basedOn chain, returning the first style whose <paramref name="select"/> yields a value.
    /// </summary>
    T? WalkStyleChain<T>(string? styleId, Func<W.Style, T?> select) where T : class
    {
        var style = GetStyle(styleId);
        var guard = 0;
        while (style != null && guard++ < 32)
        {
            var value = select(style);
            if (value != null) return value;
            style = GetStyle(style.BasedOn?.Val?.Value);
        }
        return null;
    }

    public W.RunProperties? GetStyleRunPropertiesOrNull(string? styleId) =>
        string.IsNullOrEmpty(styleId) ? null : GetStyleRunProperties(styleId);

    public T? WalkTableStyles<T>(string? styleId, Func<W.Style, T?> select) where T : class =>
        WalkStyleChain(styleId, style => style.Type?.Value == W.StyleValues.Table || style.Type == null
            ? select(style)
            : null);

    public W.RunProperties? GetStyleRunProperties(string styleId)
    {
        if (_runPropertiesByStyle.TryGetValue(styleId, out var cached)) return cached;
        var resolved = WalkStyleChain(styleId, s => CopyProperties<W.RunProperties>(s.StyleRunProperties));
        _runPropertiesByStyle[styleId] = resolved;
        return resolved;
    }

    public W.ParagraphProperties? GetStyleParagraphProperties(string styleId)
    {
        if (_paragraphPropertiesByStyle.TryGetValue(styleId, out var cached)) return cached;
        var resolved = WalkStyleChain(styleId, s => CopyProperties<W.ParagraphProperties>(s.StyleParagraphProperties));
        _paragraphPropertiesByStyle[styleId] = resolved;
        return resolved;
    }

    /// <summary>
    /// StyleRunProperties / StyleParagraphProperties carry the same children as their run/paragraph
    /// counterparts but are distinct OpenXML types, so a cast-based clone always fails; copy the children instead.
    /// </summary>
    static T? CopyProperties<T>(DocumentFormat.OpenXml.OpenXmlElement? source) where T : DocumentFormat.OpenXml.OpenXmlElement, new()
    {
        if (source == null) return null;
        var target = new T();
        foreach (var child in source.ChildElements)
            target.AppendChild(child.CloneNode(true));
        return target;
    }

    public W.RunProperties? DocDefaultsRunProperties
    {
        get { ReadDocDefaults(); return _docDefaultsRun; }
    }

    public W.ParagraphProperties? DocDefaultsParagraphProperties
    {
        get { ReadDocDefaults(); return _docDefaultsParagraph; }
    }

    void ReadDocDefaults()
    {
        if (_docDefaultsRead) return;
        _docDefaultsRead = true;
        var docDefaults = _mainPart.StyleDefinitionsPart?.Styles?.GetFirstChild<W.DocDefaults>();

        // Parsing a .docx yields the *BaseStyle types here, which are distinct from the run and
        // paragraph property types despite holding the same children; building the tree in memory
        // can yield either.
        var runDefault = docDefaults?.GetFirstChild<W.RunPropertiesDefault>();
        _docDefaultsRun = runDefault?.GetFirstChild<W.RunProperties>()
            ?? CopyProperties<W.RunProperties>(runDefault?.GetFirstChild<W.RunPropertiesBaseStyle>());

        var paragraphDefault = docDefaults?.GetFirstChild<W.ParagraphPropertiesDefault>();
        _docDefaultsParagraph = paragraphDefault?.GetFirstChild<W.ParagraphProperties>()
            ?? CopyProperties<W.ParagraphProperties>(
                paragraphDefault?.GetFirstChild<W.ParagraphPropertiesBaseStyle>());
    }

    public string? ThemeFont { get { ReadThemeFonts(); return _themeMinorFont; } }

    public string? ThemeMajorFont { get { ReadThemeFonts(); return _themeMajorFont ?? _themeMinorFont; } }

    void ReadThemeFonts()
    {
        if (_themeFontsRead) return;
        _themeFontsRead = true;
        var fontScheme = _mainPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
        _themeMinorFont = Normalize(fontScheme?.MinorFont?.LatinFont?.Typeface?.Value);
        _themeMajorFont = Normalize(fontScheme?.MajorFont?.LatinFont?.Typeface?.Value);

        static string? Normalize(string? typeface) => string.IsNullOrEmpty(typeface) ? null : typeface;
    }

    public NumberingLevel GetNumberingLevel(string? numId, int level)
    {
        if (string.IsNullOrEmpty(numId)) return NumberingLevel.Default;
        var key = (numId, level);
        if (_numberingLevels.TryGetValue(key, out var cached)) return cached;
        var resolved = ResolveNumberingLevel(numId, level);
        _numberingLevels[key] = resolved;
        return resolved;
    }

    NumberingLevel ResolveNumberingLevel(string numId, int level)
    {
        if (!_numberingRead)
        {
            _numberingRead = true;
            _numbering = _mainPart.NumberingDefinitionsPart?.Numbering;
        }
        if (_numbering == null) return NumberingLevel.Default;

        var num = _numbering.Elements<W.NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value.ToString() == numId);
        if (num == null) return NumberingLevel.Default;

        // A w:lvlOverride on the instance replaces the abstract definition for that level.
        var over = num.Elements<W.LevelOverride>().FirstOrDefault(o => o.LevelIndex?.Value == level);
        var startOverride = over?.StartOverrideNumberingValue?.Val?.Value;

        var lvl = over?.Level;
        if (lvl == null)
        {
            var abstractId = num.AbstractNumId?.Val?.Value;
            var abstractNum = abstractId == null
                ? null
                : _numbering.Elements<W.AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);
            lvl = abstractNum?.Elements<W.Level>().FirstOrDefault(l => l.LevelIndex?.Value == level);
        }
        if (lvl == null) return NumberingLevel.Default with { StartAt = startOverride };

        return new NumberingLevel(
            lvl.NumberingFormat?.Val?.InnerText ?? "decimal",
            lvl.LevelText?.Val?.Value ?? "%1.",
            startOverride ?? lvl.StartNumberingValue?.Val?.Value);
    }
}

internal readonly record struct NumberingLevel(string Format, string LevelText, int? StartAt)
{
    public static NumberingLevel Default { get; } = new("decimal", "%1.", null);
}

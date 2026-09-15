using System.Collections.Concurrent;
using PdfSharp.Fonts;

namespace PDFConverter;

internal sealed class DirectoryFontResolver : IFontResolver
{
    const string EmbeddedEmojiResource = "PDFConverter.Fonts.NotoEmoji-Regular.ttf";

    static readonly string[] FallbackFamilies =
        ["Arial", "Helvetica", "Calibri", "Segoe UI", "Liberation Sans", "DejaVu Sans", "Times New Roman"];

    readonly Dictionary<string, FontFace> _facesByKey = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, List<FontFace>> _facesByFamily = new(StringComparer.OrdinalIgnoreCase);
    readonly ConcurrentDictionary<string, byte[]> _loadedFonts = new(StringComparer.OrdinalIgnoreCase);
    readonly Action<string>? _logger;

    sealed record FontFace(string Family, string Style, string Key, string? Path, byte[]? Bytes);

    public DirectoryFontResolver(IEnumerable<string> fontFiles,
        IDictionary<string, string>? explicitMappings, Action<string>? logger)
    {
        _logger = logger;

        foreach (var file in fontFiles) Register(file, familyOverride: null);

        if (explicitMappings != null)
            foreach (var (family, path) in explicitMappings) Register(path, family);

        RegisterEmbeddedEmojiFont();
        logger?.Invoke($"Font resolver indexed {_facesByKey.Count} faces across {_facesByFamily.Count} families.");
    }

    void Register(string path, string? familyOverride)
    {
        try
        {
            if (!File.Exists(path))
            {
                _logger?.Invoke($"Font file not found: {path}");
                return;
            }

            using var stream = File.OpenRead(path);
            var names = FontUtils.ReadFontNames(stream);
            var family = familyOverride
                ?? (names.TryGetValue(1, out var parsed) ? parsed : Path.GetFileNameWithoutExtension(path));
            var style = names.TryGetValue(2, out var subFamily) ? subFamily : "Regular";
            Add(new FontFace(family, style, Key(family, style), path, null));
        }
        catch (Exception ex)
        {
            _logger?.Invoke($"Skipped unreadable font '{path}': {ex.Message}");
        }
    }

    void RegisterEmbeddedEmojiFont()
    {
        try
        {
            using var stream = typeof(DirectoryFontResolver).Assembly
                .GetManifestResourceStream(EmbeddedEmojiResource);
            if (stream == null) return;

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            var bytes = buffer.ToArray();

            var names = FontUtils.ReadFontNames(bytes);
            var family = names.TryGetValue(1, out var parsed) ? parsed : "Noto Emoji";
            var style = names.TryGetValue(2, out var subFamily) ? subFamily : "Regular";
            Add(new FontFace(family, style, Key(family, style), null, bytes));
        }
        catch (Exception ex)
        {
            _logger?.Invoke($"Failed loading the embedded emoji font: {ex.Message}");
        }
    }

    void Add(FontFace face)
    {
        if (!_facesByKey.TryAdd(face.Key, face)) return;
        if (!_facesByFamily.TryGetValue(face.Family, out var faces))
            _facesByFamily[face.Family] = faces = [];
        faces.Add(face);
    }

    static string Key(string family, string style) =>
        $"{family.Trim()}|{(string.IsNullOrWhiteSpace(style) ? "Regular" : style.Trim())}";

    static string StyleName(bool bold, bool italic) =>
        bold && italic ? "Bold Italic" : bold ? "Bold" : italic ? "Italic" : "Regular";

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var resolved = Resolve(familyName, isBold, isItalic);
        if (resolved != null) return resolved;

        foreach (var fallback in FallbackFamilies)
        {
            resolved = Resolve(fallback, isBold, isItalic);
            if (resolved != null) return resolved;
        }

        var any = _facesByKey.Values.FirstOrDefault();
        return any == null ? null : new FontResolverInfo(any.Key, isBold, isItalic);
    }

    // Without the simulate flags a family that ships only a regular face renders upright and light.
    FontResolverInfo? Resolve(string? familyName, bool isBold, bool isItalic)
    {
        if (string.IsNullOrEmpty(familyName)) return null;

        if (_facesByKey.TryGetValue(Key(familyName, StyleName(isBold, isItalic)), out var exact))
            return new FontResolverInfo(exact.Key);

        if (!_facesByFamily.TryGetValue(familyName, out var faces) || faces.Count == 0)
        {
            var match = _facesByFamily.Keys.FirstOrDefault(
                f => f.StartsWith(familyName, StringComparison.OrdinalIgnoreCase));
            if (match == null || !_facesByFamily.TryGetValue(match, out faces)) return null;
        }

        var regular = faces.FirstOrDefault(f => f.Style.Equals("Regular", StringComparison.OrdinalIgnoreCase))
            ?? faces[0];
        return new FontResolverInfo(regular.Key, isBold, isItalic);
    }

    // Faces are indexed by path and read on demand: a system font folder is hundreds of megabytes.
    public byte[]? GetFont(string faceName)
    {
        if (string.IsNullOrEmpty(faceName) || !_facesByKey.TryGetValue(faceName, out var face)) return null;
        if (face.Bytes != null) return face.Bytes;
        if (face.Path == null) return null;

        return _loadedFonts.GetOrAdd(faceName, _ =>
        {
            try
            {
                return File.ReadAllBytes(face.Path);
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"Failed reading font '{face.Path}': {ex.Message}");
                return [];
            }
        });
    }
}

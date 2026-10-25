using System.Collections.Concurrent;
using SkiaSharp;

namespace PDFConverter.Layout;

internal sealed class FontFace
{
    public required string Family { get; init; }
    public required SKTypeface Typeface { get; init; }
    public bool FakeBold { get; init; }
    public bool FakeItalic { get; init; }
    public string? Path { get; init; }
    public byte[]? Bytes { get; init; }

    public SKFont CreateFont(double size) => new(Typeface, (float)size)
    {
        Embolden = FakeBold,
        SkewX = FakeItalic ? -0.25f : 0f,
        Hinting = SKFontHinting.None,
        Subpixel = true,
    };

    public (byte[] Data, int CollectionIndex)? ReadBytes()
    {
        if (Bytes != null) return (Bytes, 0);
        try
        {
            if (Path != null) return (File.ReadAllBytes(Path), 0);
            using var stream = Typeface.OpenStream(out var collectionIndex);
            if (stream == null) return null;
            var data = new byte[stream.Length];
            stream.Read(data, data.Length);
            return (data, collectionIndex);
        }
        catch
        {
            return null;
        }
    }
}

internal sealed class FontRegistry
{
    const string EmbeddedEmojiResource = "PDFConverter.Fonts.NotoEmoji-Regular.ttf";

    static readonly string[] FallbackFamilies =
        ["Arial", "Helvetica", "Calibri", "Segoe UI", "Liberation Sans", "DejaVu Sans", "Times New Roman"];

    static readonly Lock Gate = new();
    static FontRegistry? s_current;

    public static bool IsInstalled => s_current != null;

    public static FontRegistry Current
    {
        get
        {
            if (s_current != null) return s_current;
            lock (Gate)
            {
                if (s_current == null) OpenXmlHelpers.RegisterFontsFromDirectory(null);
                return s_current ??= new FontRegistry([], null, null);
            }
        }
    }

    public static void Install(FontRegistry registry)
    {
        lock (Gate) s_current = registry;
    }

    sealed record IndexedFace(string Family, string Style, string? Path, byte[]? Bytes);

    readonly Dictionary<string, IndexedFace> _facesByKey = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, List<IndexedFace>> _facesByFamily = new(StringComparer.OrdinalIgnoreCase);
    readonly ConcurrentDictionary<string, FontFace> _resolved = new(StringComparer.OrdinalIgnoreCase);
    readonly ConcurrentDictionary<string, SKTypeface?> _typefaces = new(StringComparer.OrdinalIgnoreCase);
    readonly ConcurrentDictionary<string, FontFace> _systemFaces = new(StringComparer.OrdinalIgnoreCase);
    readonly ConcurrentDictionary<string, FontFace?> _fallbacks = new(StringComparer.Ordinal);
    readonly Action<string>? _logger;
    FontFace? _emoji;
    SKFont? _emojiFont;

    public FontRegistry(IEnumerable<string> fontFiles, IDictionary<string, string>? explicitMappings, Action<string>? logger)
    {
        _logger = logger;

        foreach (var file in fontFiles) Register(file, familyOverride: null);

        if (explicitMappings != null)
            foreach (var (family, path) in explicitMappings) Register(path, family);

        RegisterEmbeddedEmojiFont();
        logger?.Invoke($"Font registry indexed {_facesByKey.Count} faces across {_facesByFamily.Count} families.");
    }

    public int FaceCount => _facesByKey.Count;

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
                ?? (names.TryGetValue(1, out var parsed) ? parsed : System.IO.Path.GetFileNameWithoutExtension(path));
            var style = names.TryGetValue(2, out var subFamily) ? subFamily : "Regular";
            Add(new IndexedFace(family, style, path, null));
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
            using var stream = typeof(FontRegistry).Assembly.GetManifestResourceStream(EmbeddedEmojiResource);
            if (stream == null) return;

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            var bytes = buffer.ToArray();

            var names = FontUtils.ReadFontNames(bytes);
            var family = names.TryGetValue(1, out var parsed) ? parsed : "Noto Emoji";
            var style = names.TryGetValue(2, out var subFamily) ? subFamily : "Regular";
            var face = new IndexedFace(family, style, null, bytes);
            Add(face);

            var typeface = SKTypeface.FromData(SKData.CreateCopy(bytes));
            if (typeface != null)
                _emoji = new FontFace { Family = family, Typeface = typeface, Bytes = bytes };
        }
        catch (Exception ex)
        {
            _logger?.Invoke($"Failed loading the embedded emoji font: {ex.Message}");
        }
    }

    void Add(IndexedFace face)
    {
        if (!_facesByKey.TryAdd(Key(face.Family, face.Style), face)) return;
        if (!_facesByFamily.TryGetValue(face.Family, out var faces))
            _facesByFamily[face.Family] = faces = [];
        faces.Add(face);
    }

    static string Key(string family, string style) =>
        $"{family.Trim()}|{(string.IsNullOrWhiteSpace(style) ? "Regular" : style.Trim())}";

    static string StyleName(bool bold, bool italic) =>
        bold && italic ? "Bold Italic" : bold ? "Bold" : italic ? "Italic" : "Regular";

    public FontFace Resolve(string? family, bool bold, bool italic)
    {
        var key = $"{family ?? string.Empty}|{bold}|{italic}";
        return _resolved.GetOrAdd(key, _ => ResolveUncached(family, bold, italic));
    }

    FontFace ResolveUncached(string? family, bool bold, bool italic)
    {
        var resolved = ResolveFamily(family, bold, italic);
        if (resolved != null) return resolved;

        foreach (var fallback in FallbackFamilies)
        {
            resolved = ResolveFamily(fallback, bold, italic);
            if (resolved != null) return resolved;
        }

        var any = _facesByKey.Values.Select(face => Load(face)).FirstOrDefault(typeface => typeface != null);
        if (any != null)
            return new FontFace { Family = any.FamilyName, Typeface = any, FakeBold = bold, FakeItalic = italic };

        return new FontFace { Family = SKTypeface.Default.FamilyName, Typeface = SKTypeface.Default, FakeBold = bold, FakeItalic = italic };
    }

    // An indexed file with the exact style wins; otherwise the platform font manager may still hold the
    // exact style (collections are not indexed) before a nearby indexed face is synthesised.
    FontFace? ResolveFamily(string? family, bool bold, bool italic)
    {
        if (string.IsNullOrEmpty(family)) return null;

        if (_facesByKey.TryGetValue(Key(family, StyleName(bold, italic)), out var exact))
            return ToFace(exact, false, false);

        var system = ResolveSystem(family, bold, italic);
        if (system is { FakeBold: false, FakeItalic: false }) return system;
        return ResolveIndexed(family, bold, italic) ?? system;
    }

    FontFace? ResolveIndexed(string? family, bool bold, bool italic)
    {
        if (string.IsNullOrEmpty(family)) return null;

        if (!_facesByFamily.TryGetValue(family, out var faces) || faces.Count == 0)
        {
            var match = _facesByFamily.Keys.FirstOrDefault(
                f => f.StartsWith(family, StringComparison.OrdinalIgnoreCase));
            if (match == null || !_facesByFamily.TryGetValue(match, out faces)) return null;
        }

        var closest = Closest(faces, bold, italic);
        var hasBold = closest.Style.Contains("Bold", StringComparison.OrdinalIgnoreCase);
        var hasItalic = closest.Style.Contains("Italic", StringComparison.OrdinalIgnoreCase)
            || closest.Style.Contains("Oblique", StringComparison.OrdinalIgnoreCase);
        return ToFace(closest, bold && !hasBold, italic && !hasItalic);
    }

    static IndexedFace Closest(List<IndexedFace> faces, bool bold, bool italic)
    {
        IndexedFace? best = null;
        var bestScore = int.MinValue;
        foreach (var face in faces)
        {
            var faceBold = face.Style.Contains("Bold", StringComparison.OrdinalIgnoreCase);
            var faceItalic = face.Style.Contains("Italic", StringComparison.OrdinalIgnoreCase)
                || face.Style.Contains("Oblique", StringComparison.OrdinalIgnoreCase);
            var score = (faceBold == bold ? 2 : faceBold ? -2 : 0) + (faceItalic == italic ? 1 : faceItalic ? -1 : 0);
            if (score <= bestScore) continue;
            best = face;
            bestScore = score;
        }
        return best ?? faces[0];
    }

    FontFace? ToFace(IndexedFace face, bool fakeBold, bool fakeItalic)
    {
        var typeface = Load(face);
        if (typeface == null) return null;
        return new FontFace
        {
            Family = face.Family,
            Typeface = typeface,
            FakeBold = fakeBold,
            FakeItalic = fakeItalic,
            Path = face.Path,
            Bytes = face.Bytes,
        };
    }

    SKTypeface? Load(IndexedFace face)
    {
        var key = face.Path ?? Key(face.Family, face.Style);
        return _typefaces.GetOrAdd(key, _ =>
        {
            try
            {
                return face.Bytes != null
                    ? SKTypeface.FromData(SKData.CreateCopy(face.Bytes))
                    : SKTypeface.FromFile(face.Path);
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"Failed loading font '{face.Path}': {ex.Message}");
                return null;
            }
        });
    }

    static FontFace? ResolveSystem(string? family, bool bold, bool italic)
    {
        if (string.IsNullOrEmpty(family)) return null;
        try
        {
            var style = new SKFontStyle(bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal, italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);
            var typeface = SKFontManager.Default.MatchFamily(family, style);
            if (typeface == null || !typeface.FamilyName.StartsWith(family, StringComparison.OrdinalIgnoreCase))
                return null;

            return new FontFace
            {
                Family = typeface.FamilyName,
                Typeface = typeface,
                FakeBold = bold && !typeface.IsBold,
                FakeItalic = italic && !typeface.IsItalic,
            };
        }
        catch
        {
            return null;
        }
    }

    public FontFace? Fallback(FontFace primary, int codepoint)
    {
        if (_emoji != null)
        {
            _emojiFont ??= new SKFont(_emoji.Typeface);
            if (_emojiFont.GetGlyph(codepoint) != 0) return _emoji;
        }

        return _fallbacks.GetOrAdd($"{primary.Family}|{codepoint}", _ =>
        {
            try
            {
                var typeface = SKFontManager.Default.MatchCharacter(primary.Family, primary.Typeface.FontStyle, null, codepoint);
                if (typeface == null) return null;
                return _systemFaces.GetOrAdd(typeface.FamilyName, _ => new FontFace
                {
                    Family = typeface.FamilyName,
                    Typeface = typeface,
                });
            }
            catch
            {
                return null;
            }
        });
    }
}

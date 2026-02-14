using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using PdfSharp.Fonts;

namespace PDFConverter;

public static class OpenXmlHelpers
{
    /// <summary>
    /// Logger for font loading/registration messages. Set by caller if desired.
    /// </summary>
    public static Action<string>? FontLoadLogger { get; set; }

    /// <summary>
    /// Logger for image loading/resolution messages.
    /// </summary>
    public static Action<string>? ImageLoadLogger { get; set; }

    // Explicit family->file mappings persisted until next registration call
    static Dictionary<string, string>? s_explicitFamilyMappings;

    /// <summary>
    /// Register explicit mappings of font family name to a font file path. These mappings will be preferred when registering fonts.
    /// </summary>
    public static void RegisterFontMappings(IDictionary<string, string> mappings)
    {
        if (mappings == null) return;
        s_explicitFamilyMappings = new Dictionary<string, string>(mappings, StringComparer.OrdinalIgnoreCase);
        FontLoadLogger?.Invoke($"Registered {s_explicitFamilyMappings.Count} explicit font mappings.");
    }

    /// <summary>
    /// Ensure a font resolver is initialized. This will register system fonts if no resolver is present.
    /// Safe to call multiple times.
    /// </summary>
    public static void EnsureFontResolverInitialized()
    {
        if (GlobalFontSettings.FontResolver != null) return;
        // Register system fonts by default; explicit mappings (if any) will be used as fallback.
        RegisterFontsFromDirectory(null);
    }

    /// <summary>
    /// Register fonts found in the specified directory with PdfSharp's font resolver so they can be embedded.
    /// Scans for .ttf and .otf files and includes common system font folders automatically.
    /// If explicit mappings are registered via RegisterFontMappings they will be used as fallback when a family cannot be resolved from system fonts.
    /// </summary>
    public static void RegisterFontsFromDirectory(string? dir)
    {
        var files = new List<string>();
        if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
        {
            try
            {
                files.AddRange(Directory.EnumerateFiles(dir, "*.ttf"));
                files.AddRange(Directory.EnumerateFiles(dir, "*.otf"));
            }
            catch (Exception ex)
            {
                FontLoadLogger?.Invoke($"Failed to enumerate fonts in '{dir}': {ex.Message}");
            }
        }

        var sysDirs = new List<string>();
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(windir)) sysDirs.Add(Path.Combine(windir, "Fonts"));
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            sysDirs.Add("/System/Library/Fonts");
            sysDirs.Add("/Library/Fonts");
            var home = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            if (!string.IsNullOrEmpty(home)) sysDirs.Add(Path.Combine(home, "Library", "Fonts"));
        }
        else
        {
            sysDirs.Add("/usr/share/fonts");
            sysDirs.Add("/usr/local/share/fonts");
            var home = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            if (!string.IsNullOrEmpty(home)) sysDirs.Add(Path.Combine(home, ".fonts"));
        }

        foreach (var d in sysDirs.Distinct())
        {
            try
            {
                if (Directory.Exists(d))
                {
                    files.AddRange(Directory.EnumerateFiles(d, "*.ttf", SearchOption.AllDirectories));
                    files.AddRange(Directory.EnumerateFiles(d, "*.otf", SearchOption.AllDirectories));
                }
            }
            catch (Exception ex)
            {
                FontLoadLogger?.Invoke($"Failed to enumerate system fonts in '{d}': {ex.Message}");
            }
        }

        var uniqueFiles = files.Where(f => !string.IsNullOrEmpty(f)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        // If no system fonts and no explicit mappings, do not replace the default font resolver
        if ((uniqueFiles.Count == 0) && (s_explicitFamilyMappings == null || s_explicitFamilyMappings.Count == 0))
        {
            FontLoadLogger?.Invoke("No system fonts or explicit mappings found; leaving default font resolver in place.");
            return;
        }

        try
        {
            GlobalFontSettings.FontResolver = new DirectoryFontResolver(uniqueFiles, s_explicitFamilyMappings, FontLoadLogger);
            FontLoadLogger?.Invoke($"Font resolver registered with {uniqueFiles.Count} files (explicit mappings: {s_explicitFamilyMappings?.Count ?? 0}).");
        }
        catch (Exception ex)
        {
            FontLoadLogger?.Invoke($"Failed to register font resolver: {ex.Message}");
        }
    }

    // --- Word helpers (delegated) ---
    public static byte[]? GetImageBytesFromWord(WordprocessingDocument doc, string relationshipId) => WordHelpers.GetImageBytesFromWord(doc, relationshipId);
    public static string? GetRunFontFamily(W.RunProperties? rPr) => WordHelpers.GetRunFontFamily(rPr);
    public static string? GetRunColor(W.RunProperties? rPr) => WordHelpers.GetRunColor(rPr);
    internal static ParagraphFormat GetParagraphFormatting(W.ParagraphProperties? pPr) => WordHelpers.GetParagraphFormatting(pPr);
    public static W.ParagraphProperties? GetDocDefaultsParagraphProperties(MainDocumentPart mainPart) => WordHelpers.GetDocDefaultsParagraphProperties(mainPart);
    public static W.RunProperties? GetDocDefaultsRunProperties(MainDocumentPart mainPart) => WordHelpers.GetDocDefaultsRunProperties(mainPart);
    public static List<double> GetTableGridColumnWidths(W.Table table) => WordHelpers.GetTableGridColumnWidths(table);
    public static W.RunProperties? GetStyleRunProperties(MainDocumentPart mainPart, string styleId) => WordHelpers.GetStyleRunProperties(mainPart, styleId);
    public static W.ParagraphProperties? GetStyleParagraphProperties(MainDocumentPart mainPart, string styleId) => WordHelpers.GetStyleParagraphProperties(mainPart, styleId);
    internal static RunFormat ResolveRunFormatting(MainDocumentPart mainPart, W.Run run, W.Paragraph paragraph) => WordHelpers.ResolveRunFormatting(mainPart, run, paragraph);
    public static (string numFmt, string lvlText, int? startAt) GetNumberingLevelFormat(MainDocumentPart mainPart, string? numId, int ilvl) => WordHelpers.GetNumberingLevelFormat(mainPart, numId, ilvl);
    internal static BorderInfo GetWordCellBorders(W.TableCellProperties? tcPr) => WordHelpers.GetWordCellBorders(tcPr);
    internal static BorderInfo GetWordCellBorders(W.TableCellProperties? tcPr, W.TableProperties? tblPr) => WordHelpers.GetWordCellBorders(tcPr, tblPr);
    public static W.TableBorders? ResolveTableBorders(MainDocumentPart mainPart, W.TableProperties? tblPr) => WordHelpers.ResolveTableBorders(mainPart, tblPr);

    // --- Excel helpers (delegated) ---
    public static IEnumerable<byte[]> GetImagesFromWorksheet(WorksheetPart wsPart) => ExcelHelpers.GetImagesFromWorksheet(wsPart);
    internal static IEnumerable<ExcelHelpers.ExcelImageInfo> GetImagesWithPositionFromWorksheet(WorksheetPart wsPart) => ExcelHelpers.GetImagesWithPositionFromWorksheet(wsPart);
    public static List<(int startRow, int startCol, int endRow, int endCol)> GetMergeCellRanges(S.Worksheet ws) => ExcelHelpers.GetMergeCellRanges(ws);
    public static List<double> GetWorksheetColumnWidths(WorksheetPart wsPart, int maxColumns = 0) => ExcelHelpers.GetWorksheetColumnWidths(wsPart, maxColumns);
    public static string? GetNumberFormatString(WorkbookPart wbPart, uint? numFmtId) => ExcelHelpers.GetNumberFormatString(wbPart, numFmtId);
    internal static ExcelCellStyleInfo GetCellStyleInfo(WorkbookPart wbPart, uint? styleIndex) => ExcelHelpers.GetCellStyleInfo(wbPart, styleIndex);

    // DirectoryFontResolver remains here to configure PdfSharp font embedding
    class DirectoryFontResolver : IFontResolver
    {
        readonly Dictionary<string, byte[]> _fontsByKey = new(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, string> _familyToKey = new(StringComparer.OrdinalIgnoreCase);

        public DirectoryFontResolver(IEnumerable<string> fontFiles, IDictionary<string, string>? explicitMappings, Action<string>? logger)
        {
            // First register provided font files (system and/or directory) so system fonts are preferred
            foreach (var f in fontFiles)
            {
                try
                {
                    if (!File.Exists(f)) { logger?.Invoke($"Font file not found: {f}"); continue; }
                    var bytes = File.ReadAllBytes(f);
                    if (bytes == null || bytes.Length == 0) { logger?.Invoke($"Font file empty: {f}"); continue; }

                    var names = FontUtils.ReadFontNames(bytes);
                    var family = names.TryGetValue(1, out var fam) ? fam : Path.GetFileNameWithoutExtension(f);
                    var sub = names.TryGetValue(2, out var subf2) ? subf2 : "Regular";
                    var key = NormalizeKey(family, sub);

                    if (!_fontsByKey.ContainsKey(key))
                    {
                        _fontsByKey[key] = bytes;
                        if (!_familyToKey.ContainsKey(family)) _familyToKey[family] = key;
                        logger?.Invoke($"Loaded font: {family} ({sub}) from {f}");
                    }

                    // also register filename base as fallback
                    var baseKey = Path.GetFileNameWithoutExtension(f);
                    if (!_fontsByKey.ContainsKey(baseKey)) _fontsByKey[baseKey] = bytes;
                }
                catch (Exception ex)
                {
                    logger?.Invoke($"Skipped unreadable font '{f}': {ex.Message}");
                }
            }

            // Then apply explicit mappings as fallback (do not override existing keys)
            if (explicitMappings != null)
            {
                foreach (var kv in explicitMappings)
                {
                    try
                    {
                        var family = kv.Key;
                        var path = kv.Value;
                        if (!File.Exists(path)) { logger?.Invoke($"Explicit mapping file not found: {path}"); continue; }
                        var bytes = File.ReadAllBytes(path);
                        if (bytes == null || bytes.Length == 0) { logger?.Invoke($"Explicit mapping file empty: {path}"); continue; }
                        var names = FontUtils.ReadFontNames(bytes);
                        var sub = names.TryGetValue(2, out var subf) ? subf : "Regular";
                        var key = NormalizeKey(family, sub);
                        if (!_fontsByKey.ContainsKey(key)) _fontsByKey[key] = bytes;
                        if (!_familyToKey.ContainsKey(family)) _familyToKey[family] = key;
                        logger?.Invoke($"Registered explicit mapping: {family} -> {path}");
                    }
                    catch (Exception ex)
                    {
                        logger?.Invoke($"Failed to register explicit mapping for '{kv.Key}'='{kv.Value}': {ex.Message}");
                    }
                }
            }

            logger?.Invoke($"DirectoryFontResolver initialized: {_fontsByKey.Count} font entries, {_familyToKey.Count} families.");

            // Register embedded Noto Emoji font for cross-platform emoji rendering
            try
            {
                using var stream = typeof(DirectoryFontResolver).Assembly
                    .GetManifestResourceStream("PDFConverter.Fonts.NotoEmoji-Regular.ttf");
                if (stream != null)
                {
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    var emojiBytes = ms.ToArray();
                    var emojiNames = FontUtils.ReadFontNames(emojiBytes);
                    var emFamily = emojiNames.TryGetValue(1, out var ef) ? ef : "Noto Emoji";
                    var emSub = emojiNames.TryGetValue(2, out var es) ? es : "Regular";
                    var emKey = NormalizeKey(emFamily, emSub);
                    if (!_fontsByKey.ContainsKey(emKey))
                    {
                        _fontsByKey[emKey] = emojiBytes;
                        if (!_familyToKey.ContainsKey(emFamily)) _familyToKey[emFamily] = emKey;
                        logger?.Invoke($"Loaded embedded font: {emFamily} ({emSub})");
                    }
                }
            }
            catch (Exception ex)
            {
                logger?.Invoke($"Failed loading embedded emoji font: {ex.Message}");
            }
        }

        static string NormalizeKey(string family, string sub)
        {
            if (string.IsNullOrEmpty(family)) family = "";
            if (string.IsNullOrEmpty(sub)) sub = "Regular";
            return (family + "|" + sub).Trim();
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (string.IsNullOrEmpty(familyName)) familyName = _familyToKey.Keys.FirstOrDefault() ?? string.Empty;
            var style = isBold && isItalic ? "Bold Italic" : isBold ? "Bold" : isItalic ? "Italic" : "Regular";

            var desired = NormalizeKey(familyName, style);
            if (_fontsByKey.ContainsKey(desired)) return new FontResolverInfo(desired);

            // try family + common variants
            var tryKeys = new[] { NormalizeKey(familyName, "Regular"), NormalizeKey(familyName, "Bold"), NormalizeKey(familyName, "Italic"), NormalizeKey(familyName, "Bold Italic") };
            foreach (var k in tryKeys)
                if (_fontsByKey.ContainsKey(k)) return new FontResolverInfo(k);

            // try mapping by family
            if (_familyToKey.TryGetValue(familyName, out var mapped)) return new FontResolverInfo(mapped);

            // substring search
            var found = _fontsByKey.Keys.FirstOrDefault(k => k.IndexOf(familyName, StringComparison.OrdinalIgnoreCase) >= 0);
            if (found != null) return new FontResolverInfo(found);

            // fallback to any font
            var any = _fontsByKey.Keys.FirstOrDefault();
            if (any != null) return new FontResolverInfo(any);

            // as a final fallback, return Arial (font data will be empty and PdfSharp will use internal fallback)
            return new FontResolverInfo("Arial");
        }

        public byte[] GetFont(string faceName)
        {
            if (string.IsNullOrEmpty(faceName)) return Array.Empty<byte>();
            if (_fontsByKey.TryGetValue(faceName, out var b)) return b;
            // try family mapping
            if (_familyToKey.TryGetValue(faceName, out var mapped) && _fontsByKey.TryGetValue(mapped, out var mb)) return mb;
            // try contains
            var kv = _fontsByKey.FirstOrDefault(kvp => kvp.Key.IndexOf(faceName, StringComparison.OrdinalIgnoreCase) >= 0);
            if (!kv.Equals(default(KeyValuePair<string, byte[]>))) return kv.Value;
            return Array.Empty<byte>();
        }
    }
}

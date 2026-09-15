using System.Runtime.InteropServices;
using PdfSharp.Fonts;

namespace PDFConverter;

/// <summary>Font registration and diagnostic hooks shared by both converters.</summary>
public static class OpenXmlHelpers
{
    static readonly Lock Gate = new();
    static Dictionary<string, string>? s_explicitFamilyMappings;

    /// <summary>Receives font loading and registration messages.</summary>
    public static Action<string>? FontLoadLogger { get; set; }

    /// <summary>Receives image and layout diagnostics.</summary>
    public static Action<string>? ImageLoadLogger { get; set; }

    /// <summary>Maps font family names to font files, used when a family is not installed.</summary>
    public static void RegisterFontMappings(IDictionary<string, string> mappings)
    {
        if (mappings == null) return;
        lock (Gate)
        {
            s_explicitFamilyMappings = new Dictionary<string, string>(mappings, StringComparer.OrdinalIgnoreCase);
        }
        FontLoadLogger?.Invoke($"Registered {mappings.Count} explicit font mappings.");
    }

    /// <summary>Installs the font resolver if none is present. Safe to call repeatedly.</summary>
    public static void EnsureFontResolverInitialized()
    {
        if (GlobalFontSettings.FontResolver != null) return;
        lock (Gate)
        {
            if (GlobalFontSettings.FontResolver == null) RegisterFontsFromDirectory(null);
        }
    }

    /// <summary>Registers the fonts in <paramref name="dir"/> plus the platform font folders.</summary>
    public static void RegisterFontsFromDirectory(string? dir)
    {
        var files = new List<string>();
        AddFontFiles(files, dir, SearchOption.TopDirectoryOnly);
        foreach (var systemDirectory in SystemFontDirectories())
            AddFontFiles(files, systemDirectory, SearchOption.AllDirectories);

        var uniqueFiles = files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (uniqueFiles.Count == 0 && (s_explicitFamilyMappings?.Count ?? 0) == 0)
        {
            FontLoadLogger?.Invoke("No system fonts or explicit mappings found; keeping the default font resolver.");
            return;
        }

        try
        {
            GlobalFontSettings.FontResolver =
                new DirectoryFontResolver(uniqueFiles, s_explicitFamilyMappings, FontLoadLogger);
            FontLoadLogger?.Invoke($"Font resolver registered with {uniqueFiles.Count} font files.");
        }
        catch (Exception ex)
        {
            FontLoadLogger?.Invoke($"Failed to register font resolver: {ex.Message}");
        }
    }

    static void AddFontFiles(List<string> files, string? directory, SearchOption option)
    {
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory)) return;
        try
        {
            files.AddRange(Directory.EnumerateFiles(directory, "*.ttf", option));
            files.AddRange(Directory.EnumerateFiles(directory, "*.otf", option));
        }
        catch (Exception ex)
        {
            FontLoadLogger?.Invoke($"Failed to enumerate fonts in '{directory}': {ex.Message}");
        }
    }

    static IEnumerable<string> SystemFontDirectories()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(windows)) yield return Path.Combine(windows, "Fonts");
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (!string.IsNullOrEmpty(localAppData))
                yield return Path.Combine(localAppData, "Microsoft", "Windows", "Fonts");
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            yield return "/System/Library/Fonts";
            yield return "/Library/Fonts";
            if (!string.IsNullOrEmpty(home)) yield return Path.Combine(home, "Library", "Fonts");
        }
        else
        {
            yield return "/usr/share/fonts";
            yield return "/usr/local/share/fonts";
            if (!string.IsNullOrEmpty(home)) yield return Path.Combine(home, ".fonts");
        }
    }
}

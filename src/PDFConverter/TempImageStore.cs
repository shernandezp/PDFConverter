using System.Security.Cryptography;

namespace PDFConverter;

/// <summary>
/// Materialises embedded pictures as temporary files for PdfSharp, keeping one file per distinct
/// image so a picture reused across a document is embedded in the PDF only once.
/// </summary>
internal sealed class TempImageStore : IDisposable
{
    readonly Dictionary<string, string> _pathsByContent = new(StringComparer.Ordinal);

    public string? Save(byte[]? bytes, int cropLeft = 0, int cropTop = 0, int cropRight = 0, int cropBottom = 0)
    {
        if (bytes == null || bytes.Length == 0) return null;

        var key = ContentKey(bytes, cropLeft, cropTop, cropRight, cropBottom);
        if (_pathsByContent.TryGetValue(key, out var existing)) return existing;

        var cropped = ConverterExtensions.ApplySrcRectCrop(bytes, cropLeft, cropTop, cropRight, cropBottom);
        string path;
        try
        {
            path = ConverterExtensions.SaveTempImage(cropped);
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed writing temporary image: {ex.Message}");
            return null;
        }

        _pathsByContent[key] = path;
        return path;
    }

    static string ContentKey(byte[] bytes, int cropLeft, int cropTop, int cropRight, int cropBottom)
    {
        var hash = Convert.ToHexString(SHA256.HashData(bytes));
        return cropLeft == 0 && cropTop == 0 && cropRight == 0 && cropBottom == 0
            ? hash
            : $"{hash}:{cropLeft},{cropTop},{cropRight},{cropBottom}";
    }

    public void Dispose()
    {
        foreach (var path in _pathsByContent.Values)
            ConverterExtensions.TryDeleteTempFile(path);
        _pathsByContent.Clear();
    }
}

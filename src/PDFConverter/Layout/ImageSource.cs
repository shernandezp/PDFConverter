using System.Security.Cryptography;
using SkiaSharp;

namespace PDFConverter.Layout;

internal sealed class ImageSource : IDisposable
{
    readonly byte[] _bytes;
    readonly double _cropLeft, _cropTop, _cropRight, _cropBottom;
    SKImage? _image;
    bool _decoded;

    ImageSource(byte[] bytes, double cropLeft, double cropTop, double cropRight, double cropBottom)
    {
        _bytes = bytes;
        _cropLeft = cropLeft;
        _cropTop = cropTop;
        _cropRight = cropRight;
        _cropBottom = cropBottom;
    }

    public static ImageSource? Create(byte[]? bytes, int cropLeft = 0, int cropTop = 0, int cropRight = 0, int cropBottom = 0)
    {
        if (bytes == null || bytes.Length == 0) return null;
        const double whole = Units.ThousandthsOfPercentPerWhole;
        return new ImageSource(bytes, Fraction(cropLeft / whole), Fraction(cropTop / whole),
            Fraction(cropRight / whole), Fraction(cropBottom / whole));
    }

    static double Fraction(double value) => Math.Clamp(value, 0, 0.99);

    public static string KeyFor(byte[] bytes, int cropLeft, int cropTop, int cropRight, int cropBottom)
    {
        var hash = Convert.ToHexString(SHA256.HashData(bytes));
        return cropLeft == 0 && cropTop == 0 && cropRight == 0 && cropBottom == 0
            ? hash
            : $"{hash}:{cropLeft},{cropTop},{cropRight},{cropBottom}";
    }

    public SKImage? Image
    {
        get
        {
            if (_decoded) return _image;
            _decoded = true;
            _image = Decode(_bytes);
            return _image;
        }
    }

    public SKRect? SourceRect
    {
        get
        {
            var image = Image;
            if (image == null) return null;
            if (_cropLeft == 0 && _cropTop == 0 && _cropRight == 0 && _cropBottom == 0) return null;

            var left = (float)(image.Width * _cropLeft);
            var top = (float)(image.Height * _cropTop);
            var right = (float)(image.Width * (1 - _cropRight));
            var bottom = (float)(image.Height * (1 - _cropBottom));
            return new SKRect(left, top, Math.Max(right, left + 1), Math.Max(bottom, top + 1));
        }
    }

    public double HeightFor(double width)
    {
        var image = Image;
        if (image == null || image.Width == 0) return width;
        var rect = SourceRect;
        var pixelWidth = rect?.Width ?? image.Width;
        var pixelHeight = rect?.Height ?? image.Height;
        return pixelWidth <= 0 ? width : width * pixelHeight / pixelWidth;
    }

    static SKImage? Decode(byte[] bytes)
    {
        try
        {
            var image = SKImage.FromEncodedData(bytes);
            if (image != null) return image;
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Image decoding failed: {ex.Message}");
        }

        var extension = ConverterExtensions.DetectImageExtension(bytes);
        if (extension == ".tif" && TiffDecoder.Decode(bytes) is { } tiff)
        {
            using (tiff) return SKImage.FromBitmap(tiff);
        }
        if (extension is ".emf" or ".wmf" && MetafileImages.ExtractRaster(bytes) is { } raster
            && SKImage.FromEncodedData(raster) is { } embedded)
            return embedded;

        var converted = ConverterExtensions.ReencodeAsPng(bytes);
        if (converted == null)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Unsupported image format '{extension}' skipped.");
            return null;
        }
        return SKImage.FromEncodedData(converted);
    }

    public void Dispose()
    {
        _image?.Dispose();
        _image = null;
    }
}

internal sealed class ImageStore : IDisposable
{
    readonly Dictionary<string, ImageSource> _byKey = new(StringComparer.Ordinal);

    public ImageSource? Get(byte[]? bytes, int cropLeft = 0, int cropTop = 0, int cropRight = 0, int cropBottom = 0)
    {
        if (bytes == null || bytes.Length == 0) return null;

        var key = ImageSource.KeyFor(bytes, cropLeft, cropTop, cropRight, cropBottom);
        if (_byKey.TryGetValue(key, out var existing)) return existing;

        var source = ImageSource.Create(bytes, cropLeft, cropTop, cropRight, cropBottom);
        if (source != null) _byKey[key] = source;
        return source;
    }

    public void Dispose()
    {
        foreach (var source in _byKey.Values) source.Dispose();
        _byKey.Clear();
    }
}

using System.Buffers.Binary;
using PDFConverter.Layout;
using SkiaSharp;
using Xunit;

namespace PDFConverter.Tests;

public class ImageDecoderTests
{
    static byte[] Tiff(int width, int height, byte[] rgb, ushort compression = 1)
    {
        var entries = new List<(ushort Tag, ushort Type, uint Count, uint Value)>
        {
            (256, 3, 1, (uint)width), (257, 3, 1, (uint)height), (258, 3, 3, 0), (259, 3, 1, compression),
            (262, 3, 1, 2), (273, 4, 1, 0), (277, 3, 1, 3), (278, 3, 1, (uint)height), (279, 4, 1, (uint)rgb.Length),
        };
        var ifdOffset = 8;
        var ifdSize = 2 + entries.Count * 12 + 4;
        var bitsOffset = ifdOffset + ifdSize;
        var dataOffset = bitsOffset + 6;
        var file = new byte[dataOffset + rgb.Length];
        file[0] = (byte)'I';
        file[1] = (byte)'I';
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(2), 42);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(4), (uint)ifdOffset);
        BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(ifdOffset), (ushort)entries.Count);
        for (var i = 0; i < entries.Count; i++)
        {
            var (tag, type, count, value) = entries[i];
            if (tag == 258) value = (uint)bitsOffset;
            if (tag == 273) value = (uint)dataOffset;
            var at = ifdOffset + 2 + i * 12;
            BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(at), tag);
            BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(at + 2), type);
            BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(at + 4), count);
            if (type == 3 && count == 1) BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(at + 8), (ushort)value);
            else BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(at + 8), value);
        }
        for (var i = 0; i < 3; i++) BinaryPrimitives.WriteUInt16LittleEndian(file.AsSpan(bitsOffset + i * 2), 8);
        rgb.CopyTo(file, dataOffset);
        return file;
    }

    [Fact]
    public void Tiff_UncompressedRgb_IsDecoded()
    {
        var pixels = new byte[] { 255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255 };

        using var bitmap = TiffDecoder.Decode(Tiff(2, 2, pixels));

        Assert.NotNull(bitmap);
        Assert.Equal(2, bitmap.Width);
        Assert.Equal(SKColors.Red, bitmap.GetPixel(0, 0));
        Assert.Equal(SKColors.Lime, bitmap.GetPixel(1, 0));
        Assert.Equal(SKColors.Blue, bitmap.GetPixel(0, 1));
        Assert.Equal(SKColors.White, bitmap.GetPixel(1, 1));
    }

    [Fact]
    public void Tiff_PackBits_IsDecoded()
    {
        var packed = new byte[] { 0xFE, 0x80, 0x02, 0x10, 0x20, 0x30 };

        using var bitmap = TiffDecoder.Decode(Tiff(2, 1, packed, compression: 32773));

        Assert.NotNull(bitmap);
        Assert.Equal(new SKColor(0x80, 0x80, 0x80), bitmap.GetPixel(0, 0));
        Assert.Equal(new SKColor(0x10, 0x20, 0x30), bitmap.GetPixel(1, 0));
    }

    [Fact]
    public void Tiff_Garbage_ReturnsNull()
    {
        Assert.Null(TiffDecoder.Decode([1, 2, 3]));
        Assert.Null(TiffDecoder.Decode("II*\0garbage"u8.ToArray()));
    }

    static byte[] Dib(int width, int height, byte[] bgr)
    {
        var header = new byte[40];
        BinaryPrimitives.WriteUInt32LittleEndian(header, 40);
        BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(4), width);
        BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(8), height);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(12), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(14), 24);
        return [.. header, .. bgr];
    }

    static byte[] Emf(byte[] dibHeaderAndBits)
    {
        var header = new byte[108];
        BinaryPrimitives.WriteUInt32LittleEndian(header, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(4), 108);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(40), 0x464D4520);

        var record = new byte[80 + dibHeaderAndBits.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(record, 81);
        BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(4), (uint)record.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(48), 80);
        BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(52), 40);
        BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(56), 120);
        BinaryPrimitives.WriteUInt32LittleEndian(record.AsSpan(60), (uint)(dibHeaderAndBits.Length - 40));
        dibHeaderAndBits.CopyTo(record, 80);
        return [.. header, .. record];
    }

    [Fact]
    public void Emf_WithStretchDiBitsRecord_YieldsADecodableBitmap()
    {
        var dib = Dib(1, 1, [0, 0, 255, 0]);

        var raster = MetafileImages.ExtractRaster(Emf(dib));

        Assert.NotNull(raster);
        using var image = SKImage.FromEncodedData(raster);
        Assert.NotNull(image);
        Assert.Equal(1, image.Width);
        using var bitmap = SKBitmap.FromImage(image);
        Assert.Equal(SKColors.Red, bitmap.GetPixel(0, 0));
    }

    [Fact]
    public void Emf_WithEmbeddedPng_YieldsThePng()
    {
        using var source = new SKBitmap(3, 2);
        source.Erase(SKColors.Green);
        using var encoded = SKImage.FromBitmap(source).Encode(SKEncodedImageFormat.Png, 100);
        var png = encoded.ToArray();

        var header = new byte[108];
        BinaryPrimitives.WriteUInt32LittleEndian(header, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(4), 108);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(40), 0x464D4520);
        var comment = new byte[12 + png.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(comment, 70);
        BinaryPrimitives.WriteUInt32LittleEndian(comment.AsSpan(4), (uint)comment.Length);
        png.CopyTo(comment, 12);

        var raster = MetafileImages.ExtractRaster([.. header, .. comment]);

        Assert.NotNull(raster);
        using var image = SKImage.FromEncodedData(raster);
        Assert.Equal(3, image!.Width);
    }

    [Fact]
    public void Metafile_WithoutRaster_ReturnsNull()
    {
        var header = new byte[108];
        BinaryPrimitives.WriteUInt32LittleEndian(header, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(4), 108);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(40), 0x464D4520);

        Assert.Null(MetafileImages.ExtractRaster(header));
    }
}

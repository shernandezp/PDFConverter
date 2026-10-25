using System.Buffers.Binary;

namespace PDFConverter.Layout;

internal static class MetafileImages
{
    static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    static readonly byte[] JpegSignature = [0xFF, 0xD8, 0xFF];

    public static byte[]? ExtractRaster(byte[] data)
    {
        try
        {
            if (data.Length > 44 && BinaryPrimitives.ReadUInt32LittleEndian(data) == 1
                && BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(40)) == 0x464D4520)
                return FromEmf(data);
            if (data.Length > 22 && BinaryPrimitives.ReadUInt32LittleEndian(data) == 0x9AC6CDD7)
                return FromWmf(data, 22);
            if (data.Length > 18 && BinaryPrimitives.ReadUInt16LittleEndian(data) is 1 or 2)
                return FromWmf(data, 0);
        }
        catch
        {
        }
        return null;
    }

    static byte[]? FromEmf(byte[] data)
    {
        var position = 0;
        byte[]? best = null;
        var bestPixels = 0L;
        while (position + 8 <= data.Length)
        {
            var type = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(position));
            var size = (int)BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(position + 4));
            if (size < 8 || position + size > data.Length) break;
            var record = data.AsSpan(position, size);

            byte[]? candidate = type switch
            {
                76 => Dib(record, 84, 88, 92, 96),
                77 => Dib(record, 84, 88, 92, 96),
                80 => Dib(record, 48, 52, 56, 60),
                81 => Dib(record, 48, 52, 56, 60),
                70 => EmbeddedEncoded(record),
                _ => null,
            };
            if (candidate != null)
            {
                var pixels = Pixels(candidate);
                if (pixels > bestPixels)
                {
                    best = candidate;
                    bestPixels = pixels;
                }
            }
            position += size;
        }
        return best;
    }

    static byte[]? FromWmf(byte[] data, int start)
    {
        var position = start + 18;
        byte[]? best = null;
        var bestPixels = 0L;
        while (position + 6 <= data.Length)
        {
            var size = (int)(BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(position)) * 2);
            var function = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(position + 4));
            if (size < 6 || position + size > data.Length) break;

            var dibOffset = function switch
            {
                0x0B41 => 6 + 2 * 10,
                0x0F43 => 6 + 2 * 11,
                0x0D33 => 6 + 2 * 9,
                _ => -1,
            };
            if (dibOffset > 0 && dibOffset < size)
            {
                var candidate = BmpFromDib(data.AsSpan(position + dibOffset, size - dibOffset));
                var pixels = Pixels(candidate);
                if (candidate != null && pixels > bestPixels)
                {
                    best = candidate;
                    bestPixels = pixels;
                }
            }
            position += size;
        }
        return best;
    }

    static byte[]? Dib(ReadOnlySpan<byte> record, int offBmi, int cbBmi, int offBits, int cbBits)
    {
        if (record.Length < offBits + 4) return null;
        var headerOffset = (int)BinaryPrimitives.ReadUInt32LittleEndian(record[offBmi..]);
        var headerSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(record[cbBmi..]);
        var bitsOffset = (int)BinaryPrimitives.ReadUInt32LittleEndian(record[offBits..]);
        var bitsSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(record[cbBits..]);
        if (headerSize <= 0 || bitsSize <= 0 || headerOffset + headerSize > record.Length || bitsOffset + bitsSize > record.Length)
            return null;

        var header = record.Slice(headerOffset, headerSize);
        var bits = record.Slice(bitsOffset, bitsSize);
        return Bmp(header, bits);
    }

    static byte[]? BmpFromDib(ReadOnlySpan<byte> dib)
    {
        if (dib.Length < 40) return null;
        var headerSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(dib);
        if (headerSize < 40 || headerSize > dib.Length) return null;
        var bitCount = BinaryPrimitives.ReadUInt16LittleEndian(dib[14..]);
        var colorsUsed = (int)BinaryPrimitives.ReadUInt32LittleEndian(dib[32..]);
        var paletteEntries = bitCount <= 8 ? (colorsUsed > 0 ? colorsUsed : 1 << bitCount) : 0;
        var compression = BinaryPrimitives.ReadUInt32LittleEndian(dib[16..]);
        var masks = compression == 3 ? 12 : 0;
        var headerTotal = headerSize + masks + paletteEntries * 4;
        if (headerTotal > dib.Length) return null;
        return Bmp(dib[..headerTotal], dib[headerTotal..]);
    }

    static byte[] Bmp(ReadOnlySpan<byte> header, ReadOnlySpan<byte> bits)
    {
        var file = new byte[14 + header.Length + bits.Length];
        file[0] = (byte)'B';
        file[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(2), (uint)file.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(10), (uint)(14 + header.Length));
        header.CopyTo(file.AsSpan(14));
        bits.CopyTo(file.AsSpan(14 + header.Length));
        return file;
    }

    static byte[]? EmbeddedEncoded(ReadOnlySpan<byte> record)
    {
        var png = record.IndexOf(PngSignature);
        if (png >= 0) return record[png..].ToArray();
        var jpeg = record.IndexOf(JpegSignature);
        return jpeg >= 0 ? record[jpeg..].ToArray() : null;
    }

    static long Pixels(byte[]? image)
    {
        if (image == null || image.Length < 26) return 0;
        if (image[0] == 'B' && image[1] == 'M')
            return Math.Abs((long)BinaryPrimitives.ReadInt32LittleEndian(image.AsSpan(18)) * BinaryPrimitives.ReadInt32LittleEndian(image.AsSpan(22)));
        if (image.AsSpan().StartsWith(PngSignature))
            return (long)BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(16)) * BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(20));
        return 1;
    }
}

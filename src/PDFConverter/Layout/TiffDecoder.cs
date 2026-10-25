using System.Buffers.Binary;
using SkiaSharp;

namespace PDFConverter.Layout;

internal static class TiffDecoder
{
    const int Compression = 259;
    const int Photometric = 262;
    const int StripOffsets = 273;
    const int SamplesPerPixel = 277;
    const int RowsPerStrip = 278;
    const int StripByteCounts = 279;
    const int PlanarConfiguration = 284;
    const int Predictor = 317;
    const int ColorMap = 320;

    public static SKBitmap? Decode(byte[] data)
    {
        try
        {
            return DecodeCore(data);
        }
        catch
        {
            return null;
        }
    }

    static SKBitmap? DecodeCore(byte[] data)
    {
        if (data.Length < 8) return null;
        var little = data[0] == 'I' && data[1] == 'I';
        if (!little && !(data[0] == 'M' && data[1] == 'M')) return null;
        var reader = new Reader(data, little);
        if (reader.U16(2) != 42) return null;

        var ifd = (int)reader.U32(4);
        var entries = reader.U16(ifd);
        var tags = new Dictionary<int, (int Type, int Count, int ValueOffset)>();
        for (var i = 0; i < entries; i++)
        {
            var entry = ifd + 2 + i * 12;
            tags[reader.U16(entry)] = (reader.U16(entry + 2), (int)reader.U32(entry + 4), entry + 8);
        }

        var width = (int)reader.Value(tags, 256);
        var height = (int)reader.Value(tags, 257);
        if (width <= 0 || height <= 0 || (long)width * height > 50_000_000) return null;

        var bitsPerSample = reader.Values(tags, 258, 1);
        var bits = (int)bitsPerSample[0];
        var samples = (int)reader.Value(tags, SamplesPerPixel, 1);
        var compression = (int)reader.Value(tags, Compression, 1);
        var photometric = (int)reader.Value(tags, Photometric, samples >= 3 ? 2 : 1);
        var planar = (int)reader.Value(tags, PlanarConfiguration, 1);
        var predictor = (int)reader.Value(tags, Predictor, 1);
        var rowsPerStrip = (int)Math.Min(reader.Value(tags, RowsPerStrip, height), height);
        var offsets = reader.Values(tags, StripOffsets, 0);
        var counts = reader.Values(tags, StripByteCounts, 0);
        if (planar != 1 || offsets.Length == 0 || counts.Length != offsets.Length) return null;
        if (bits is not (1 or 4 or 8 or 16) || samples is not (1 or 2 or 3 or 4)) return null;
        if (compression is not (1 or 5 or 32773)) return null;

        var bytesPerRow = (width * samples * bits + 7) / 8;
        var pixels = new byte[bytesPerRow * height];
        var row = 0;
        for (var strip = 0; strip < offsets.Length && row < height; strip++)
        {
            var raw = data.AsSpan((int)offsets[strip], (int)counts[strip]);
            var rows = Math.Min(rowsPerStrip, height - row);
            var expected = rows * bytesPerRow;
            var decoded = compression switch
            {
                1 => raw.ToArray(),
                5 => Lzw.Decode(raw, expected),
                _ => PackBits.Decode(raw, expected),
            };
            if (decoded.Length < expected) Array.Resize(ref decoded, expected);
            decoded.AsSpan(0, expected).CopyTo(pixels.AsSpan(row * bytesPerRow));
            row += rows;
        }

        if (predictor == 2 && bits == 8)
        {
            for (var y = 0; y < height; y++)
            {
                var line = pixels.AsSpan(y * bytesPerRow, bytesPerRow);
                for (var i = samples; i < line.Length; i++) line[i] += line[i - samples];
            }
        }

        uint[]? palette = null;
        if (photometric == 3 && tags.ContainsKey(ColorMap))
        {
            var map = reader.Values(tags, ColorMap, 0);
            var size = map.Length / 3;
            palette = new uint[size];
            for (var i = 0; i < size; i++)
                palette[i] = Argb((byte)(map[i] >> 8), (byte)(map[size + i] >> 8), (byte)(map[2 * size + i] >> 8));
        }

        var bitmap = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Unpremul);
        var output = new uint[width * height];
        for (var y = 0; y < height; y++)
        {
            var line = pixels.AsSpan(y * bytesPerRow, bytesPerRow);
            for (var x = 0; x < width; x++)
                output[y * width + x] = Pixel(line, x, bits, samples, photometric, palette);
        }
        System.Runtime.InteropServices.MemoryMarshal.AsBytes(output.AsSpan()).CopyTo(bitmap.GetPixelSpan());
        return bitmap;
    }

    static uint Pixel(ReadOnlySpan<byte> line, int x, int bits, int samples, int photometric, uint[]? palette)
    {
        var channel = new int[samples];
        for (var s = 0; s < samples; s++) channel[s] = Sample(line, x * samples + s, bits);
        var max = (1 << Math.Min(bits, 8)) - 1;

        switch (photometric)
        {
            case 0:
            {
                var gray = (byte)(255 - channel[0] * 255 / max);
                return Argb(gray, gray, gray, samples == 2 ? (byte)(channel[1] * 255 / max) : (byte)255);
            }
            case 1:
            {
                var gray = (byte)(channel[0] * 255 / max);
                return Argb(gray, gray, gray, samples == 2 ? (byte)(channel[1] * 255 / max) : (byte)255);
            }
            case 3:
                return palette != null && channel[0] < palette.Length ? palette[channel[0]] : Argb(0, 0, 0);
            default:
            {
                if (samples < 3) return Argb((byte)channel[0], (byte)channel[0], (byte)channel[0]);
                var alpha = samples == 4 ? (byte)(channel[3] * 255 / max) : (byte)255;
                return Argb((byte)(channel[0] * 255 / max), (byte)(channel[1] * 255 / max), (byte)(channel[2] * 255 / max), alpha);
            }
        }
    }

    static int Sample(ReadOnlySpan<byte> line, int index, int bits) => bits switch
    {
        1 => (line[index / 8] >> (7 - index % 8)) & 1,
        4 => (line[index / 2] >> (index % 2 == 0 ? 4 : 0)) & 0xF,
        8 => line[index],
        _ => line[index * 2],
    };

    static uint Argb(byte r, byte g, byte b, byte a = 255) => (uint)(a << 24 | r << 16 | g << 8 | b);

    readonly struct Reader(byte[] data, bool little)
    {
        public ushort U16(int offset) => little
            ? BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(offset))
            : BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(offset));

        public uint U32(int offset) => little
            ? BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(offset))
            : BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(offset));

        public long Value(Dictionary<int, (int Type, int Count, int ValueOffset)> tags, int tag, long fallback = 0)
        {
            var values = Values(tags, tag, fallback);
            return values.Length > 0 ? values[0] : fallback;
        }

        public long[] Values(Dictionary<int, (int Type, int Count, int ValueOffset)> tags, int tag, long fallback)
        {
            if (!tags.TryGetValue(tag, out var entry)) return fallback == 0 ? [] : [fallback];
            var size = entry.Type switch { 1 or 2 or 6 or 7 => 1, 3 or 8 => 2, 4 or 9 => 4, _ => 0 };
            if (size == 0) return [fallback];
            var total = size * entry.Count;
            var offset = total <= 4 ? entry.ValueOffset : (int)U32(entry.ValueOffset);
            var values = new long[entry.Count];
            for (var i = 0; i < entry.Count; i++)
            {
                var at = offset + i * size;
                values[i] = size switch { 1 => data[at], 2 => U16(at), _ => U32(at) };
            }
            return values;
        }
    }

    static class PackBits
    {
        public static byte[] Decode(ReadOnlySpan<byte> input, int expected)
        {
            var output = new List<byte>(expected);
            var i = 0;
            while (i < input.Length && output.Count < expected)
            {
                var n = (sbyte)input[i++];
                if (n >= 0)
                {
                    for (var k = 0; k <= n && i < input.Length; k++) output.Add(input[i++]);
                }
                else if (n != -128 && i < input.Length)
                {
                    var value = input[i++];
                    for (var k = 0; k < 1 - n; k++) output.Add(value);
                }
            }
            return output.ToArray();
        }
    }

    static class Lzw
    {
        public static byte[] Decode(ReadOnlySpan<byte> input, int expected)
        {
            var output = new List<byte>(expected);
            var table = new List<byte[]>();
            void Reset()
            {
                table.Clear();
                for (var i = 0; i < 256; i++) table.Add([(byte)i]);
                table.Add([]);
                table.Add([]);
            }
            Reset();

            var codeSize = 9;
            long bitBuffer = 0;
            var bitCount = 0;
            var position = 0;
            byte[]? previous = null;

            while (output.Count < expected)
            {
                while (bitCount < codeSize && position < input.Length)
                {
                    bitBuffer = (bitBuffer << 8) | input[position++];
                    bitCount += 8;
                }
                if (bitCount < codeSize) break;
                var code = (int)((bitBuffer >> (bitCount - codeSize)) & ((1 << codeSize) - 1));
                bitCount -= codeSize;

                if (code == 256)
                {
                    Reset();
                    codeSize = 9;
                    previous = null;
                    continue;
                }
                if (code == 257) break;

                byte[] entry;
                if (code < table.Count)
                {
                    entry = table[code];
                    if (previous != null) table.Add([.. previous, entry[0]]);
                }
                else if (previous != null)
                {
                    entry = [.. previous, previous[0]];
                    table.Add(entry);
                }
                else
                {
                    break;
                }

                output.AddRange(entry);
                previous = entry;
                if (table.Count + 1 >= 1 << codeSize && codeSize < 12) codeSize++;
            }
            return output.ToArray();
        }
    }
}

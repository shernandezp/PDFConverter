using System.Buffers.Binary;

namespace PDFConverter.Layout;

internal static class FontSubsetter
{
    static readonly string[] KeptTables = ["OS/2", "cmap", "head", "hhea", "hmtx", "maxp", "name"];

    public static byte[]? StripUnusedGlyphs(byte[] font, IReadOnlySet<ushort> usedGlyphs, int collectionIndex = 0)
    {
        try
        {
            return Strip(font, usedGlyphs, collectionIndex);
        }
        catch
        {
            return null;
        }
    }

    static byte[]? Strip(byte[] font, IReadOnlySet<ushort> usedGlyphs, int collectionIndex)
    {
        var data = font.AsSpan();
        var directory = 0;
        if (BinaryPrimitives.ReadUInt32BigEndian(data) == 0x74746366)
        {
            var fontCount = (int)BinaryPrimitives.ReadUInt32BigEndian(data[8..]);
            if (collectionIndex < 0 || collectionIndex >= fontCount) return null;
            directory = (int)BinaryPrimitives.ReadUInt32BigEndian(data[(12 + collectionIndex * 4)..]);
            if (directory < 0 || directory + 12 > data.Length) return null;
        }

        var version = BinaryPrimitives.ReadUInt32BigEndian(data[directory..]);
        if (version != 0x00010000 && version != 0x74727565) return null;

        var tableCount = BinaryPrimitives.ReadUInt16BigEndian(data[(directory + 4)..]);
        var tables = new Dictionary<string, (int Offset, int Length)>(StringComparer.Ordinal);
        for (var i = 0; i < tableCount; i++)
        {
            var record = data[(directory + 12 + i * 16)..];
            var tag = System.Text.Encoding.ASCII.GetString(record[..4]);
            var offset = (int)BinaryPrimitives.ReadUInt32BigEndian(record[8..]);
            var length = (int)BinaryPrimitives.ReadUInt32BigEndian(record[12..]);
            if (offset < 0 || length < 0 || offset + length > data.Length) return null;
            tables[tag] = (offset, length);
        }

        if (!tables.TryGetValue("glyf", out var glyf) || !tables.TryGetValue("loca", out var loca)
            || !tables.TryGetValue("head", out var head) || !tables.TryGetValue("maxp", out var maxp))
            return null;

        var longLoca = BinaryPrimitives.ReadInt16BigEndian(data[(head.Offset + 50)..]) == 1;
        var glyphCount = BinaryPrimitives.ReadUInt16BigEndian(data[(maxp.Offset + 4)..]);
        var offsets = ReadLoca(data.Slice(loca.Offset, loca.Length), glyphCount, longLoca);
        if (offsets == null) return null;

        var keep = new HashSet<ushort> { 0 };
        var pending = new Stack<ushort>(usedGlyphs);
        while (pending.Count > 0)
        {
            var glyph = pending.Pop();
            if (glyph >= glyphCount || !keep.Add(glyph)) continue;
            foreach (var component in Components(data.Slice(glyf.Offset, glyf.Length), offsets, glyph))
                pending.Push(component);
        }

        var newGlyf = new MemoryStream();
        var newOffsets = new uint[glyphCount + 1];
        for (var glyph = 0; glyph < glyphCount; glyph++)
        {
            newOffsets[glyph] = (uint)newGlyf.Length;
            if (!keep.Contains((ushort)glyph)) continue;

            var start = (int)offsets[glyph];
            var length = (int)(offsets[glyph + 1] - offsets[glyph]);
            if (length <= 0 || start + length > glyf.Length) continue;

            newGlyf.Write(data.Slice(glyf.Offset + start, length));
            while (newGlyf.Length % 4 != 0) newGlyf.WriteByte(0);
        }
        newOffsets[glyphCount] = (uint)newGlyf.Length;

        var newLoca = new byte[(glyphCount + 1) * 4];
        for (var i = 0; i <= glyphCount; i++)
            BinaryPrimitives.WriteUInt32BigEndian(newLoca.AsSpan(i * 4), newOffsets[i]);

        var headCopy = data.Slice(head.Offset, head.Length).ToArray();
        BinaryPrimitives.WriteInt16BigEndian(headCopy.AsSpan(50), 1);
        BinaryPrimitives.WriteUInt32BigEndian(headCopy.AsSpan(8), 0);

        var output = new List<(string Tag, byte[] Data)>
        {
            ("glyf", newGlyf.ToArray()),
            ("loca", newLoca),
            ("head", headCopy),
            ("post", [0, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
        };
        foreach (var tag in KeptTables)
            if (tables.TryGetValue(tag, out var table))
                output.Add((tag, data.Slice(table.Offset, table.Length).ToArray()));

        return Assemble(output);
    }

    static uint[]? ReadLoca(ReadOnlySpan<byte> loca, int glyphCount, bool longFormat)
    {
        var entrySize = longFormat ? 4 : 2;
        if (loca.Length < (glyphCount + 1) * entrySize) return null;

        var offsets = new uint[glyphCount + 1];
        for (var i = 0; i <= glyphCount; i++)
        {
            offsets[i] = longFormat
                ? BinaryPrimitives.ReadUInt32BigEndian(loca[(i * 4)..])
                : (uint)BinaryPrimitives.ReadUInt16BigEndian(loca[(i * 2)..]) * 2;
        }
        return offsets;
    }

    static List<ushort> Components(ReadOnlySpan<byte> glyf, uint[] offsets, ushort glyph)
    {
        var components = new List<ushort>();
        var start = (int)offsets[glyph];
        var length = (int)(offsets[glyph + 1] - offsets[glyph]);
        if (length < 10 || start + length > glyf.Length) return components;

        var contours = BinaryPrimitives.ReadInt16BigEndian(glyf[start..]);
        if (contours >= 0) return components;

        var position = start + 10;
        while (position + 4 <= start + length)
        {
            var flags = BinaryPrimitives.ReadUInt16BigEndian(glyf[position..]);
            components.Add(BinaryPrimitives.ReadUInt16BigEndian(glyf[(position + 2)..]));
            position += 4;
            position += (flags & 0x0001) != 0 ? 4 : 2;
            if ((flags & 0x0008) != 0) position += 2;
            else if ((flags & 0x0040) != 0) position += 4;
            else if ((flags & 0x0080) != 0) position += 8;
            if ((flags & 0x0020) == 0) break;
        }
        return components;
    }

    static byte[] Assemble(List<(string Tag, byte[] Data)> tables)
    {
        tables.Sort((a, b) => string.CompareOrdinal(a.Tag, b.Tag));

        var count = tables.Count;
        var entrySelector = 0;
        while ((2 << entrySelector) <= count) entrySelector++;
        var searchRange = (1 << entrySelector) * 16;

        var directorySize = 12 + count * 16;
        var totalSize = directorySize;
        foreach (var (_, table) in tables) totalSize += (table.Length + 3) & ~3;

        var output = new byte[totalSize];
        var span = output.AsSpan();
        BinaryPrimitives.WriteUInt32BigEndian(span, 0x00010000);
        BinaryPrimitives.WriteUInt16BigEndian(span[4..], (ushort)count);
        BinaryPrimitives.WriteUInt16BigEndian(span[6..], (ushort)searchRange);
        BinaryPrimitives.WriteUInt16BigEndian(span[8..], (ushort)entrySelector);
        BinaryPrimitives.WriteUInt16BigEndian(span[10..], (ushort)(count * 16 - searchRange));

        var offset = directorySize;
        var headOffset = -1;
        for (var i = 0; i < count; i++)
        {
            var (tag, table) = tables[i];
            var record = span[(12 + i * 16)..];
            System.Text.Encoding.ASCII.GetBytes(tag.PadRight(4), record[..4]);
            table.CopyTo(span[offset..]);
            BinaryPrimitives.WriteUInt32BigEndian(record[4..], Checksum(span.Slice(offset, (table.Length + 3) & ~3)));
            BinaryPrimitives.WriteUInt32BigEndian(record[8..], (uint)offset);
            BinaryPrimitives.WriteUInt32BigEndian(record[12..], (uint)table.Length);
            if (tag == "head") headOffset = offset;
            offset += (table.Length + 3) & ~3;
        }

        if (headOffset >= 0)
            BinaryPrimitives.WriteUInt32BigEndian(span[(headOffset + 8)..], 0xB1B0AFBA - Checksum(span));

        return output;
    }

    static uint Checksum(ReadOnlySpan<byte> data)
    {
        uint sum = 0;
        var i = 0;
        for (; i + 4 <= data.Length; i += 4) sum += BinaryPrimitives.ReadUInt32BigEndian(data[i..]);
        if (i < data.Length)
        {
            Span<byte> tail = stackalloc byte[4];
            data[i..].CopyTo(tail);
            sum += BinaryPrimitives.ReadUInt32BigEndian(tail);
        }
        return sum;
    }
}

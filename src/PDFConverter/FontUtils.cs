using System.Text;

namespace PDFConverter;

/// <summary>Reads the name table of a TrueType/OpenType font.</summary>
public static class FontUtils
{
    /// <summary>Maps nameId to value (1 = family, 2 = subfamily). Returns an empty map on failure.</summary>
    public static Dictionary<int, string> ReadFontNames(byte[] data)
    {
        using var stream = new MemoryStream(data);
        return ReadFontNames(stream);
    }

    internal static Dictionary<int, string> ReadFontNames(Stream stream)
    {
        var result = new Dictionary<int, string>();
        try
        {
            using var reader = new BinaryReader(stream, Encoding.ASCII, leaveOpen: true);

            ReadUInt32BE(reader);
            var tableCount = ReadUInt16BE(reader);
            reader.ReadBytes(6);

            uint nameTableOffset = 0;
            for (var i = 0; i < tableCount; i++)
            {
                var tag = Encoding.ASCII.GetString(reader.ReadBytes(4));
                ReadUInt32BE(reader);
                var offset = ReadUInt32BE(reader);
                ReadUInt32BE(reader);
                if (tag == "name") nameTableOffset = offset;
            }
            if (nameTableOffset == 0) return result;

            stream.Position = nameTableOffset;
            ReadUInt16BE(reader);
            var recordCount = ReadUInt16BE(reader);
            var storageOffset = ReadUInt16BE(reader);

            var records = new List<(ushort PlatformId, ushort NameId, ushort Length, ushort Offset)>(recordCount);
            for (var i = 0; i < recordCount; i++)
            {
                var platformId = ReadUInt16BE(reader);
                ReadUInt16BE(reader);
                ReadUInt16BE(reader);
                var nameId = ReadUInt16BE(reader);
                var length = ReadUInt16BE(reader);
                var offset = ReadUInt16BE(reader);
                records.Add((platformId, nameId, length, offset));
            }

            foreach (var record in records)
            {
                if (result.ContainsKey(record.NameId)) continue;
                stream.Position = nameTableOffset + storageOffset + record.Offset;
                var raw = reader.ReadBytes(record.Length);
                var encoding = record.PlatformId is 0 or 3 ? Encoding.BigEndianUnicode : Encoding.Latin1;
                result[record.NameId] = encoding.GetString(raw).Trim('\0').Trim();
            }
        }
        catch
        {
            return result;
        }
        return result;
    }

    static ushort ReadUInt16BE(BinaryReader reader)
    {
        var bytes = reader.ReadBytes(2);
        return (ushort)((bytes[0] << 8) | bytes[1]);
    }

    static uint ReadUInt32BE(BinaryReader reader)
    {
        var bytes = reader.ReadBytes(4);
        return ((uint)bytes[0] << 24) | ((uint)bytes[1] << 16) | ((uint)bytes[2] << 8) | bytes[3];
    }
}

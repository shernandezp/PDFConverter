namespace PDFConverter;

/// <summary>
/// Utility helpers for font processing and name table parsing.
/// </summary>
public static class FontUtils
{
    /// <summary>
    /// Read simple name records from a TrueType/OpenType font byte array.
    /// Returns a mapping of nameId -> string (e.g., 1 = family, 2 = subfamily).
    /// This is defensive and returns an empty map on failure.
    /// </summary>
    public static Dictionary<int, string> ReadFontNames(byte[] data)
    {
        var result = new Dictionary<int, string>();
        try
        {
            using var ms = new MemoryStream(data);
            using var br = new BinaryReader(ms);
            // Offset table
            ReadUInt32BE(br);
            var numTables = ReadUInt16BE(br);
            br.ReadUInt16(); // searchRange
            br.ReadUInt16(); // entrySelector
            br.ReadUInt16(); // rangeShift

            uint nameOffset = 0;
            for (int i = 0; i < numTables; i++)
            {
                var tagChars = br.ReadChars(4);
                var tag = new string(tagChars);
                var checkSum = ReadUInt32BE(br);
                var offset = ReadUInt32BE(br);
                var length = ReadUInt32BE(br);
                if (tag == "name") nameOffset = offset;
            }

            if (nameOffset == 0) return result;
            ms.Position = nameOffset;
            var format = ReadUInt16BE(br);
            var count = ReadUInt16BE(br);
            var stringOffset = ReadUInt16BE(br);

            var records = new List<(ushort platformId, ushort encodingId, ushort languageId, ushort nameId, ushort length, ushort offset)>();
            for (int i = 0; i < count; i++)
            {
                var platformId = ReadUInt16BE(br);
                var encodingId = ReadUInt16BE(br);
                var languageId = ReadUInt16BE(br);
                var nameId = ReadUInt16BE(br);
                var length = ReadUInt16BE(br);
                var offset = ReadUInt16BE(br);
                records.Add((platformId, encodingId, languageId, nameId, length, offset));
            }

            var storagePos = nameOffset + stringOffset;
            foreach (var rec in records)
            {
                ms.Position = storagePos + rec.offset;
                var raw = br.ReadBytes(rec.length);
                string str;
                if (rec.platformId == 0 || rec.platformId == 3)
                {
                    str = System.Text.Encoding.BigEndianUnicode.GetString(raw);
                }
                else
                {
                    try { str = System.Text.Encoding.UTF8.GetString(raw); }
                    catch { str = System.Text.Encoding.Latin1.GetString(raw); }
                }

                if (!result.ContainsKey(rec.nameId)) result[rec.nameId] = str.Trim('\0').Trim();
            }
        }
        catch { }
        return result;
    }

    static ushort ReadUInt16BE(BinaryReader br)
    {
        var b1 = br.ReadByte();
        var b2 = br.ReadByte();
        return (ushort)((b1 << 8) | b2);
    }

    static uint ReadUInt32BE(BinaryReader br)
    {
        var b1 = br.ReadByte();
        var b2 = br.ReadByte();
        var b3 = br.ReadByte();
        var b4 = br.ReadByte();
        return (uint)((b1 << 24) | (b2 << 16) | (b3 << 8) | b4);
    }
}

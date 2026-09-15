using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter.Tests;

public class WordNumberingTests : IDisposable
{
    readonly MemoryStream _stream = new();
    readonly WordprocessingDocument _document;

    public WordNumberingTests()
    {
        _document = WordprocessingDocument.Create(_stream, WordprocessingDocumentType.Document);
        _document.AddMainDocumentPart().Document = new W.Document(new W.Body());
    }

    public void Dispose()
    {
        _document.Dispose();
        _stream.Dispose();
        GC.SuppressFinalize(this);
    }

    WordNumbering BuildNumbering(params W.AbstractNum[] definitions)
    {
        var part = _document.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>();
        var numbering = new W.Numbering();
        foreach (var definition in definitions) numbering.AppendChild(definition);
        foreach (var definition in definitions)
        {
            numbering.AppendChild(new W.NumberingInstance(
                new W.AbstractNumId { Val = definition.AbstractNumberId! })
            { NumberID = definition.AbstractNumberId!.Value + 100 });
        }
        part.Numbering = numbering;
        return new WordNumbering(WordStyleCache.For(_document.MainDocumentPart!));
    }

    static W.Level Level(int index, string format, string text, int? start = null)
    {
        var level = new W.Level
        {
            LevelIndex = index,
            NumberingFormat = new W.NumberingFormat { Val = new EnumValue<W.NumberFormatValues>(new W.NumberFormatValues(format)) },
            LevelText = new W.LevelText { Val = text },
        };
        if (start != null) level.StartNumberingValue = new W.StartNumberingValue { Val = start.Value };
        return level;
    }

    [Fact]
    public void NextLabel_BulletFormat_RendersBulletGlyphNotANumber()
    {
        var numbering = BuildNumbering(new W.AbstractNum(Level(0, "bullet", "")) { AbstractNumberId = 1 });

        Assert.Equal("• ", numbering.NextLabel("101", 0));
    }

    [Fact]
    public void NextLabel_DecimalFormat_SubstitutesPercentPlaceholder()
    {
        var numbering = BuildNumbering(new W.AbstractNum(Level(0, "decimal", "%1.")) { AbstractNumberId = 1 });

        Assert.Equal("1. ", numbering.NextLabel("101", 0));
        Assert.Equal("2. ", numbering.NextLabel("101", 0));
    }

    [Fact]
    public void NextLabel_MultiLevel_CombinesParentAndChildCounters()
    {
        var numbering = BuildNumbering(new W.AbstractNum(
            Level(0, "decimal", "%1."),
            Level(1, "decimal", "%1.%2.")) { AbstractNumberId = 1 });

        Assert.Equal("1. ", numbering.NextLabel("101", 0));
        Assert.Equal("1.1. ", numbering.NextLabel("101", 1));
        Assert.Equal("1.2. ", numbering.NextLabel("101", 1));
        Assert.Equal("2. ", numbering.NextLabel("101", 0));
        Assert.Equal("2.1. ", numbering.NextLabel("101", 1));
    }

    [Fact]
    public void NextLabel_StartValue_OffsetsTheFirstNumber()
    {
        var numbering = BuildNumbering(new W.AbstractNum(Level(0, "decimal", "%1.", start: 5)) { AbstractNumberId = 1 });

        Assert.Equal("5. ", numbering.NextLabel("101", 0));
        Assert.Equal("6. ", numbering.NextLabel("101", 0));
    }

    [Fact]
    public void NextLabel_LetterAndRomanFormats_AreRendered()
    {
        var numbering = BuildNumbering(
            new W.AbstractNum(Level(0, "lowerLetter", "%1)")) { AbstractNumberId = 1 },
            new W.AbstractNum(Level(0, "upperRoman", "%1.")) { AbstractNumberId = 2 });

        Assert.Equal("a) ", numbering.NextLabel("101", 0));
        Assert.Equal("b) ", numbering.NextLabel("101", 0));
        Assert.Equal("I. ", numbering.NextLabel("102", 0));
        Assert.Equal("II. ", numbering.NextLabel("102", 0));
    }

    [Fact]
    public void NextLabel_NoneFormat_ProducesNoLabel()
    {
        var numbering = BuildNumbering(new W.AbstractNum(Level(0, "none", "")) { AbstractNumberId = 1 });

        Assert.Null(numbering.NextLabel("101", 0));
    }

    [Fact]
    public void NextLabel_UnknownList_FallsBackToDecimal()
    {
        var numbering = BuildNumbering(new W.AbstractNum(Level(0, "decimal", "%1.")) { AbstractNumberId = 1 });

        Assert.Equal("1. ", numbering.NextLabel("999", 0));
    }
}

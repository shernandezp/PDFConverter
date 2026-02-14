using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.DocumentObjectModel;

namespace PDFConverter.Tests;

public class WordHelpersTests
{
    #region GetParagraphFormatting

    [Fact]
    public void GetParagraphFormatting_NullInput_ReturnsDefaults()
    {
        var fmt = WordHelpers.GetParagraphFormatting(null);

        Assert.Equal(ParagraphAlignment.Left, fmt.Alignment);
        Assert.Equal(0, fmt.LeftIndent);
        Assert.Equal(0, fmt.RightIndent);
        Assert.Equal(0, fmt.FirstLineIndent);
        Assert.Equal(0, fmt.SpacingBefore);
        Assert.Equal(0, fmt.SpacingAfter);
        Assert.Null(fmt.LineSpacing);
        Assert.Null(fmt.LineRule);
        Assert.False(fmt.HasExplicitSpacingBefore);
        Assert.False(fmt.HasExplicitSpacingAfter);
    }

    [Fact]
    public void GetParagraphFormatting_CenterAlignment()
    {
        var pPr = new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center });
        Assert.Equal(ParagraphAlignment.Center, WordHelpers.GetParagraphFormatting(pPr).Alignment);
    }

    [Fact]
    public void GetParagraphFormatting_RightAlignment()
    {
        var pPr = new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Right });
        Assert.Equal(ParagraphAlignment.Right, WordHelpers.GetParagraphFormatting(pPr).Alignment);
    }

    [Fact]
    public void GetParagraphFormatting_JustifyAlignment()
    {
        var pPr = new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Both });
        Assert.Equal(ParagraphAlignment.Justify, WordHelpers.GetParagraphFormatting(pPr).Alignment);
    }

    [Fact]
    public void GetParagraphFormatting_LeftAlignment()
    {
        var pPr = new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Left });
        Assert.Equal(ParagraphAlignment.Left, WordHelpers.GetParagraphFormatting(pPr).Alignment);
    }

    [Fact]
    public void GetParagraphFormatting_LeftIndent_ConvertsTwipsToPoints()
    {
        // 720 twips = 36 points
        var pPr = new W.ParagraphProperties(
            new W.Indentation { Left = "720" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.Equal(36.0, fmt.LeftIndent, 1);
    }

    [Fact]
    public void GetParagraphFormatting_RightIndent_ConvertsTwipsToPoints()
    {
        // 138 twips = 6.9 points
        var pPr = new W.ParagraphProperties(
            new W.Indentation { Right = "138" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.Equal(6.9, fmt.RightIndent, 1);
    }

    [Fact]
    public void GetParagraphFormatting_FirstLineIndent_ConvertsTwipsToPoints()
    {
        // 400 twips = 20 points
        var pPr = new W.ParagraphProperties(
            new W.Indentation { FirstLine = "400" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.Equal(20.0, fmt.FirstLineIndent, 1);
    }

    [Fact]
    public void GetParagraphFormatting_HangingIndent_ProducesNegativeFirstLine()
    {
        // Hanging indent of 422 twips = -21.1 points
        var pPr = new W.ParagraphProperties(
            new W.Indentation { Hanging = "422" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.Equal(-21.1, fmt.FirstLineIndent, 1);
    }

    [Fact]
    public void GetParagraphFormatting_HangingOverridesFirstLine()
    {
        // When both FirstLine and Hanging are set, Hanging wins (it's checked second)
        var pPr = new W.ParagraphProperties(
            new W.Indentation { FirstLine = "200", Hanging = "300" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        // Hanging = 300 twips → -15pt
        Assert.Equal(-15.0, fmt.FirstLineIndent, 1);
    }

    [Fact]
    public void GetParagraphFormatting_SpacingBefore_ConvertsTwipsToPoints()
    {
        // 240 twips = 12 points
        var pPr = new W.ParagraphProperties(
            new W.SpacingBetweenLines { Before = "240" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.Equal(12.0, fmt.SpacingBefore, 1);
        Assert.True(fmt.HasExplicitSpacingBefore);
    }

    [Fact]
    public void GetParagraphFormatting_SpacingAfter_ConvertsTwipsToPoints()
    {
        // 160 twips = 8 points
        var pPr = new W.ParagraphProperties(
            new W.SpacingBetweenLines { After = "160" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.Equal(8.0, fmt.SpacingAfter, 1);
        Assert.True(fmt.HasExplicitSpacingAfter);
    }

    [Fact]
    public void GetParagraphFormatting_SpacingNotSet_HasExplicitFlagsAreFalse()
    {
        var pPr = new W.ParagraphProperties();

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.False(fmt.HasExplicitSpacingBefore);
        Assert.False(fmt.HasExplicitSpacingAfter);
    }

    [Fact]
    public void GetParagraphFormatting_AutoLineSpacing_ConvertsTo240thsMultiple()
    {
        // Auto line spacing: 276/240 = 1.15 multiple
        var pPr = new W.ParagraphProperties(
            new W.SpacingBetweenLines { Line = "276" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.NotNull(fmt.LineSpacing);
        Assert.Equal(1.15, fmt.LineSpacing!.Value, 2);
        Assert.Equal("Auto", fmt.LineRule);
    }

    [Fact]
    public void GetParagraphFormatting_ExactLineSpacing_ConvertsTwipsToPoints()
    {
        // Exact: 240 twips / 20 = 12 points
        var pPr = new W.ParagraphProperties(
            new W.SpacingBetweenLines
            {
                Line = "240",
                LineRule = W.LineSpacingRuleValues.Exact
            });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.NotNull(fmt.LineSpacing);
        Assert.Equal(12.0, fmt.LineSpacing!.Value, 1);
        Assert.Equal("Exact", fmt.LineRule);
    }

    [Fact]
    public void GetParagraphFormatting_AtLeastLineSpacing_ConvertsTwipsToPoints()
    {
        var pPr = new W.ParagraphProperties(
            new W.SpacingBetweenLines
            {
                Line = "300",
                LineRule = W.LineSpacingRuleValues.AtLeast
            });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.NotNull(fmt.LineSpacing);
        Assert.Equal(15.0, fmt.LineSpacing!.Value, 1);
        Assert.Equal("AtLeast", fmt.LineRule);
    }

    [Fact]
    public void GetParagraphFormatting_CombinedProperties_AllMapped()
    {
        var pPr = new W.ParagraphProperties(
            new W.Justification { Val = W.JustificationValues.Center },
            new W.Indentation { Left = "720", Right = "360", FirstLine = "200" },
            new W.SpacingBetweenLines { Before = "120", After = "60" });

        var fmt = WordHelpers.GetParagraphFormatting(pPr);

        Assert.Equal(ParagraphAlignment.Center, fmt.Alignment);
        Assert.Equal(36.0, fmt.LeftIndent, 1);
        Assert.Equal(18.0, fmt.RightIndent, 1);
        Assert.Equal(10.0, fmt.FirstLineIndent, 1);
        Assert.Equal(6.0, fmt.SpacingBefore, 1);
        Assert.Equal(3.0, fmt.SpacingAfter, 1);
    }

    #endregion

    #region GetRunFontFamily

    [Fact]
    public void GetRunFontFamily_NullRunProperties_ReturnsNull()
    {
        Assert.Null(WordHelpers.GetRunFontFamily(null));
    }

    [Fact]
    public void GetRunFontFamily_AsciiFont_ReturnsAscii()
    {
        var rPr = new W.RunProperties(
            new W.RunFonts { Ascii = "Arial" });

        Assert.Equal("Arial", WordHelpers.GetRunFontFamily(rPr));
    }

    [Fact]
    public void GetRunFontFamily_NoAscii_FallsToHighAnsi()
    {
        var rPr = new W.RunProperties(
            new W.RunFonts { HighAnsi = "Verdana" });

        Assert.Equal("Verdana", WordHelpers.GetRunFontFamily(rPr));
    }

    [Fact]
    public void GetRunFontFamily_NoAsciiNoHighAnsi_FallsToComplexScript()
    {
        var rPr = new W.RunProperties(
            new W.RunFonts { ComplexScript = "Tahoma" });

        Assert.Equal("Tahoma", WordHelpers.GetRunFontFamily(rPr));
    }

    #endregion

    #region GetRunColor

    [Fact]
    public void GetRunColor_NullRunProperties_ReturnsNull()
    {
        Assert.Null(WordHelpers.GetRunColor(null));
    }

    [Fact]
    public void GetRunColor_WithColor_ReturnsValue()
    {
        var rPr = new W.RunProperties(
            new W.Color { Val = "FF0000" });

        Assert.Equal("FF0000", WordHelpers.GetRunColor(rPr));
    }

    #endregion

    #region GetTableGridColumnWidths

    [Fact]
    public void GetTableGridColumnWidths_FromTableGrid_ConvertsTwipsToPoints()
    {
        var table = new W.Table(
            new W.TableGrid(
                new W.GridColumn { Width = "1440" },  // 72pt
                new W.GridColumn { Width = "2880" })); // 144pt

        var widths = WordHelpers.GetTableGridColumnWidths(table);

        Assert.Equal(2, widths.Count);
        Assert.Equal(72.0, widths[0], 1);
        Assert.Equal(144.0, widths[1], 1);
    }

    [Fact]
    public void GetTableGridColumnWidths_NoGrid_FallsToCellWidths()
    {
        var table = new W.Table(
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(
                        new W.TableCellWidth { Width = "2000", Type = W.TableWidthUnitValues.Dxa }),
                    new W.Paragraph()),
                new W.TableCell(
                    new W.TableCellProperties(
                        new W.TableCellWidth { Width = "3000", Type = W.TableWidthUnitValues.Dxa }),
                    new W.Paragraph())));

        var widths = WordHelpers.GetTableGridColumnWidths(table);

        Assert.Equal(2, widths.Count);
        Assert.Equal(100.0, widths[0], 1);  // 2000/20
        Assert.Equal(150.0, widths[1], 1);  // 3000/20
    }

    [Fact]
    public void GetTableGridColumnWidths_CellWithGridSpan_SplitsEvenly()
    {
        var table = new W.Table(
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(
                        new W.TableCellWidth { Width = "4000", Type = W.TableWidthUnitValues.Dxa },
                        new W.GridSpan { Val = 2 }),
                    new W.Paragraph())));

        var widths = WordHelpers.GetTableGridColumnWidths(table);

        Assert.Equal(2, widths.Count);
        Assert.Equal(100.0, widths[0], 1);  // 4000/20/2 = 100
        Assert.Equal(100.0, widths[1], 1);
    }

    #endregion

    #region GetWordCellBorders

    [Fact]
    public void GetWordCellBorders_NullInput_ReturnsZeroBorders()
    {
        var borders = WordHelpers.GetWordCellBorders(null);

        Assert.Equal(0, borders.TopWidth);
        Assert.Equal(0, borders.BottomWidth);
        Assert.Equal(0, borders.LeftWidth);
        Assert.Equal(0, borders.RightWidth);
    }

    [Fact]
    public void GetWordCellBorders_WithCellBorders_ParsesSizeInEighths()
    {
        // sz=8 → 8/8 = 1pt
        var tcPr = new W.TableCellProperties(
            new W.TableCellBorders(
                new W.TopBorder { Val = W.BorderValues.Single, Size = 8, Color = "000000" },
                new W.BottomBorder { Val = W.BorderValues.Single, Size = 16, Color = "FF0000" }));

        var borders = WordHelpers.GetWordCellBorders(tcPr);

        Assert.Equal(1.0, borders.TopWidth, 2);
        Assert.Equal("#000000", borders.TopColor);
        Assert.Equal(2.0, borders.BottomWidth, 2);
        Assert.Equal("#FF0000", borders.BottomColor);
    }

    [Fact]
    public void GetWordCellBorders_TableBordersAsDefaults_CellOverrides()
    {
        var tblPr = new W.TableProperties(
            new W.TableBorders(
                new W.TopBorder { Val = W.BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new W.BottomBorder { Val = W.BorderValues.Single, Size = 4, Color = "AAAAAA" }));

        var tcPr = new W.TableCellProperties(
            new W.TableCellBorders(
                new W.TopBorder { Val = W.BorderValues.Single, Size = 12, Color = "0000FF" }));

        var borders = WordHelpers.GetWordCellBorders(tcPr, tblPr);

        // Top overridden by cell border
        Assert.Equal(1.5, borders.TopWidth, 2);   // 12/8
        Assert.Equal("#0000FF", borders.TopColor);
        // Bottom from table defaults
        Assert.Equal(0.5, borders.BottomWidth, 2); // 4/8
        Assert.Equal("#AAAAAA", borders.BottomColor);
    }

    [Fact]
    public void GetWordCellBorders_InsideH_AppliesAsDefaultTopBottom()
    {
        var tblPr = new W.TableProperties(
            new W.TableBorders(
                new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4, Color = "999999" }));

        var borders = WordHelpers.GetWordCellBorders(null, tblPr);

        Assert.Equal(0.5, borders.TopWidth, 2);    // insideH applies to top when top=0
        Assert.Equal(0.5, borders.BottomWidth, 2);
        Assert.Equal("#999999", borders.TopColor);
    }

    [Fact]
    public void GetWordCellBorders_InsideV_AppliesAsDefaultLeftRight()
    {
        var tblPr = new W.TableProperties(
            new W.TableBorders(
                new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 8, Color = "333333" }));

        var borders = WordHelpers.GetWordCellBorders(null, tblPr);

        Assert.Equal(1.0, borders.LeftWidth, 2);
        Assert.Equal(1.0, borders.RightWidth, 2);
        Assert.Equal("#333333", borders.LeftColor);
    }

    [Fact]
    public void GetWordCellBorders_CellMargins_ParsesCorrectly()
    {
        var tcPr = new W.TableCellProperties(
            new W.TableCellMargin(
                new W.TopMargin { Width = "60", Type = W.TableWidthUnitValues.Dxa },
                new W.BottomMargin { Width = "40", Type = W.TableWidthUnitValues.Dxa },
                new W.LeftMargin { Width = "100", Type = W.TableWidthUnitValues.Dxa },
                new W.RightMargin { Width = "80", Type = W.TableWidthUnitValues.Dxa }));

        var borders = WordHelpers.GetWordCellBorders(tcPr);

        Assert.Equal(3.0, borders.PaddingTop, 1);     // 60/20
        Assert.Equal(2.0, borders.PaddingBottom, 1);   // 40/20
        Assert.Equal(5.0, borders.PaddingLeft, 1);     // 100/20
        Assert.Equal(4.0, borders.PaddingRight, 1);    // 80/20
    }

    #endregion

    #region GetStyleParagraphProperties

    [Fact]
    public void GetStyleParagraphProperties_ReturnsNonNull_WhenStyleHasParagraphProps()
    {
        // This validates the critical fix: StyleParagraphProperties → ParagraphProperties conversion
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.Style(
                new W.StyleParagraphProperties(
                    new W.Justification { Val = W.JustificationValues.Center }))
            {
                Type = W.StyleValues.Paragraph,
                StyleId = "TestStyle"
            });

        var result = WordHelpers.GetStyleParagraphProperties(mainPart, "TestStyle");

        Assert.NotNull(result);
        Assert.NotNull(result!.Justification);
        Assert.Equal(W.JustificationValues.Center, result.Justification!.Val!.Value);
    }

    [Fact]
    public void GetStyleParagraphProperties_ReturnsNull_WhenStyleNotFound()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles();

        var result = WordHelpers.GetStyleParagraphProperties(mainPart, "NonExistent");

        Assert.Null(result);
    }

    [Fact]
    public void GetStyleParagraphProperties_WalksBasedOnChain()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            // Child style — no paragraph properties
            new W.Style(
                new W.BasedOn { Val = "ParentStyle" })
            {
                Type = W.StyleValues.Paragraph,
                StyleId = "ChildStyle"
            },
            // Parent style — has paragraph properties
            new W.Style(
                new W.StyleParagraphProperties(
                    new W.Justification { Val = W.JustificationValues.Right },
                    new W.Indentation { Left = "480" }))
            {
                Type = W.StyleValues.Paragraph,
                StyleId = "ParentStyle"
            });

        var result = WordHelpers.GetStyleParagraphProperties(mainPart, "ChildStyle");

        Assert.NotNull(result);
        Assert.Equal(W.JustificationValues.Right, result!.Justification!.Val!.Value);
        Assert.Equal("480", result.Indentation!.Left!.Value);
    }

    [Fact]
    public void GetStyleParagraphProperties_ReturnType_IsParagraphProperties_NotStyleParagraphProperties()
    {
        // This is THE critical test — the original bug was that CloneNode(true) as ParagraphProperties
        // returned null because StyleParagraphProperties is a different type
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.Style(
                new W.StyleParagraphProperties(
                    new W.SpacingBetweenLines { Before = "100", After = "200" }))
            {
                Type = W.StyleValues.Paragraph,
                StyleId = "SpacedStyle"
            });

        var result = WordHelpers.GetStyleParagraphProperties(mainPart, "SpacedStyle");

        // Must return a real ParagraphProperties, not null
        Assert.NotNull(result);
        Assert.IsType<W.ParagraphProperties>(result);

        // Verify children were properly transferred
        var spacing = result!.SpacingBetweenLines;
        Assert.NotNull(spacing);
        Assert.Equal("100", spacing!.Before!.Value);
        Assert.Equal("200", spacing.After!.Value);
    }

    #endregion

    #region GetStyleRunProperties

    [Fact]
    public void GetStyleRunProperties_ReturnsProperties_WhenStyleExists()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.Style(
                new W.StyleRunProperties(
                    new W.Bold(),
                    new W.FontSize { Val = "24" },
                    new W.Color { Val = "112233" }))
            {
                Type = W.StyleValues.Paragraph,
                StyleId = "BoldStyle"
            });

        var result = WordHelpers.GetStyleRunProperties(mainPart, "BoldStyle");

        Assert.NotNull(result);
        Assert.NotNull(result!.Bold);
        Assert.Equal("24", result.FontSize!.Val!.Value);
        Assert.Equal("112233", result.Color!.Val!.Value);
    }

    #endregion

    #region ResolveRunFormatting

    [Fact]
    public void ResolveRunFormatting_InlineProperties_TakePrecedence()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());
        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles();

        var para = new W.Paragraph();
        var run = new W.Run(
            new W.RunProperties(
                new W.Bold(),
                new W.Italic(),
                new W.Underline { Val = W.UnderlineValues.Single },
                new W.FontSize { Val = "28" },    // 14pt
                new W.Color { Val = "FF0000" },
                new W.RunFonts { Ascii = "Courier" }),
            new W.Text("test"));

        var fmt = WordHelpers.ResolveRunFormatting(mainPart, run, para);

        Assert.True(fmt.Bold);
        Assert.True(fmt.Italic);
        Assert.True(fmt.Underline);
        Assert.Equal(14.0, fmt.Size);
        Assert.Equal("FF0000", fmt.Color);
        Assert.Equal("Courier", fmt.FontFamily);
    }

    [Fact]
    public void ResolveRunFormatting_FallsBackToStyle()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.Style(
                new W.StyleRunProperties(
                    new W.RunFonts { Ascii = "Times" },
                    new W.FontSize { Val = "20" }))   // 10pt
            {
                Type = W.StyleValues.Paragraph,
                StyleId = "Normal"
            });

        var para = new W.Paragraph(
            new W.ParagraphProperties(
                new W.ParagraphStyleId { Val = "Normal" }));
        var run = new W.Run(new W.Text("test"));

        var fmt = WordHelpers.ResolveRunFormatting(mainPart, run, para);

        Assert.Equal("Times", fmt.FontFamily);
        Assert.Equal(10.0, fmt.Size);
    }

    [Fact]
    public void ResolveRunFormatting_FallsBackToDocDefaults()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.DocDefaults(
                new W.RunPropertiesDefault(
                    new W.RunProperties(
                        new W.RunFonts { Ascii = "Calibri" },
                        new W.FontSize { Val = "22" }))));   // 11pt

        var para = new W.Paragraph();
        var run = new W.Run(new W.Text("test"));

        var fmt = WordHelpers.ResolveRunFormatting(mainPart, run, para);

        Assert.Equal("Calibri", fmt.FontFamily);
        Assert.Equal(11.0, fmt.Size);
    }

    #endregion

    #region ResolveTableBorders

    [Fact]
    public void ResolveTableBorders_InlineBordersOverrideStyle()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.Style(
                new W.StyleTableProperties(
                    new W.TableBorders(
                        new W.TopBorder { Val = W.BorderValues.Single, Size = 4 })))
            {
                Type = W.StyleValues.Table,
                StyleId = "TableGrid"
            });

        var tblPr = new W.TableProperties(
            new W.TableStyle { Val = "TableGrid" },
            new W.TableBorders(
                new W.TopBorder { Val = W.BorderValues.Double, Size = 12 }));

        var result = WordHelpers.ResolveTableBorders(mainPart, tblPr);

        Assert.NotNull(result);
        // Inline should override style
        Assert.Equal(W.BorderValues.Double, result!.TopBorder!.Val!.Value);
    }

    [Fact]
    public void ResolveTableBorders_FallsBackToStyle_WhenNoInline()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.Style(
                new W.StyleTableProperties(
                    new W.TableBorders(
                        new W.BottomBorder { Val = W.BorderValues.Dotted, Size = 8 })))
            {
                Type = W.StyleValues.Table,
                StyleId = "DottedTable"
            });

        var tblPr = new W.TableProperties(
            new W.TableStyle { Val = "DottedTable" });

        var result = WordHelpers.ResolveTableBorders(mainPart, tblPr);

        Assert.NotNull(result);
        Assert.Equal(W.BorderValues.Dotted, result!.BottomBorder!.Val!.Value);
    }

    #endregion

    #region GetDocDefaults

    [Fact]
    public void GetDocDefaultsParagraphProperties_ReturnsProperties()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.DocDefaults(
                new W.ParagraphPropertiesDefault(
                    new W.ParagraphProperties(
                        new W.SpacingBetweenLines { After = "160" }))));

        var result = WordHelpers.GetDocDefaultsParagraphProperties(mainPart);

        Assert.NotNull(result);
        Assert.Equal("160", result!.SpacingBetweenLines!.After!.Value);
    }

    [Fact]
    public void GetDocDefaultsRunProperties_ReturnsProperties()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(
            new W.DocDefaults(
                new W.RunPropertiesDefault(
                    new W.RunProperties(
                        new W.FontSize { Val = "22" }))));

        var result = WordHelpers.GetDocDefaultsRunProperties(mainPart);

        Assert.NotNull(result);
        Assert.Equal("22", result!.FontSize!.Val!.Value);
    }

    [Fact]
    public void GetDocDefaultsParagraphProperties_ReturnsNull_WhenNoDefaults()
    {
        using var ms = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles();

        Assert.Null(WordHelpers.GetDocDefaultsParagraphProperties(mainPart));
    }

    #endregion

    #region IsConditionalRow / IsConditionalColumn

    [Fact]
    public void IsConditionalRow_FirstRow_WithTblLook_ReturnsTrue()
    {
        var tblPr = new W.TableProperties(
            new W.TableLook { FirstRow = true });

        var row = new W.TableRow();

        Assert.True(WordHelpers.IsConditionalRow(row, tblPr, 0, 5,
            W.TableStyleOverrideValues.FirstRow));
    }

    [Fact]
    public void IsConditionalRow_FirstRow_MiddlePosition_ReturnsFalse()
    {
        var tblPr = new W.TableProperties(
            new W.TableLook { FirstRow = true });

        var row = new W.TableRow();

        Assert.False(WordHelpers.IsConditionalRow(row, tblPr, 2, 5,
            W.TableStyleOverrideValues.FirstRow));
    }

    [Fact]
    public void IsConditionalRow_LastRow_WithTblLook_ReturnsTrue()
    {
        var tblPr = new W.TableProperties(
            new W.TableLook { LastRow = true });

        var row = new W.TableRow();

        Assert.True(WordHelpers.IsConditionalRow(row, tblPr, 4, 5,
            W.TableStyleOverrideValues.LastRow));
    }

    [Fact]
    public void IsConditionalRow_CnfStyleFirstRow_ReturnsTrue()
    {
        var row = new W.TableRow(
            new W.TableRowProperties(
                new W.ConditionalFormatStyle { FirstRow = true }));

        Assert.True(WordHelpers.IsConditionalRow(row, null, 0, 5,
            W.TableStyleOverrideValues.FirstRow));
    }

    [Fact]
    public void IsConditionalColumn_FirstColumn_ReturnsTrue()
    {
        var tblPr = new W.TableProperties(
            new W.TableLook { FirstColumn = true });

        var cell = new W.TableCell(new W.Paragraph());

        Assert.True(WordHelpers.IsConditionalColumn(cell, tblPr, 0, 5));
    }

    [Fact]
    public void IsConditionalColumn_NonFirstColumn_ReturnsFalse()
    {
        var tblPr = new W.TableProperties(
            new W.TableLook { FirstColumn = true });

        var cell = new W.TableCell(new W.Paragraph());

        Assert.False(WordHelpers.IsConditionalColumn(cell, tblPr, 2, 5));
    }

    #endregion
}

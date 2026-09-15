using System.Globalization;
using Xunit;

namespace PDFConverter.Tests;

public class ExcelNumberFormatTests
{
    static string Apply(string raw, string format)
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            return ExcelNumberFormat.Apply(raw, format);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void Apply_TimeFormat_TreatsMAsMinutesNotMonths()
    {
        Assert.Equal("14:05:09", Apply("0.58690972222", "h:mm:ss"));
    }

    [Fact]
    public void Apply_DateFormat_TreatsMAsMonths()
    {
        Assert.Equal("25/12/2020", Apply("44190", "dd/mm/yyyy"));
    }

    [Fact]
    public void Apply_DateTimeFormat_DistinguishesMonthsFromMinutes()
    {
        Assert.Equal("25/12/2020 14:05", Apply("44190.58690972222", "dd/mm/yyyy hh:mm"));
    }

    [Fact]
    public void Apply_CurrencyWithColorSection_UsesPositiveSectionAndStripsColor()
    {
        Assert.Equal("$90.00", Apply("90", "$#,##0.00;[Red]-$#,##0.00"));
    }

    [Fact]
    public void Apply_NegativeValue_UsesNegativeSection()
    {
        Assert.Equal("(90.00)", Apply("-90", "#,##0.00;(#,##0.00)"));
    }

    [Fact]
    public void Apply_ZeroSection_IsUsedForZero()
    {
        Assert.Equal("-", Apply("0", "#,##0.00;(#,##0.00);\"-\""));
    }

    [Fact]
    public void Apply_QuotedLiteralsAndPaddingTokens_AreHandled()
    {
        Assert.Equal("1,234 kg", Apply("1234", "_(#,##0\" kg\"_)"));
    }

    [Fact]
    public void Apply_Percentage_ScalesTheValue()
    {
        Assert.Equal("12.5%", Apply("0.125", "0.0%"));
    }

    [Fact]
    public void Apply_UnparseableValue_IsReturnedUnchanged()
    {
        Assert.Equal("N/A", Apply("N/A", "#,##0.00"));
    }

    [Fact]
    public void Apply_NoFormat_ReturnsRawValue()
    {
        Assert.Equal("1234.5", ExcelNumberFormat.Apply("1234.5", null));
    }

    [Fact]
    public void Apply_ParsesValueInvariantly_RegardlessOfThreadCulture()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("es-ES");
            Assert.Equal("1.234,50", ExcelNumberFormat.Apply("1234.5", "#,##0.00"));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }
}

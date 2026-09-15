using System.Globalization;
using Xunit;

namespace PDFConverter.Tests;

public class UnitsTests
{
    [Theory]
    [InlineData("1440", 72.0)]
    [InlineData("240", 12.0)]
    [InlineData("-120", -6.0)]
    public void TwipsToPoints_ConvertsTwentiethsOfAPoint(string twips, double expected)
    {
        Assert.Equal(expected, Units.TwipsToPoints(twips)!.Value, 3);
    }

    [Fact]
    public void EmuToPoints_ConvertsEnglishMetricUnits()
    {
        Assert.Equal(72.0, Units.EmuToPoints(914400), 3);
    }

    [Fact]
    public void HalfPointsToPoints_HalvesTheValue()
    {
        Assert.Equal(11.0, Units.HalfPointsToPoints("22")!.Value, 3);
    }

    [Fact]
    public void EighthPointsToPoints_ConvertsBorderSizes()
    {
        Assert.Equal(1.0, Units.EighthPointsToPoints(8), 3);
    }

    [Fact]
    public void TryParseDouble_UsesInvariantCulture_RegardlessOfThreadCulture()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("es-ES");
            Assert.True(Units.TryParseDouble("1234.5", out var value));
            Assert.Equal(1234.5, value, 3);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void TryParseDouble_RejectsMalformedInput()
    {
        Assert.False(Units.TryParseDouble("auto", out _));
        Assert.False(Units.TryParseDouble(null, out _));
    }
}

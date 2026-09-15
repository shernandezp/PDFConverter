using Xunit;

namespace PDFConverter.Tests;

public class ColorUtilsTests
{
    [Theory]
    [InlineData("FF0000")]
    [InlineData("#FF0000")]
    [InlineData("00FF0000")]
    public void TryParse_AcceptsOpenXmlColorSpellings(string value)
    {
        Assert.True(ColorUtils.TryParse(value, out var color));
        Assert.Equal(MigraDoc.DocumentObjectModel.Color.Parse("#FF0000"), color);
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("AUTO")]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("ZZZZZZ")]
    [InlineData("12345")]
    public void TryParse_RejectsSentinelsAndMalformedValues(string? value)
    {
        Assert.False(ColorUtils.TryParse(value, out _));
    }
}

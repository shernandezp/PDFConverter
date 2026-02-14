using PDFConverter;


Console.WriteLine("PDFConverter Test Console");

if (args.Length < 2)
{
    Console.WriteLine("Usage: TestConsole <docx|xlsx> <input> [output.pdf]");
    return;
}

var type = args[0].ToLowerInvariant();
var input = args[1];
var output = args.Length >= 3 ? args[2] : Path.ChangeExtension(input, ".pdf");

try
{
    if (type == "docx")
    {
        if (File.Exists(input))
        {
            Converters.DocxToPdf(input, output);
        }
        else
        {
            var bytes = Convert.FromBase64String(File.ReadAllText(input));
            Converters.DocxToPdf(bytes, output);
        }
    }
    else if (type == "xlsx")
    {
        if (File.Exists(input))
        {
            Converters.XlsxToPdf(input, output);
        }
        else
        {
            var bytes = Convert.FromBase64String(File.ReadAllText(input));
            Converters.XlsxToPdf(bytes, output);
        }
    }
    else
    {
        Console.WriteLine("Unknown type. Use 'docx' or 'xlsx'.");
    }

    Console.WriteLine($"Saved PDF: {output}");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}

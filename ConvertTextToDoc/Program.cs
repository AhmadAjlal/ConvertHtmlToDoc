using HtmlAgilityPack;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main()
    {
        Console.WriteLine("Enter the directory path to scan for files:");
        string? directoryPath = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
        {
            Console.WriteLine("Directory does not exist or input was invalid.");
            return;
        }

        Console.WriteLine("Enter the Extension of the file you want to read (No need to add \".\"):");
        string? extension = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(extension))
        {
            Console.WriteLine("Extension input is invalid.");
            return;
        }

        string outputDocxPath = Path.Combine(Directory.GetCurrentDirectory(), "CleanCshtmlText.docx");

        var cshtmlFiles = Directory.GetFiles(directoryPath, $"*.{extension}", SearchOption.AllDirectories);

        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputDocxPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            Body body = mainPart.Document.Body ?? throw new InvalidOperationException("Body is null.");

            foreach (var file in cshtmlFiles)
            {
                //if (file.Contains("Blog", StringComparison.OrdinalIgnoreCase))
                //continue;

                string rawContent = File.ReadAllText(file);
                string cleanText = ExtractTextFromHtml(rawContent);

                if (!string.IsNullOrWhiteSpace(cleanText))
                {
                    string fileName = Path.GetFileNameWithoutExtension(file) ?? "UnknownFile";

                    body.Append(new Paragraph(new Run(new Text($"File: {fileName}")))
                    {
                        ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Center })
                    });

                    body.Append(new Paragraph(new Run(new Text(cleanText.Trim()))));

                    body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }
            }

            mainPart.Document.Save();
        }

        Console.WriteLine($"Clean document created at: {outputDocxPath}");
    }

    // Strip HTML and Razor syntax using HtmlAgilityPack
    static string ExtractTextFromHtml(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
            return string.Empty;

        // Remove Razor code blocks: @{ ... }
        html = System.Text.RegularExpressions.Regex.Replace(html, @"@\{[\s\S]*?\}", "", System.Text.RegularExpressions.RegexOptions.Multiline);

        // Remove Razor inline expressions: @something
        html = System.Text.RegularExpressions.Regex.Replace(html, @"@\w+(\([^\)]*\))?", "", System.Text.RegularExpressions.RegexOptions.Multiline);

        // Remove common Razor keywords like ViewData["Title"], Layout = "...";
        html = System.Text.RegularExpressions.Regex.Replace(html, @"ViewData\[[^\]]+\]\s*=\s*""[^""]*"";", "", System.Text.RegularExpressions.RegexOptions.Multiline);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"Layout\s*=\s*""[^""]*"";", "", System.Text.RegularExpressions.RegexOptions.Multiline);

        // Remove <script> and <style> blocks
        html = System.Text.RegularExpressions.Regex.Replace(html, @"<script[\s\S]*?</script>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"<style[\s\S]*?</style>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        string? text = doc.DocumentNode?.InnerText;
        if (text == null)
            return string.Empty;

        text = System.Text.RegularExpressions.Regex.Replace(text, @"\s{2,}", " ");
        text = System.Text.RegularExpressions.Regex.Replace(text, @"\n+", "\n");

        return text.Trim();
    }
}

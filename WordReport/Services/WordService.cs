using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordReport.ViewModels;

namespace WordReport.Services;

public class WordService
{
    public byte[] GenerateDocument(DocumentViewModel model, string templatePath)
    {
        byte[] byteArray = File.ReadAllBytes(templatePath);

        using (MemoryStream mem = new MemoryStream())
        {
            mem.Write(byteArray, 0, byteArray.Length);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
            {
                var placeholders = new Dictionary<string, string>
                {
                    { "{Title}", model.Title },
                    { "{Date}", model.Date }
                };

                // Replace in body
                ReplacePlaceholders(wordDoc.MainDocumentPart.Document.Body, placeholders);

                // Replace in headers
                foreach (var header in wordDoc.MainDocumentPart.HeaderParts)
                {
                    ReplacePlaceholders(header.Header, placeholders);
                }

                // Replace in footers
                foreach (var footer in wordDoc.MainDocumentPart.FooterParts)
                {
                    ReplacePlaceholders(footer.Footer, placeholders);
                }

                // Add rows to the table
                AddTableRows(wordDoc.MainDocumentPart.Document.Body, model.Items);

                wordDoc.MainDocumentPart.Document.Save();

                wordDoc.MainDocumentPart.Document.Save();
            }

            return mem.ToArray();
        }
    }

    private void ReplacePlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        foreach (var text in element.Descendants<Text>())
        {
            foreach (var key in placeholders.Keys)
            {
                if (text.Text.Contains(key))
                {
                    text.Text = text.Text.Replace(key, placeholders[key]);
                }
            }
        }
    }

    private void AddTableRows(Body body, List<TableItemViewModel> items)
    {
        var table = body.Descendants<Table>().FirstOrDefault();
        if (table == null || items == null || !items.Any()) return;

        // Skip the header row
        var headerRow = table.Elements<TableRow>().FirstOrDefault();

        foreach (var item in items)
        {
            var newRow = new TableRow();

            newRow.Append(CreateCell(item.Name));
            newRow.Append(CreateCell(item.Quantity.ToString()));

            table.Append(newRow);
        }
    }

    // Helper to create table cell
    private TableCell CreateCell(string text)
    {
        return new TableCell(
            new Paragraph(
                new Run(
                    new Text(text)
                )
            )
        );
    }
}

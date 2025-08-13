using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using WordReport.Models;
using WordReport.ViewModels;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using GemBox.Document;

namespace WordReport.Services;

public class CleanWordService
{
    public byte[] GenerateDocument(DocumentViewModel model, string templatePath)
    {
        var products = DataService.GetProducts();
        var tests = DataService.GetTests();
        var quotations = DataService.GetQuotations();

        byte[] byteArray = File.ReadAllBytes(templatePath);

        using (MemoryStream mem = new MemoryStream())
        {
            mem.Write(byteArray, 0, byteArray.Length);
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
            {
                var placeholders = new Dictionary<string, string>
                {
                    { "{{For}}", "The Great User" },
                    { "{{Subject}}", "Important Subject" },
                    { "{{ProposalReference}}", "No Proposal Reference" },
                    { "{{ProposalDate}}", DateTime.Now.ToString("dd MMMM yyyy") },
                    { "{{ValidFor}}", "7 Days" }
                };

                var bodyPlaceholders = new Dictionary<string, string>
                {
                    { "{{For}}", "The Great User" },
                    { "{{Company}}", "SomeCompany" },
                    { "{{Location}}", "Some Location" },
                    { "{{RoutineAnalysis}}", "80 Days" },
                    { "{{SubcontractedParameters}}", "70 Days" }
                };

                ReplacePlaceholdersAdvanced(wordDoc.MainDocumentPart.Document.Body, placeholders);
                //ReplacePlaceholdersAdvanced();
                ReplacePlaceholdersInitialTable(wordDoc.MainDocumentPart.Document.Body, placeholders);
                AddQuotationTableAfterTitle(wordDoc.MainDocumentPart.Document.Body, quotations);
                AddTableRowsByIndex(wordDoc.MainDocumentPart.Document.Body, tests, 4);

                wordDoc.MainDocumentPart.Document.Save();

                //  Generate XML file from the document
                var xmlPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(templatePath), "document.xml");
                var docPart = wordDoc.MainDocumentPart;

                using (var reader = new StreamReader(docPart.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    var xmlContent = reader.ReadToEnd();
                    File.WriteAllText(xmlPath, xmlContent);
                }
            }

            return mem.ToArray();
        }
    }

    public static void ReplacePlaceholdersInitialTable(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        if (element == null || placeholders == null || placeholders.Count == 0)
            return;

        var firstTable = element.Elements<W.Table>().FirstOrDefault();

        if (firstTable == null)
            return;

        foreach (var para in firstTable.Descendants<W.Paragraph>())
        {
            var texts = para.Descendants<W.Text>().ToList();
            if (!texts.Any()) continue;

            // Combine all runs into a single string
            string combinedText = string.Concat(texts.Select(t => t.Text));

            // Replace all placeholders in the combined string
            foreach (var kvp in placeholders)
            {
                combinedText = combinedText.Replace(kvp.Key, kvp.Value ?? string.Empty);
            }

            // Put the modified text back into the first <w:t> and clear the rest
            texts.First().Text = combinedText;
            for (int i = 1; i < texts.Count; i++)
            {
                texts[i].Text = string.Empty;
            }
        }
    }


    private void AddTableRowsByIndex<T>(W.Body body, List<T> items, int tableIndex)
    {
        if (items == null || items.Count == 0) return;

        var table = body.Descendants<W.Table>().ElementAtOrDefault(tableIndex);

        if (table == null)
            throw new InvalidOperationException($"No table found at index {tableIndex}");

        var rows = table.Elements<W.TableRow>().ToList();

        if (rows.Count < 2)
            throw new InvalidOperationException($"Table {tableIndex} does not have enough rows for label/property rows.");

        PopulateTableWithItems(table, items);
    }

    private void PopulateTableWithItems<T>(W.Table table, List<T> items)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count < 2) return;


        // The Opposite => Because no Keys this time in the table 
        var labelRow = rows[1]; // First row (display names)
        var propertyRow = rows[0]; // Second row (property names)

        // Read property names from table
        var propertyNames = propertyRow.Elements<W.TableCell>()
            .Select(cell => cell?.InnerText?.Trim() ?? "")
            .ToList();

        // Map each property name to its PropertyInfo (or null if not found)
        var props = propertyNames
            .Select(name => typeof(T).GetProperties()
                .FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        // Remove old data rows
        for (int i = 1; i < rows.Count; i++)
        {
            rows[i].Remove();
        }

        int index = 1;
        foreach (var item in items)
        {
            var row = new W.TableRow();

            // First column: index
            row.Append(CreateCell(index.ToString()));

            // Remaining columns: property values or empty cells
            foreach (var prop in props.Skip(1)) // Skip first property name (index column)
            {
                string value = string.Empty;
                if (prop != null)
                {
                    var propValue = prop.GetValue(item);
                    value = propValue?.ToString() ?? "";
                }
                row.Append(CreateCell(value));
            }

            table.Append(row);
            index++;
        }
    }

    private W.TableCell CreateCell(string text)
    {
        return new W.TableCell(
            new W.Paragraph(
                new W.Run(
                    new W.Text(text ?? "")
                )
            )
        );
    }

    private W.TableCell CreateCenteredCell(string text)
    {
        return new W.TableCell(
            new W.Paragraph(
                new W.ParagraphProperties(
                    new W.Justification { Val = W.JustificationValues.Center } // Horizontal center
                ),
                new W.Run(
                    new W.Text(text ?? "")
                )
            ),
            new W.TableCellProperties(
                new W.TableCellVerticalAlignment { Val = W.TableVerticalAlignmentValues.Center } // Vertical center
            )
        );
    }

    private void AddQuotationTableAfterTitle(W.Body body, Quotation quotation)
    {
        // Find the Title Quotation
        var titleParagraph = body.Descendants<W.Paragraph>()
        .FirstOrDefault(p => p.InnerText.Trim().Equals("Quotation", StringComparison.OrdinalIgnoreCase) && p.Ancestors<W.Table>().Count() == 0);

        if (titleParagraph == null)
            throw new InvalidOperationException("No 'Quotation' title found in the document.");

        // Finding The First Table After The Title Quotation
        var table = titleParagraph.ElementsAfter().OfType<W.Table>().FirstOrDefault();
        if (table == null)
            throw new InvalidOperationException("No table found after 'Quotation' title.");

        ApplyTableBorders(table);

        // Remove All The Rows
        var rows = table.Elements<W.TableRow>().ToList();
        for (int i = 1; i < rows.Count; i++)
            rows[i].Remove();

        int index = 1;
        decimal subtotal = 0;


        // Normal Cells
        foreach (var item in quotation.Items)
        {
            decimal unitCost = item.UnitCost ?? 0;
            int qty = item.Quantity ?? 0;
            decimal total = item.TotalCost ?? (unitCost * qty);

            subtotal += total;

            var row = new W.TableRow();
            row.Append(CreateCenteredCell(index.ToString("00")));       
            row.Append(CreateCenteredCell(item.Deliverable ?? ""));     
            row.Append(CreateCenteredCell(item.Unit ?? ""));            
            row.Append(CreateCenteredCell(unitCost.ToString("N2")));    
            row.Append(CreateCenteredCell(qty.ToString()));             
            row.Append(CreateCenteredCell(total.ToString("N2")));       
            table.Append(row);
            index++;
        }


        //  Spanned Cells
        table.Append(CreateSummaryRow("Subtotal (without VAT)", subtotal, 5));
        table.Append(CreateSummaryRow("VAT", quotation.Vat ?? 0, 5));
        table.Append(CreateSummaryRow("Total", quotation.FinalTotal ?? 0, 5, bold: true));
    }


    private W.TableRow CreateSummaryRow(string label, decimal amount, int mergeColumns, bool bold = false)
    {
        var row = new W.TableRow();

        // Merged label cell
        row.Append(CreateCell(label, bold, mergeColumns));

        // Last column (Total Cost)
        row.Append(CreateCell(amount.ToString("N2"), bold));

        return row;
    }

    private W.TableCell CreateCell(string text, bool bold = false, int colspan = 1)
    {
        // Create run
        var run = new W.Run(new W.Text(text ?? ""));
        if (bold)
            run.RunProperties = new W.RunProperties(new W.Bold());

        // Paragraph with center alignment
        var para = new W.Paragraph(run)
        {
            ParagraphProperties = new W.ParagraphProperties
            {
                Justification = new W.Justification { Val = W.JustificationValues.Center }
            }
        };

        // Cell properties (center vertically + optional colspan)
        var cellProps = new W.TableCellProperties(
            new W.TableCellVerticalAlignment { Val = W.TableVerticalAlignmentValues.Center }
        );

        if (colspan > 1)
            cellProps.Append(new W.GridSpan { Val = colspan });

        var cell = new W.TableCell(para);
        cell.Append(cellProps);

        return cell;
    }

    private void ApplyTableBorders(W.Table table)
    {
        var tblBorders = new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Single, Size = 6 },
            new W.BottomBorder { Val = W.BorderValues.Single, Size = 6 },
            new W.LeftBorder { Val = W.BorderValues.Single, Size = 6 },
            new W.RightBorder { Val = W.BorderValues.Single, Size = 6 },
            new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 6 },
            new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 6 }
        );

        var tblProps = table.GetFirstChild<W.TableProperties>();
        if (tblProps == null)
        {
            tblProps = new W.TableProperties();
            table.PrependChild(tblProps);
        }

        tblProps.TableBorders = tblBorders;
    }

    private void ReplacePlaceholdersInParagraph(W.Paragraph paragraph, Dictionary<string, string> placeholders)
    {
        // Get all text content from the paragraph
        string paragraphText = GetParagraphText(paragraph);

        // Check if any placeholder exists in the paragraph
        bool hasReplacements = false;
        string modifiedText = paragraphText;

        foreach (var kvp in placeholders)
        {
            if (modifiedText.Contains(kvp.Key))
            {
                modifiedText = modifiedText.Replace(kvp.Key, kvp.Value);
                hasReplacements = true;
            }
        }

        // If no replacements needed, return early
        if (!hasReplacements)
            return;

        // Get the first run's formatting to preserve style
        var firstRun = paragraph.Descendants<W.Run>().FirstOrDefault();
        var runProperties = firstRun?.RunProperties?.CloneNode(true) as W.RunProperties;

        // Clear all existing runs
        paragraph.RemoveAllChildren<W.Run>();

        // Create a new run with the modified text and original formatting
        var newRun = new W.Run();
        if (runProperties != null)
            newRun.RunProperties = runProperties;

        newRun.AppendChild(new W.Text(modifiedText) { Space = SpaceProcessingModeValues.Preserve });
        paragraph.AppendChild(newRun);
    }

    private string GetParagraphText(W.Paragraph paragraph)
    {
        var textBuilder = new StringBuilder();

        foreach (var text in paragraph.Descendants<W.Text>())
        {
            textBuilder.Append(text.Text);
        }

        return textBuilder.ToString();
    }

    // Alternative approach that preserves more complex formatting
    private void ReplacePlaceholdersAdvanced(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        foreach (var paragraph in element.Descendants<W.Paragraph>())
        {
            if (paragraph.Ancestors<W.Table>().Any())
                continue;

            ReplacePlaceholdersWithFormattingPreservation(paragraph, placeholders);
        }
    }

    private void ReplacePlaceholdersWithFormattingPreservation(W.Paragraph paragraph, Dictionary<string, string> placeholders)
    {
        string paragraphText = GetParagraphText(paragraph);

        foreach (var kvp in placeholders)
        {
            if (paragraphText.Contains(kvp.Key))
            {
                // Find the position of the placeholder
                int placeholderStart = paragraphText.IndexOf(kvp.Key);
                int placeholderEnd = placeholderStart + kvp.Key.Length;

                // Find which runs contain the placeholder
                var runs = paragraph.Descendants<W.Run>().ToList();
                int currentPosition = 0;
                W.Run startRun = null;
                W.Run endRun = null;
                int startRunOffset = 0;
                int endRunOffset = 0;

                foreach (var run in runs)
                {
                    string runText = run.InnerText;
                    int runLength = runText.Length;

                    if (startRun == null && currentPosition + runLength > placeholderStart)
                    {
                        startRun = run;
                        startRunOffset = placeholderStart - currentPosition;
                    }

                    if (currentPosition + runLength >= placeholderEnd)
                    {
                        endRun = run;
                        endRunOffset = placeholderEnd - currentPosition;
                        break;
                    }

                    currentPosition += runLength;
                }

                if (startRun != null && endRun != null)
                {
                    ReplaceTextInRuns(paragraph, startRun, endRun, startRunOffset, endRunOffset, kvp.Key, kvp.Value);
                    // Refresh paragraph text for next replacement
                    paragraphText = GetParagraphText(paragraph);
                }
            }
        }
    }

    private void ReplaceTextInRuns(W.Paragraph paragraph, W.Run startRun, W.Run endRun, int startOffset, int endOffset, string placeholder, string replacement)
    {
        if (startRun == endRun)
        {
            // Placeholder is within a single run
            var textElement = startRun.Descendants<W.Text>().FirstOrDefault();
            if (textElement != null)
            {
                string originalText = textElement.Text;
                string beforePlaceholder = originalText.Substring(0, startOffset);
                string afterPlaceholder = originalText.Substring(endOffset);
                textElement.Text = beforePlaceholder + replacement + afterPlaceholder;
            }
        }
        else
        {
            // Placeholder spans multiple runs - more complex handling needed
            // This is a simplified approach - you might need more sophisticated logic
            var runsToProcess = GetRunsBetween(paragraph, startRun, endRun);

            // Get the formatting from the first run
            var formatting = startRun.RunProperties?.CloneNode(true) as W.RunProperties;

            // Remove the runs that contain the placeholder
            foreach (var run in runsToProcess)
            {
                run.Remove();
            }

            // Create a new run with the replacement text
            var newRun = new W.Run();
            if (formatting != null)
                newRun.RunProperties = formatting;

            newRun.AppendChild(new W.Text(replacement) { Space = SpaceProcessingModeValues.Preserve });

            // Insert the new run at the position of the first removed run
            if (startRun.NextSibling() != null)
                startRun.NextSibling().InsertBeforeSelf(newRun);
            else
                paragraph.AppendChild(newRun);
        }
    }

    private List<W.Run> GetRunsBetween(W.Paragraph paragraph, W.Run startRun, W.Run endRun)
    {
        var runs = paragraph.Descendants<W.Run>().ToList();
        var result = new List<W.Run>();
        bool collecting = false;

        foreach (var run in runs)
        {
            if (run == startRun)
                collecting = true;

            if (collecting)
                result.Add(run);

            if (run == endRun)
                break;
        }

        return result;
    }


}

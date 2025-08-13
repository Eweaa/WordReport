using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using System.IO.Compression;
using System.Text;
using WordReport.Models;
using WordReport.ViewModels;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace WordReport.Services;

public class NewWordService
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

                var xmlPathBefore = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(templatePath), "documentBefore.xml");
                var docPartBefore = wordDoc.MainDocumentPart;

                using (var reader = new StreamReader(docPartBefore.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    var xmlContent = reader.ReadToEnd();
                    File.WriteAllText(xmlPathBefore, xmlContent);
                }


                //var placeholders = model.GetType()
                //.GetProperties()
                //.Where(p => p.PropertyType == typeof(string) || p.PropertyType.IsValueType)
                //.Select(p => new { Name = p.Name, Value = p.GetValue(model) })
                //.Where(p => p.Value != null)
                //.ToDictionary(p => p.Name, p => p.Value!.ToString());


                var placeholders = new Dictionary<string, string>
                {
                    { "{{For}}", "The Great User" },
                    { "{{Subject}}", "Important Subject" },
                    { "{{ProposalReference}}", "No Proposal Reference" },
                    { "{{ProposalDate}}", DateTime.Now.ToString("dd MMMM yyyy") },
                    { "{{ValidFor}}", "7 Days" }
                };

                // Replace in body
                ReplacePlaceholders(wordDoc.MainDocumentPart.Document.Body, placeholders);

                // Replace in headers and handle logo
                //foreach (var header in wordDoc.MainDocumentPart.HeaderParts)
                //{
                //    ReplacePlaceholders(header.Header, placeholders);
                //    if (model.Logo != null)
                //    {
                //        ReplaceLogoInHeader(header, model.Logo);
                //    }
                //}

                // Replace in footers
                //foreach (var footer in wordDoc.MainDocumentPart.FooterParts)
                //{
                //    ReplacePlaceholders(footer.Footer, placeholders);
                //}

                // Add rows to the table
                //AddTableRows(wordDoc.MainDocumentPart.Document.Body, products);

                AddQuotationTableAfterTitle(wordDoc.MainDocumentPart.Document.Body, quotations);
                AddTableRowsByIndex(wordDoc.MainDocumentPart.Document.Body, tests, 4);


                //UpdateFooterDatePlaceholder(wordDoc);
                //AddQuotationTable(wordDoc.MainDocumentPart.Document.Body, quotations, 5);

                wordDoc.MainDocumentPart.Document.Save();


                // Generate XML file from the header
                //var xmlHeaderPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(templatePath), "header.xml");
                //var docHeaderPart = wordDoc.MainDocumentPart.HeaderParts.First();

                //using (var reader = new StreamReader(docHeaderPart.GetStream(FileMode.Open, FileAccess.Read)))
                //{
                //    var xmlContent = reader.ReadToEnd();
                //    File.WriteAllText(xmlHeaderPath, xmlContent);
                //}


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

    //private void ReplacePlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    //{
    //    foreach (var text in element.Descendants<W.Text>())
    //    {
    //        if (text.Ancestors<W.Table>().Any())
    //            continue;

    //        foreach (var key in placeholders.Keys)
    //        {
    //            if (text.Text.Contains(key))
    //            {
    //                text.Text = text.Text.Replace(key, placeholders[key]);
    //            }
    //        }
    //    }
    //}


    //private void ReplacePlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    //{
    //    // Get all text elements that are not inside tables
    //    var texts = element
    //        .Descendants<W.Text>()
    //        .Where(t => !t.Ancestors<W.Table>().Any())
    //        .ToList();

    //    if (!texts.Any()) return;

    //    // Combine all text into one string
    //    string combinedText = string.Join("", texts.Select(t => t.Text));

    //    // Replace placeholders
    //    foreach (var kvp in placeholders)
    //    {
    //        combinedText = combinedText.Replace(kvp.Key, kvp.Value);
    //    }

    //    // Clear all but first run's text
    //    texts[0].Text = combinedText;
    //    for (int i = 1; i < texts.Count; i++)
    //    {
    //        texts[i].Text = string.Empty;
    //    }
    //}

    private void ReplaceLogoInHeader(HeaderPart headerPart, IFormFile logoFile)
    {
        // Find the "{Logo}" placeholder in the header
        var logoPlaceholder = headerPart.Header.Descendants<W.Text>()
            .FirstOrDefault(t => t.Text.Contains("Logo"));

        if (logoPlaceholder == null) return;

        // Get the parent run of the placeholder
        var run = logoPlaceholder.Ancestors<W.Run>().FirstOrDefault();
        if (run == null) return;

        // Create image part with explicit type handling
        ImagePart imagePart;
        string extension = System.IO.Path.GetExtension(logoFile.FileName).ToLower();

        switch (extension)
        {
            case ".jpg":
            case ".jpeg":
                imagePart = headerPart.AddImagePart(ImagePartType.Jpeg);
                break;
            case ".png":
                imagePart = headerPart.AddImagePart(ImagePartType.Png);
                break;
            case ".gif":
                imagePart = headerPart.AddImagePart(ImagePartType.Gif);
                break;
            case ".bmp":
                imagePart = headerPart.AddImagePart(ImagePartType.Bmp);
                break;
            case ".tiff":
            case ".tif":
                imagePart = headerPart.AddImagePart(ImagePartType.Tiff);
                break;
            default:
                imagePart = headerPart.AddImagePart(ImagePartType.Jpeg);
                break;
        }

        // Copy image data
        using (var stream = logoFile.OpenReadStream())
        {
            imagePart.FeedData(stream);
        }

        // Get image dimensions
        var imageSize = GetImageSize(logoFile);
        long widthEmu = imageSize.Width * 9525; // Convert pixels to EMUs
        long heightEmu = imageSize.Height * 9525;

        // Limit maximum size
        const long maxWidthEmu = 1440000; // about 1 inch
        const long maxHeightEmu = 1440000;

        if (widthEmu > maxWidthEmu)
        {
            heightEmu = (long)(heightEmu * ((double)maxWidthEmu / widthEmu));
            widthEmu = maxWidthEmu;
        }
        if (heightEmu > maxHeightEmu)
        {
            widthEmu = (long)(widthEmu * ((double)maxHeightEmu / heightEmu));
            heightEmu = maxHeightEmu;
        }

        string relationshipId = headerPart.GetIdOfPart(imagePart);

        // Create the drawing element with proper namespace prefixes
        var drawing = new W.Drawing(
            new DW.Inline(
                new DW.Extent() { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties()
                {
                    Id = 1U,
                    Name = "Logo"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties()
                                {
                                    Id = 0U,
                                    Name = "Logo"
                                },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip() { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry(
                                    new A.AdjustValueList()
                                )
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                }
                            )
                        )
                    )
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            }
        );

        // Replace the placeholder text with the image
        logoPlaceholder.Text = "";
        run.AppendChild(drawing);
    }

    private Size GetImageSize(IFormFile imageFile)
    {
        try
        {
            using var stream = imageFile.OpenReadStream();
            using var image = Image.FromStream(stream);
            return new Size(image.Width, image.Height);
        }
        catch
        {
            // Return default size if unable to read image dimensions
            return new Size(100, 100);
        }
    }

    private void AddTableRows<T>(W.Body body, List<T> items)
    {
        if (items == null || items.Count == 0) return;

        var table = body.Descendants<W.Table>().FirstOrDefault();
        if (table == null) return;

        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count < 2) return;

        var labelRow = rows[0]; // First row (display names)
        var propertyRow = rows[1]; // Second row (property names)

        // Get property names from the second row
        var propertyNames = propertyRow.Elements<W.TableCell>()
            .Select(cell => cell.InnerText.Trim())
            .ToList();

        // Map property names to actual Product properties
        var productType = typeof(T);

        var props = propertyNames
            .Select(name => productType.GetProperties().FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)))
            //.Where(p => p != null)
            .ToList();

        // Remove all rows except the first (label) row
        for (int i = 1; i < rows.Count; i++)
        {
            rows[i].Remove();
        }

        // Add data rows
        foreach (var item in items)
        {
            var row = new W.TableRow();
            foreach (var prop in props)
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


    private void AddTableRowsByBookmark<T>(W.Body body, List<T> items, string bookmarkName)
    {
        if (items == null || items.Count == 0) return;

        var bookmark = body.Descendants<W.BookmarkStart>()
                           .FirstOrDefault(b => b.Name == bookmarkName);
        if (bookmark == null) return;

        var table = bookmark.Parent.Descendants<W.Table>().FirstOrDefault();
        if (table == null) return;

        PopulateTableWithItems(table, items);
    }


    private void AddTableRowsByKeyword<T>(W.Body body, List<T> items, string keyword)
    {
        if (items == null || items.Count == 0) return;

        var table = body.Descendants<W.Table>()
                        .FirstOrDefault(t => t.InnerText.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        if (table == null) return;

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
        var titleParagraph = body
        .Descendants<W.Paragraph>()
        .FirstOrDefault(p =>
            p.InnerText.Trim().Equals("Quotation", StringComparison.OrdinalIgnoreCase) &&
            p.Ancestors<W.Table>().Count() == 0
        );

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
            row.Append(CreateCenteredCell(index.ToString("00")));          // No.
            row.Append(CreateCenteredCell(item.Deliverable ?? ""));        // Deliverable
            row.Append(CreateCenteredCell(item.Unit ?? ""));               // Unit
            row.Append(CreateCenteredCell(unitCost.ToString("N2")));       // Unit Cost
            row.Append(CreateCenteredCell(qty.ToString()));                // Qty
            row.Append(CreateCenteredCell(total.ToString("N2")));          // Total Cost
            table.Append(row);
            index++;
        }


        //  Spanned Cells
        table.Append(CreateSummaryRow("Subtotal (without VAT)", subtotal, 5));
        table.Append(CreateSummaryRow("VAT", quotation.Vat ?? 0, 5));
        table.Append(CreateSummaryRow("Total", quotation.FinalTotal ?? 0, 5, bold: true));
    }

    private void AddQuotationTable(W.Body body, Quotation quotation, int tableIndex)
    {
        var table = body.Descendants<W.Table>().ElementAtOrDefault(tableIndex);
        if (table == null) throw new InvalidOperationException($"No table found at index {tableIndex}");

        ApplyTableBorders(table);

        // Clear old rows except header
        var rows = table.Elements<W.TableRow>().ToList();
        for (int i = 1; i < rows.Count; i++)
            rows[i].Remove();

        int index = 1;
        decimal subtotal = 0;

        // Add item rows
        foreach (var item in quotation.Items)
        {
            decimal unitCost = item.UnitCost ?? 0;
            int qty = item.Quantity ?? 0;
            decimal total = item.TotalCost ?? (unitCost * qty);

            subtotal += total;

            var row = new W.TableRow();
            row.Append(CreateCell(index.ToString("00")));                      // No.
            row.Append(CreateCell(item.Deliverable ?? ""));                    // Deliverable
            row.Append(CreateCell(item.Unit ?? ""));                           // Unit
            row.Append(CreateCell(unitCost.ToString("N2")));                   // Unit Cost
            row.Append(CreateCell(qty.ToString()));                            // Qty
            row.Append(CreateCell(total.ToString("N2")));                      // Total Cost
            table.Append(row);
            index++;
        }

        // Summary rows
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


    private void UpdateFooterDatePlaceholder(WordprocessingDocument wordDoc)
    {
        string today = DateTime.Now.ToString("dd MMMM yyyy");

        foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
        {
            foreach (var paragraph in footerPart.Footer.Descendants<W.Paragraph>())
            {
                var texts = paragraph.Descendants<W.Text>().ToList();
                if (texts.Count == 0) continue;

                // Combine text values to detect placeholder even if split
                string fullText = string.Concat(texts.Select(t => t.Text));

                if (fullText.Contains("{{Date}}"))
                {
                    // New combined string
                    string updatedText = fullText.Replace("{{Date}}", today);

                    // Now redistribute updated text back into the same runs
                    int pos = 0;
                    foreach (var text in texts)
                    {
                        int remaining = updatedText.Length - pos;
                        if (remaining <= 0)
                        {
                            text.Text = "";
                            continue;
                        }

                        // Fill run's text with a portion of updated string
                        int length = Math.Min(text.Text.Length, remaining);
                        text.Text = updatedText.Substring(pos, length);
                        pos += length;
                    }
                }
            }
        }
    }


    //private void ReplacePlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    //{
    //    // Process paragraphs instead of individual text elements
    //    foreach (var paragraph in element.Descendants<W.Paragraph>())
    //    {
    //        if (paragraph.Ancestors<W.Table>().Any())
    //            continue;

    //        ReplacePlaceholdersInParagraph(paragraph, placeholders);
    //    }
    //}

    //private void ReplacePlaceholdersInParagraph(W.Paragraph paragraph, Dictionary<string, string> placeholders)
    //{
    //    // Get all text content from the paragraph
    //    string paragraphText = GetParagraphText(paragraph);

    //    // Check if any placeholder exists in the paragraph
    //    bool hasReplacements = false;
    //    string modifiedText = paragraphText;

    //    foreach (var kvp in placeholders)
    //    {
    //        if (modifiedText.Contains(kvp.Key))
    //        {
    //            modifiedText = modifiedText.Replace(kvp.Key, kvp.Value);
    //            hasReplacements = true;
    //        }
    //    }

    //    // If no replacements needed, return early
    //    if (!hasReplacements)
    //        return;

    //    // Get the first run's formatting to preserve style
    //    var firstRun = paragraph.Descendants<W.Run>().FirstOrDefault();
    //    var runProperties = firstRun?.RunProperties?.CloneNode(true) as W.RunProperties;

    //    // Clear all existing runs
    //    paragraph.RemoveAllChildren<W.Run>();

    //    // Create a new run with the modified text and original formatting
    //    var newRun = new W.Run();
    //    if (runProperties != null)
    //        newRun.RunProperties = runProperties;

    //    newRun.AppendChild(new W.Text(modifiedText) { Space = SpaceProcessingModeValues.Preserve });
    //    paragraph.AppendChild(newRun);
    //}

    //private string GetParagraphText(W.Paragraph paragraph)
    //{
    //    var textBuilder = new StringBuilder();

    //    foreach (var text in paragraph.Descendants<W.Text>())
    //    {
    //        textBuilder.Append(text.Text);
    //    }

    //    return textBuilder.ToString();
    //}

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

    private void ReplaceTextInRuns(W.Paragraph paragraph, W.Run startRun, W.Run endRun,
        int startOffset, int endOffset, string placeholder, string replacement)
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



    private void ReplacePlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        // Process paragraphs in the main document body (outside tables)
        foreach (var paragraph in element.Descendants<W.Paragraph>().Where(p => !p.Ancestors<W.Table>().Any()))
        {
            ReplacePlaceholdersInParagraph(paragraph, placeholders);
        }

        // Process paragraphs in tables (first table only, or all tables as needed)
        ProcessTablesPlaceholders(element, placeholders);

        // Process paragraphs in text boxes and other drawing elements
        ProcessDrawingElementsPlaceholders(element, placeholders);
    }

    private void ProcessTablesPlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        // Get the first table only (adjust logic if you need to process all tables)
        var firstTable = element.Descendants<W.Table>().FirstOrDefault();
        if (firstTable != null)
        {
            foreach (var paragraph in firstTable.Descendants<W.Paragraph>())
            {
                ReplacePlaceholdersInParagraph(paragraph, placeholders);
            }
        }

        // If you want to process ALL tables instead of just the first one, use this:
        /*
        foreach (var table in element.Descendants<W.Table>())
        {
            foreach (var paragraph in table.Descendants<W.Paragraph>())
            {
                ReplacePlaceholdersInParagraph(paragraph, placeholders);
            }
        }
        */
    }

    private void ProcessDrawingElementsPlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        // Process text boxes in drawings (like in your XML example)
        foreach (var textBox in element.Descendants<W.Drawing>())
        {
            // Handle both modern drawing format and legacy VML format
            var textBoxContent = textBox.Descendants().Where(e =>
                e.LocalName == "txbxContent" &&
                (e.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ||
                 e.NamespaceUri.Contains("office")));

            foreach (var content in textBoxContent)
            {
                foreach (var paragraph in content.Descendants<W.Paragraph>())
                {
                    ReplacePlaceholdersInParagraph(paragraph, placeholders);
                }
            }
        }

        // Handle VML text boxes (legacy format like v:textbox in your XML)
        var vmlNamespace = "urn:schemas-microsoft-com:vml";
        foreach (var vmlElement in element.Descendants().Where(e => e.NamespaceUri == vmlNamespace))
        {
            foreach (var paragraph in vmlElement.Descendants<W.Paragraph>())
            {
                ReplacePlaceholdersInParagraph(paragraph, placeholders);
            }
        }
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
        var runsToRemove = paragraph.Descendants<W.Run>().ToList();
        foreach (var run in runsToRemove)
        {
            run.Remove();
        }

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

    // Alternative comprehensive method that handles all scenarios in one go
    private void ReplacePlaceholdersComprehensive(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        // Get all paragraphs from all locations (main document, tables, text boxes, etc.)
        var allParagraphs = GetAllParagraphs(element);

        foreach (var paragraph in allParagraphs)
        {
            ReplacePlaceholdersInParagraph(paragraph, placeholders);
        }
    }

    private IEnumerable<W.Paragraph> GetAllParagraphs(OpenXmlElement element)
    {
        var paragraphs = new List<W.Paragraph>();

        // Main document paragraphs (outside tables)
        paragraphs.AddRange(element.Descendants<W.Paragraph>().Where(p => !p.Ancestors<W.Table>().Any()));

        // Table paragraphs (first table only - modify as needed)
        var firstTable = element.Descendants<W.Table>().FirstOrDefault();
        if (firstTable != null)
        {
            paragraphs.AddRange(firstTable.Descendants<W.Paragraph>());
        }

        // Drawing/TextBox paragraphs
        foreach (var drawing in element.Descendants<W.Drawing>())
        {
            paragraphs.AddRange(drawing.Descendants<W.Paragraph>());
        }

        // VML TextBox paragraphs (legacy format)
        var vmlNamespace = "urn:schemas-microsoft-com:vml";
        foreach (var vmlElement in element.Descendants().Where(e => e.NamespaceUri == vmlNamespace))
        {
            paragraphs.AddRange(vmlElement.Descendants<W.Paragraph>());
        }

        // Header/Footer paragraphs if needed
        foreach (var headerFooter in element.Descendants().Where(e =>
            e.LocalName == "hdr" || e.LocalName == "ftr"))
        {
            paragraphs.AddRange(headerFooter.Descendants<W.Paragraph>());
        }

        return paragraphs;
    }

    // Utility method to safely handle null placeholders
    private void ReplacePlaceholdersSafe(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        if (element == null || placeholders == null || !placeholders.Any())
            return;

        // Filter out null or empty keys/values
        var safePlaceholders = placeholders
            .Where(kvp => !string.IsNullOrEmpty(kvp.Key) && kvp.Value != null)
            .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

        if (safePlaceholders.Any())
        {
            ReplacePlaceholdersComprehensive(element, safePlaceholders);
        }
    }

}

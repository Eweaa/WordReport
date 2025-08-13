using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Routing.Template;
using System.Diagnostics;
using System.Drawing;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using WordReport.Models;
using WordReport.ViewModels;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
//using GemBox.Document;

namespace WordReport.Services;

public class WordService
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
                //var placeholders = model.GetType()
                //.GetProperties()
                //.Where(p => p.PropertyType == typeof(string) || p.PropertyType.IsValueType)
                //.Select(p => new { Name = p.Name, Value = p.GetValue(model) })
                //.Where(p => p.Value != null)
                //.ToDictionary(p => p.Name, p => p.Value!.ToString());


                var placeholders = new Dictionary<string, string>
                {
                    { "{For}", "The Great User" },
                    { "{Subject}", "Important Subject" },
                    { "{ProposalReference}", "No Proposal Reference" },
                    { "{ProposalDate}", DateTime.Now.ToString("dd MMMM yyyy") },
                    { "{ValidFor}", "7 Days" }
                };

                var bodyPlaceholders = new Dictionary<string, string>
                {
                    { "{For}", "The Great User" },
                    { "{Company}", "SomeCompany" },
                    { "{Location}", "Some Location" },
                    { "{RoutineAnalysis}", "80 Days" },         
                    { "{SubContractedParameters}", "70 Days" }  
                };

                // Replace in body
                //ReplacePlaceholders(wordDoc.MainDocumentPart.Document.Body, placeholders);

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

                //ReplacePlaceholdersFirstPage(wordDoc.MainDocumentPart.Document.Body, "{For}", "The Great User");
                //ReplacePlaceholders1(wordDoc.MainDocumentPart.Document.Body, bodyPlaceholders);
                //ReplacePlaceholders2(wordDoc.MainDocumentPart.Document.Body, bodyPlaceholders);

                ReplacePlaceholders3(wordDoc.MainDocumentPart.Document.Body, bodyPlaceholders);
                ReplacePlaceholdersInitialTable(wordDoc.MainDocumentPart.Document.Body, placeholders);
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


                //  Generate XML file from the edited document
                
                //  NEEDS TO BE TESTED BEFORE DELETING CODE BENEATH IT
                GenerateXmlFile(templatePath, wordDoc, "document");

                //var xmlPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(templatePath), "document.xml");
                //var docPart = wordDoc.MainDocumentPart;

                //using (var reader = new StreamReader(docPart.GetStream(FileMode.Open, FileAccess.Read)))
                //{
                //    var xmlContent = reader.ReadToEnd();
                //    File.WriteAllText(xmlPath, xmlContent);
                //}
            }

            return mem.ToArray();

        }
    }

    public byte[] GenerateProposalPdfWithLibreOffice(DocumentViewModel model, string templatePath)
    {
        // Create a temp working folder
        string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        // Paths for the DOCX and final PDF
        string editedDocxPath = System.IO.Path.Combine(tempDir, "proposal.docx");

        // Copy template so we don't overwrite the original
        File.Copy(templatePath, editedDocxPath, true);

        // Edit Word File
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(editedDocxPath, true))
        {
            var products = DataService.GetProducts();
            var tests = DataService.GetTests();
            var quotations = DataService.GetQuotations();

            var placeholders = new Dictionary<string, string>
            {
                { "{For}", "The Great User" },
                { "{Subject}", "Important Subject" },
                { "{ProposalReference}", "No Proposal Reference" },
                { "{ProposalDate}", DateTime.Now.ToString("dd MMMM yyyy") },
                { "{ValidFor}", "7 Days" }
            };

            var bodyPlaceholders = new Dictionary<string, string>
            {
                { "{For}", "The Great User" },
                { "{Company}", "SomeCompany" },
                { "{Location}", "Some Location" },
                { "{RoutineAnalysis}", "80 Days" },
                { "{SubContractedParameters}", "70 Days" }
            };

            // Replace in body
            ReplacePlaceholders3(wordDoc.MainDocumentPart.Document.Body, bodyPlaceholders);
            ReplacePlaceholdersInitialTable(wordDoc.MainDocumentPart.Document.Body, placeholders);

            // Add tables
            AddQuotationTableAfterTitle(wordDoc.MainDocumentPart.Document.Body, quotations);
            AddTableRowsByIndex(wordDoc.MainDocumentPart.Document.Body, tests, 4);

            // Save changes
            wordDoc.MainDocumentPart.Document.Save();


            // Generate XML file for the edited word file
            GenerateXmlFile(templatePath, wordDoc, "document");
        }

        // Convert Word to PDF
        var processInfo = new ProcessStartInfo
        {
            FileName = "soffice",
            Arguments = $"--headless --convert-to pdf --outdir \"{tempDir}\" \"{editedDocxPath}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        string stdOut, stdErr;
        using (var process = Process.Start(processInfo))
        {
            stdOut = process.StandardOutput.ReadToEnd();
            stdErr = process.StandardError.ReadToEnd();
            process.WaitForExit();
        }

        Console.WriteLine("LibreOffice Output:");
        Console.WriteLine(stdOut);
        Console.WriteLine("LibreOffice Error:");
        Console.WriteLine(stdErr);

        // Wait for PDF to appear
        string pdfPath = null;
        for (int i = 0; i < 10; i++) // Try up to ~5 seconds
        {
            pdfPath = Directory.GetFiles(tempDir, "*.pdf").FirstOrDefault();
            if (pdfPath != null && File.Exists(pdfPath))
                break;

            Thread.Sleep(500);
        }

        if (pdfPath == null || !File.Exists(pdfPath))
            throw new FileNotFoundException("LibreOffice did not create a PDF file.", pdfPath ?? "(unknown)");

        // Read PDF into byte array
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);

        // Clean up
        try { Directory.Delete(tempDir, true); } catch { }

        return pdfBytes;
    }

    // Example document generation stub
    private void GenerateWordDocumentFromTemplate(DocumentViewModel model, string templatePath, string outputPath)
    {
        // TODO: Fill DOCX generation logic here
        File.Copy(templatePath, outputPath, true);
    }

    public void GenerateXmlFile(string templatePath, WordprocessingDocument wordDocument, string outputFileName)
    {
        var xmlPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(templatePath), $"{outputFileName}.xml");
        var docPart = wordDocument.MainDocumentPart;

        using (var reader = new StreamReader(docPart.GetStream(FileMode.Open, FileAccess.Read)))
        {
            var xmlContent = reader.ReadToEnd();
            File.WriteAllText(xmlPath, xmlContent);
        }
    }


    //public byte[] GeneratePdfFromTemplate(string templatePath)
    //{
    //    var products = DataService.GetProducts();
    //    var tests = DataService.GetTests();
    //    var quotations = DataService.GetQuotations();

    //    byte[] byteArray = File.ReadAllBytes(templatePath);

    //    using (MemoryStream mem = new MemoryStream())
    //    {
    //        mem.Write(byteArray, 0, byteArray.Length);

    //        // 1️⃣ Modify the DOCX in memory
    //        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
    //        {
    //            var placeholders = new Dictionary<string, string>
    //            {
    //                { "{{For}}", "The Great User" },
    //                { "{{Subject}}", "Important Subject" },
    //                { "{{ProposalReference}}", "No Proposal Reference" },
    //                { "{{ProposalDate}}", DateTime.Now.ToString("dd MMMM yyyy") },
    //                { "{{ValidFor}}", "7 Days" }
    //            };

    //            var bodyPlaceholders = new Dictionary<string, string>
    //            {
    //                { "{{For}}", "The Great User" },
    //                { "{{Company}}", "SomeCompany" },
    //                { "{{Location}}", "Some Location" },
    //                { "{{RoutineAnalysis}}", "80 Days" },
    //                { "{{SubcontractedParameters}}", "70 Days" }
    //            };

    //            ReplacePlaceholdersInitialTable(wordDoc.MainDocumentPart.Document.Body, placeholders);
    //            AddQuotationTableAfterTitle(wordDoc.MainDocumentPart.Document.Body, quotations);
    //            AddTableRowsByIndex(wordDoc.MainDocumentPart.Document.Body, tests, 4);

    //            wordDoc.MainDocumentPart.Document.Save();

    //            // Optional: Debugging - save XML
    //            var xmlPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(templatePath), "document.xml");
    //            var docPart = wordDoc.MainDocumentPart;

    //            using (var reader = new StreamReader(docPart.GetStream(FileMode.Open, FileAccess.Read)))
    //            {
    //                var xmlContent = reader.ReadToEnd();
    //                File.WriteAllText(xmlPath, xmlContent);
    //            }
    //        }

    //        // 2️⃣ Convert the updated DOCX in memory to PDF using GemBox.Document
    //        mem.Position = 0; // reset stream position before reading
    //        var document = DocumentModel.Load(mem, GemBox.Document.LoadOptions.DocxDefault);

    //        using (MemoryStream pdfStream = new MemoryStream())
    //        {
    //            document.Save(pdfStream, GemBox.Document.SaveOptions.PdfDefault);
    //            return pdfStream.ToArray();
    //        }
    //    }
    //}


    public static void ReplacePlaceholdersInitialTable(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        if (element == null || placeholders == null || placeholders.Count == 0)
            return;

        var firstTable = element.Elements<W.Table>().FirstOrDefault();

        if (firstTable == null)
            return;

        //foreach (var text in firstTable.Descendants<W.Text>())
        //{
        //    foreach (var kvp in placeholders)
        //    {
        //        if (!string.IsNullOrEmpty(text.Text) && text.Text.Contains(kvp.Key))
        //        {
        //            text.Text = text.Text.Replace(kvp.Key, kvp.Value ?? string.Empty);
        //        }
        //    }
        //}


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



    //private void replaceplaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
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
    //    var texts = element.Descendants<W.Text>().Where(t => !t.Ancestors<W.Table>().Any()).ToList();

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

    private void ReplacePlaceholdersFirstPage(OpenXmlElement element, string Key, string Value)
    {
        var texts = element.Descendants<W.Text>().Where(t => !t.Ancestors<W.Table>().Any()).ToList();

        foreach (var text in texts)
        {
            string originalText = text.Text;
            originalText = originalText.Replace(Key, Value);
            text.Text = originalText;
        }
    }

    private void ReplacePlaceholders1(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        var texts = element.Descendants<W.Text>().Where(t => !t.Ancestors<W.Table>().Any()).ToList();

        foreach (var text in texts)
        {
            string originalText = text.Text;
            foreach (var kvp in placeholders)
            {
                originalText = originalText.Replace(kvp.Key, kvp.Value);
            }
            text.Text = originalText;
        }
    }

    private void ReplacePlaceholders2(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        var paragraphs = element.Descendants<W.Paragraph>()
        .Where(p => !p.Ancestors<W.Table>().Any()).ToList();

        foreach (var paragraph in paragraphs)
        {
            var texts = paragraph.Descendants<W.Text>().ToList();
            if (!texts.Any()) continue;

            string combinedText = string.Join("", texts.Select(t => t.Text));

            foreach (var kvp in placeholders)
            {
                combinedText = combinedText.Replace(kvp.Key, kvp.Value);
            }

            // Only modify if there are changes
            if (texts.Count > 1 || texts[0].Text != combinedText)
            {
                texts[0].Text = combinedText;
                for (int i = 1; i < texts.Count; i++)
                {
                    texts[i].Text = string.Empty;
                }
            }
        }
    }

    private void ReplacePlaceholders3(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        // Get all paragraphs that are not inside tables
        var paragraphs = element.Descendants<W.Paragraph>()
            .Where(p => !p.Ancestors<W.Table>().Any())
            .ToList();

        foreach (var paragraph in paragraphs)
        {
            ReplacePlaceholdersInParagraph1(paragraph, placeholders);
        }
    }


    // Don't Remove yet to Test
    //private void ReplacePlaceholdersAdvanced(OpenXmlElement element, Dictionary<string, string> placeholders)
    //{
    //    var paragraphs = element.Descendants<W.Paragraph>()
    //        .Where(p => !p.Ancestors<W.Table>().Any())
    //        .ToList();

    //    foreach (var paragraph in paragraphs)
    //    {
    //        // Check if this paragraph contains any placeholder
    //        string fullText = paragraph.InnerText;
    //        bool containsPlaceholder = placeholders.Keys.Any(key => fullText.Contains(key));

    //        if (!containsPlaceholder)
    //            continue;

    //        // Try individual text replacement first
    //        var texts = paragraph.Descendants<W.Text>().ToList();
    //        bool simpleReplacementWorked = false;

    //        foreach (var text in texts)
    //        {
    //            string originalText = text.Text;
    //            string newText = originalText;

    //            foreach (var kvp in placeholders)
    //            {
    //                newText = newText.Replace(kvp.Key, kvp.Value);
    //            }

    //            if (newText != originalText)
    //            {
    //                text.Text = newText;
    //                simpleReplacementWorked = true;
    //            }
    //        }

    //        // If simple replacement didn't work, use consolidation approach
    //        if (!simpleReplacementWorked)
    //        {
    //            string combinedText = string.Join("", texts.Select(t => t.Text));
    //            string processedText = combinedText;

    //            foreach (var kvp in placeholders)
    //            {
    //                processedText = processedText.Replace(kvp.Key, kvp.Value);
    //            }

    //            if (processedText != combinedText)
    //            {
    //                texts[0].Text = processedText;
    //                for (int i = 1; i < texts.Count; i++)
    //                {
    //                    texts[i].Text = string.Empty;
    //                }
    //            }
    //        }
    //    }
    //}


    private void ReplacePlaceholdersInParagraph1(W.Paragraph paragraph, Dictionary<string, string> placeholders)
    {
        var texts = paragraph.Descendants<W.Text>().ToList();
        if (!texts.Any()) return;

        // First => Simple Replace If The Entire Key is in One Tag
        bool anyReplaced = false;
        foreach (var text in texts)
        {
            string originalText = text.Text;
            string modifiedText = originalText;

            foreach (var kvp in placeholders)
            {
                modifiedText = modifiedText.Replace(kvp.Key, kvp.Value);
            }

            if (modifiedText != originalText)
            {
                text.Text = modifiedText;
                anyReplaced = true;
            }
        }

        // If Simple Replace Worked Thank You Very Much and سلامو عليكوا
        if (anyReplaced)
            return;


        // If Simple Replce Didn't Work We Need To Try Complex Replace (Where The Key Is Split Across Different Tags)
        string combinedText = string.Join("", texts.Select(t => t.Text));
        string processedText = combinedText;

        foreach (var kvp in placeholders)
        {
            processedText = processedText.Replace(kvp.Key, kvp.Value);
        }

        // If Something Is Replaced
        if (processedText != combinedText)
        {
            // Combine The Text into The First Element & Make The Rest Empty
            texts[0].Text = processedText;
            for (int i = 1; i < texts.Count; i++)
            {
                texts[i].Text = string.Empty;
            }
        }
    }

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


}

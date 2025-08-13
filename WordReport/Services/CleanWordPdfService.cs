using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Diagnostics;
using WordReport.Models;
using WordReport.ViewModels;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace WordReport.Services;

public class CleanWordPdfService
{
    public byte[] GenerateProposalPdfWithLibreOffice(DocumentViewModel model, string templatePath)
    {
        // Create A Temp Working Folder
        string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        // Paths for the DOCX and final PDF
        string editedDocxPath = Path.Combine(tempDir, "Report.docx");

        // Copy template so we don't overwrite the original
        File.Copy(templatePath, editedDocxPath, true);

        // Edit Word File
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(editedDocxPath, true))
        {

            // Getting The Data (In-Memory)
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

            // Replacing The Data in Word File
            ReplacePlaceholders(wordDoc.MainDocumentPart.Document.Body, bodyPlaceholders);
            ReplacePlaceholdersInitialTable(wordDoc.MainDocumentPart.Document.Body, placeholders);
            AddQuotationTableAfterTitle(wordDoc.MainDocumentPart.Document.Body, quotations);
            AddTableRowsByIndex(wordDoc.MainDocumentPart.Document.Body, tests, 4);

            // Save changes
            wordDoc.MainDocumentPart.Document.Save();

            // Generate XML File for The Edited Word File
            GenerateXmlFile(templatePath, wordDoc, "document");
        }



        // Convert Word to PDF
        // VERY VERY VERY VERY IMPORTANT NOTE:
        // For this to work, you need to have LibreOffice installed and available in your system's PATH (Add it to Environment Varaible).
        // Ensure LibreOffice is installed and available in PATH
        // You May Need to Close The Terminal if it's open and Close Visual Studio if it's open
        // Run soffice -- version in the terminal to check if it's available
        // Restart the Computer if it's not showing in Terminal
        // Then Run This Code & It Should Generate The PDF File
        // قول يا رب
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


        // Debugging Code عدي يا معلم متركزش معاه إلا لو الكود ضرب
        Console.WriteLine("LibreOffice Output:");
        Console.WriteLine(stdOut);
        Console.WriteLine("LibreOffice Error:");
        Console.WriteLine(stdErr);

        // Wait for PDF to appear
        // This Used to Fail Sometimes Because LibreOffice Takes Time to Generate the PDF
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

    /// <summary>
    /// Generates an XML file from the Word document.
    /// </summary>
    /// <param name="templatePath">The Folder Path That The XML File Will Be Saved in</param>
    /// <param name="wordDocument">The Word Document That is Going to Be Transformed Into XML</param>
    /// <param name="outputFileName">The Output File Name -- Don't Add Extension Just The Name. The Method Will Add The Extension</param>
    public void GenerateXmlFile(string templatePath, WordprocessingDocument wordDocument, string outputFileName)
    {
        var xmlPath = Path.Combine(Path.GetDirectoryName(templatePath), $"{outputFileName}.xml");
        var docPart = wordDocument.MainDocumentPart;

        using (var reader = new StreamReader(docPart.GetStream(FileMode.Open, FileAccess.Read)))
        {
            var xmlContent = reader.ReadToEnd();
            File.WriteAllText(xmlPath, xmlContent);
        }
    }

    /// <summary>
    /// Replaces Keys in The Word Document Body Not The Tables
    /// </summary>
    /// <param name="element">The Element That is Going to Be Searched</param>
    /// <param name="placeholders">The Keys and The Values</param>
    private void ReplacePlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        // Get all paragraphs that are not inside tables
        var paragraphs = element.Descendants<W.Paragraph>()
            .Where(p => !p.Ancestors<W.Table>().Any())
            .ToList();

        foreach (var paragraph in paragraphs)
        {
            ReplacePlaceholdersInParagraph(paragraph, placeholders);
        }
    }


    private void ReplacePlaceholdersInParagraph(W.Paragraph paragraph, Dictionary<string, string> placeholders)
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


    /// <summary>
    /// Replaces Keys in The First Table of The Word Document Body
    /// </summary>
    /// <param name="element">The Element That is Going to Be Searched</param>
    /// <param name="placeholders">The Keys and The Values</param>
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

            // Combine All Runs into One String
            string combinedText = string.Concat(texts.Select(t => t.Text));

            // Replace All Placeholders in The Combined String
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

    /// <summary>
    /// The Method Finds The Table By Index and Adds Rows to It
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="body"></param>
    /// <param name="items"></param>
    /// <param name="tableIndex"></param>
    /// <exception cref="InvalidOperationException"></exception>
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


    /// <summary>
    /// The Most Important Method That Populates The Table With Items. Works For Now
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="table"></param>
    /// <param name="items"></param>
    private void PopulateTableWithItems<T>(W.Table table, List<T> items)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rows.Count < 2) return;


        // The Opposite => Because no Keys this time in the table 
        var labelRow = rows[1]; // First Row (display names)
        var propertyRow = rows[0]; // Second Row (property names)

        // Read Property Names From Table
        var propertyNames = propertyRow.Elements<W.TableCell>()
            .Select(cell => cell?.InnerText?.Trim() ?? "")
            .ToList();

        // Map each property name to its PropertyInfo (or null if not found)
        var props = propertyNames
            .Select(name => typeof(T).GetProperties()
            .FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        // Remove All The Rows Except The First One
        for (int i = 1; i < rows.Count; i++)
        {
            rows[i].Remove();
        }

        int index = 1;
        foreach (var item in items)
        {
            var row = new W.TableRow();

            // First Column => Index
            row.Append(CreateCell(index.ToString()));

            // Remaining Columns
            foreach (var prop in props.Skip(1)) // Skip Index
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

    /// <summary>
    /// Adds A Cell to The Table
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
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


    /// <summary>
    /// Add A Cell to The Table With Centered Text
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
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



    /// <summary>
    /// This Method Searches For The Table That Contains The Quotation Title
    /// </summary>
    /// <param name="body">Takes The Element That Is Going to Be Searched</param>
    /// <param name="quotation">Takes The Data That is Going to Be Used</param>
    /// <exception cref="InvalidOperationException"></exception>
    private void AddQuotationTableAfterTitle(W.Body body, Quotation quotation)
    {
        // Can Be Updated To Take The Title as A Parameter And Search For It


        // Find the Title Quotation
        var titleParagraph = body.Descendants<W.Paragraph>()
        .FirstOrDefault(p => p.InnerText.Trim().Equals("Quotation", StringComparison.OrdinalIgnoreCase) && p.Ancestors<W.Table>().Count() == 0);

        if (titleParagraph == null)
            throw new InvalidOperationException("No Quotation Title Found in The Document.");

        // Finding The First Table After The Title Quotation
        var table = titleParagraph.ElementsAfter().OfType<W.Table>().FirstOrDefault();

        if (table == null)
        {
            throw new InvalidOperationException("No Table Found After Quotation Title.");
        }

        ApplyTableBorders(table);

        // Remove All The Rows
        var rows = table.Elements<W.TableRow>().ToList();

        for (int i = 1; i < rows.Count; i++)
        {
            rows[i].Remove();
        }

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


    /// <summary>
    /// Adds A Cell to The Table That Is Spanned Across Multiple Columns
    /// </summary>
    /// <param name="label"></param>
    /// <param name="amount"></param>
    /// <param name="mergeColumns"></param>
    /// <param name="bold"></param>
    /// <returns></returns>
    private W.TableRow CreateSummaryRow(string label, decimal amount, int mergeColumns, bool bold = false)
    {
        var row = new W.TableRow();

        // Merged label cell
        row.Append(CreateCell(label, bold, mergeColumns));

        // Last column (Total Cost)
        row.Append(CreateCell(amount.ToString("N2"), bold));

        return row;
    }


    /// <summary>
    /// Overloaded Method to Create A Cell with Text, Bold Option, and Colspan
    /// </summary>
    /// <param name="text"></param>
    /// <param name="bold"></param>
    /// <param name="colspan"></param>
    /// <returns></returns>
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


    /// <summary>
    /// Add Borders to The Table
    /// </summary>
    /// <param name="table"></param>
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

}

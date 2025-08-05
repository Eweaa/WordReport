using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using System.IO.Compression;
using WordReport.ViewModels;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

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
                var placeholders = model.GetType()
                .GetProperties()
                .Where(p => p.PropertyType == typeof(string) || p.PropertyType.IsValueType)
                .Select(p => new { Name = p.Name, Value = p.GetValue(model) })
                .Where(p => p.Value != null)
                .ToDictionary(p => p.Name, p => p.Value!.ToString());

                // Replace in body
                ReplacePlaceholders(wordDoc.MainDocumentPart.Document.Body, placeholders);

                // Replace in headers and handle logo
                foreach (var header in wordDoc.MainDocumentPart.HeaderParts)
                {
                    ReplacePlaceholders(header.Header, placeholders);
                    if (model.Logo != null)
                    {
                        ReplaceLogoInHeader(wordDoc, header, model.Logo);
                    }
                }

                // Replace in footers
                foreach (var footer in wordDoc.MainDocumentPart.FooterParts)
                {
                    ReplacePlaceholders(footer.Footer, placeholders);
                }

                // Add rows to the table
                AddTableRows(wordDoc.MainDocumentPart.Document.Body, model.Tables);
                wordDoc.MainDocumentPart.Document.Save();


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

    private void ReplacePlaceholders(OpenXmlElement element, Dictionary<string, string> placeholders)
    {
        foreach (var text in element.Descendants<W.Text>())
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

    private void ReplaceLogoInHeader(WordprocessingDocument wordDoc, HeaderPart headerPart, IFormFile logoFile)
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

    private void AddTableRows(W.Body body, List<TableItemGroupViewModel> groups)
    {
        if (groups == null || groups.Count == 0) return;

        var tables = body.Descendants<W.Table>().ToList();

        foreach (var group in groups)
        {
            // Find the table that contains the placeholder
            var table = tables.FirstOrDefault(t =>
                t.Descendants<W.Text>().Any(txt => txt.Text.Contains(group.Key)));

            if (table == null) continue;

            // Find the row with the placeholder and remove it
            var placeholderRow = table.Descendants<W.TableRow>()
                .FirstOrDefault(r => r.InnerText.Contains(group.Key));

            if (placeholderRow != null)
            {
                table.RemoveChild(placeholderRow);
            }

            // Insert new rows
            foreach (var item in group.Items)
            {
                var newRow = new W.TableRow();
                newRow.Append(CreateCell(item.Name));
                newRow.Append(CreateCell(item.Quantity.ToString()));
                table.Append(newRow);
            }
        }
    }

    private W.TableCell CreateCell(string text)
    {
        return new W.TableCell(
            new W.Paragraph(
                new W.Run(
                    new W.Text(text)
                )
            )
        );
    }
}

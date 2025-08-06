using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using System.IO.Compression;
using WordReport.Models;
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
        var products = new List<Product>
        {
            new Product
            {
                ProductId = 1,
                NameEn = "Laptop",
                NameAr = "حاسوب محمول",
                Description = "A powerful laptop for work.",
                Category = "Electronics",
                Price = 1500,
                PriceAfterDiscount = 1200,
                QuantityInStock = 10,
                ExpiryDate = null
            },
            new Product
            {
                ProductId = 2,
                NameEn = "Smartphone",
                NameAr = "هاتف ذكي",
                Description = "Latest smartphone model.",
                Category = "Electronics",
                Price = 900,
                PriceAfterDiscount = 800,
                QuantityInStock = 25,
                ExpiryDate = null
            },
            new Product
            {
                ProductId = 3,
                NameEn = "Milk",
                NameAr = "حليب",
                Description = "Organic milk 1L",
                Category = "Groceries",
                Price = 2.5m,
                PriceAfterDiscount = 2.0m,
                QuantityInStock = 100,
                ExpiryDate = DateTime.Now.AddDays(10)
            }
        };

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
                AddTableRows(wordDoc.MainDocumentPart.Document.Body, products);

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
            if (text.Ancestors<W.Table>().Any())
                continue;

            foreach (var key in placeholders.Keys)
            {
                if (text.Text.Contains(key))
                {
                    text.Text = text.Text.Replace(key, placeholders[key], StringComparison.OrdinalIgnoreCase);
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
        var productType = typeof(Product);
        var props = propertyNames
            .Select(name => productType
                .GetProperties()
                .FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)))
            .Where(p => p != null)
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
                var value = prop.GetValue(item);
                row.Append(CreateCell(value?.ToString()));
            }
            table.Append(row);
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

using System.IO.Compression;
using System.Xml.Linq;
using WordReport.ViewModels;

namespace WordReport.Services;

public class DirectWordXmlService
{
    public byte[] GenerateDocumentManual(DocumentViewModel model, string templatePath)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(tempDir);

        // Unzip .docx (which is a zip archive)
        ZipFile.ExtractToDirectory(templatePath, tempDir);

        var placeholders = model.GetType()
        .GetProperties()
        .Where(p => p.PropertyType == typeof(string) || p.PropertyType.IsValueType)
        .ToDictionary(p => p.Name, p => p.GetValue(model)?.ToString() ?? "");

        // Replace in body
        ReplaceInXml(Path.Combine(tempDir, "word", "document.xml"), placeholders);

        // Replace in all headers
        foreach (var headerPath in Directory.GetFiles(Path.Combine(tempDir, "word"), "header*.xml"))
        {
            ReplaceInXml(headerPath, placeholders);
        }

        // Replace in all footers
        foreach (var footerPath in Directory.GetFiles(Path.Combine(tempDir, "word"), "footer*.xml"))
        {
            ReplaceInXml(footerPath, placeholders);
        }

        string outputDocxPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        ZipFile.CreateFromDirectory(tempDir, outputDocxPath);

        byte[] result = File.ReadAllBytes(outputDocxPath);

        Directory.Delete(tempDir, true);
        File.Delete(outputDocxPath);

        return result;
    }


    private void ReplaceInXml(string xmlPath, Dictionary<string, string> placeholders)
    {
        if (!File.Exists(xmlPath)) return;

        string content = File.ReadAllText(xmlPath);

        foreach (var pair in placeholders)
        {
            content = content.Replace(pair.Key, pair.Value);
        }

        File.WriteAllText(xmlPath, content);
    }


    public string ReadWordAsXml(string templatePath)
    {
        using (ZipArchive archive = ZipFile.OpenRead(templatePath))
        {
            var documentEntry = archive.GetEntry("word/document.xml");
            if (documentEntry == null) throw new FileNotFoundException("document.xml not found in the .docx file.");

            using (var reader = new StreamReader(documentEntry.Open()))
            {
                return reader.ReadToEnd();
            }
        }
    }


    public string ReplacePlaceholdersInXml(string xml, Dictionary<string, string> placeholders)
    {
        var doc = XDocument.Parse(xml);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        foreach (var textNode in doc.Descendants(w + "t"))
        {
            foreach (var key in placeholders.Keys)
            {
                if (textNode.Value.Contains(key))
                {
                    textNode.Value = textNode.Value.Replace(key, placeholders[key]);
                }
            }
        }

        return doc.ToString();
    }


    public void SaveXmlAsDocx(string originalDocxPath, string modifiedXml, string outputPath)
    {
        File.Copy(originalDocxPath, outputPath, true);
        using (var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("word/document.xml");
            entry.Delete();
            var newEntry = archive.CreateEntry("word/document.xml");

            using (var writer = new StreamWriter(newEntry.Open()))
            {
                writer.Write(modifiedXml);
            }
        }
    }
}

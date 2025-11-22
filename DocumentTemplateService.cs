using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class DocumentTemplateService
{
    public byte[] FillTemplate(string templatePath, Dictionary<string, string> fields)
    {
        using var mem = new MemoryStream();

        using (var original = File.OpenRead(templatePath))
            original.CopyTo(mem);

        mem.Position = 0;

        using (var doc = WordprocessingDocument.Open(mem, true))
        {
            var paragraphs = doc.MainDocumentPart.Document.Descendants<Paragraph>();

            foreach (var p in paragraphs)
            {
                var textNodes = p.Descendants<Text>().ToList();
                if (!textNodes.Any()) continue;

                string combinedText = string.Concat(textNodes.Select(t => t.Text));

                bool replaced = false;
                foreach (var pair in fields)
                {
                    string marker = $"{{{{{pair.Key}}}}}";
                    if (combinedText.Contains(marker))
                    {
                        combinedText = combinedText.Replace(marker, pair.Value);
                        replaced = true;
                    }
                }

                if (replaced)
                {
                    int pos = 0;

                    for (int i = 0; i < textNodes.Count; i++)
                    {
                        var t = textNodes[i];

                        if (i == textNodes.Count - 1)
                        {
                            // последний Text получает остаток строки
                            t.Text = combinedText.Substring(pos);
                        }
                        else
                        {
                            int len = t.Text.Length;
                            t.Text = combinedText.Substring(pos, len);
                            pos += len;
                        }
                    }
                }
            }

            doc.MainDocumentPart.Document.Save();
        }

        return mem.ToArray();
    }
}

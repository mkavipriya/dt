using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public static class WordTranslator
{
    public static async Task RunAsync()
    {
        string originalPath = @"C:\KAVI\translator\SampleDoc.docx";
        string translatedPath = @"C:\KAVI\translator\SampleTranslatedDoc.docx";

        File.Copy(originalPath, translatedPath, true);

        using (WordprocessingDocument doc = WordprocessingDocument.Open(translatedPath, true))
        {
            var tasks = new List<Task>();

            tasks.AddRange(TranslateParagraphs(doc.MainDocumentPart.Document.Body.Descendants<Paragraph>()));
            tasks.AddRange(TranslateParagraphs(doc.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Paragraph>())));
            tasks.AddRange(TranslateParagraphs(doc.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Paragraph>())));
            tasks.AddRange(TranslateParagraphs(doc.MainDocumentPart.FootnotesPart?.Footnotes.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>()));
            tasks.AddRange(TranslateParagraphs(doc.MainDocumentPart.EndnotesPart?.Endnotes.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>()));
            tasks.AddRange(TranslateComments(doc.MainDocumentPart.WordprocessingCommentsPart));
            tasks.AddRange(TranslateTextBoxes(doc));
            tasks.AddRange(TranslateShapes(doc));
            tasks.AddRange(TranslateFields(doc));
            tasks.AddRange(TranslateHyperlinks(doc));
            tasks.AddRange(TranslateBookmarks(doc));
            tasks.AddRange(TranslateContentControls(doc));
            tasks.AddRange(TranslateSmartArt(doc));

            await Task.WhenAll(tasks);
            doc.MainDocumentPart.Document.Save();
        }

        Console.WriteLine("✅ Translated document saved to: " + translatedPath);

        static IEnumerable<Task> TranslateParagraphs(IEnumerable<Paragraph> paragraphs)
    => paragraphs.Select(p => TranslateParagraphPreservingFormatting(p));

        static IEnumerable<Task> TranslateComments(WordprocessingCommentsPart commentsPart)
        {
            if (commentsPart == null) return Enumerable.Empty<Task>();
            return commentsPart.Comments.Descendants<Comment>().SelectMany(c => TranslateParagraphs(c.Descendants<Paragraph>()));
        }

        static IEnumerable<Task> TranslateTextBoxes(WordprocessingDocument doc)
            => TranslateParagraphs(doc.MainDocumentPart.Document.Body.Descendants<SdtBlock>().Where(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val?.Value?.Contains("TextBox") ?? false).SelectMany(sdt => sdt.Descendants<Paragraph>()));

        static IEnumerable<Task> TranslateShapes(WordprocessingDocument doc)
            => TranslateParagraphs(doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().Where(p => p.Ancestors<Drawing>().Any()));

        static IEnumerable<Task> TranslateFields(WordprocessingDocument doc)
            => TranslateParagraphs(doc.MainDocumentPart.Document.Body.Descendants<SimpleField>().SelectMany(f => f.Descendants<Paragraph>()));

        static IEnumerable<Task> TranslateHyperlinks(WordprocessingDocument doc)
            => TranslateParagraphs(doc.MainDocumentPart.HyperlinkRelationships.SelectMany(h => doc.MainDocumentPart.Document.Body.Descendants<Hyperlink>().Where(link => link.Id == h.Id).SelectMany(link => link.Descendants<Paragraph>())));

        static IEnumerable<Task> TranslateBookmarks(WordprocessingDocument doc)
            => TranslateParagraphs(doc.MainDocumentPart.Document.Body.Descendants<BookmarkStart>().SelectMany(b => b.Parent.Descendants<Paragraph>()));

        static IEnumerable<Task> TranslateContentControls(WordprocessingDocument doc)
            => TranslateParagraphs(doc.MainDocumentPart.Document.Body.Descendants<SdtBlock>().SelectMany(sdt => sdt.Descendants<Paragraph>()));

        static IEnumerable<Task> TranslateSmartArt(WordprocessingDocument doc)
        {
            var tasks = new List<Task>();

            foreach (var diagramPart in doc.MainDocumentPart.DiagramDataParts)
            {
                var xmlDoc = new System.Xml.XmlDocument();
                using (var stream = diagramPart.GetStream())
                {
                    xmlDoc.Load(stream);
                }

                var textNodes = xmlDoc.GetElementsByTagName("a:t"); // SmartArt text nodes
                foreach (System.Xml.XmlNode node in textNodes)
                {
                    string originalText = node.InnerText;
                    if (string.IsNullOrWhiteSpace(originalText)) continue;

                    tasks.Add(Task.Run(async () =>
                    {
                        string translated = await TranslateText(originalText);
                        if (!string.IsNullOrWhiteSpace(translated) && !translated.StartsWith("[Translation Error]"))
                        {
                            node.InnerText = translated;
                            using var outStream = diagramPart.GetStream(FileMode.Create, FileAccess.Write);
                            xmlDoc.Save(outStream);
                        }
                    }));
                }
            }
            static IEnumerable<Task> TranslateFootnoteReferences(WordprocessingDocument doc)
            {
                var footnotesPart = doc.MainDocumentPart.FootnotesPart;
                if (footnotesPart == null) return Enumerable.Empty<Task>();

                var footnoteIds = doc.MainDocumentPart.Document.Body
                    .Descendants<FootnoteReference>()
                    .Select(r => r.Id?.Value.ToString()) // ✅ Convert long? to string
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Distinct();

                var tasks = new List<Task>();

                foreach (var id in footnoteIds)
                {
                    if (!long.TryParse(id, out long parsedId)) continue;

                    var paragraphs = footnotesPart.Footnotes.Elements<Footnote>()
                        .Where(f => f.Id?.Value == parsedId) // ✅ Compare long? to long
                        .SelectMany(f => f.Descendants<Paragraph>());

                    tasks.AddRange(TranslateParagraphs(paragraphs));
                }

                return tasks;
            }





            return tasks;
        }

    }

    static async Task TranslateParagraphPreservingFormatting(Paragraph paragraph)
    {
        var runs = paragraph.Elements<Run>().ToList();
        if (runs.Count == 0) return;

        var sentenceGroups = GroupRunsBySentence(runs);

        foreach (var runGroup in sentenceGroups)
        {
            string originalSentence = GetTextFromRuns(runGroup);
            if (string.IsNullOrWhiteSpace(originalSentence)) continue;

            string translatedSentence = await TranslateText(originalSentence);
            if (string.IsNullOrEmpty(translatedSentence) || translatedSentence.StartsWith("[Translation Error]"))
                continue;

            ReplaceRunTextPreservingFormatting(runGroup, translatedSentence);
        }
    }

    static List<List<Run>> GroupRunsBySentence(List<Run> runs)
    {
        var result = new List<List<Run>>();
        var currentGroup = new List<Run>();
        var sentenceEndRegex = new Regex(@"(?<=[.!?])\s+|(?<=\n)");

        foreach (var run in runs)
        {
            currentGroup.Add(run);
            var textElement = run.GetFirstChild<Text>();
            if (textElement != null && sentenceEndRegex.IsMatch(textElement.Text))
            {
                result.Add(new List<Run>(currentGroup));
                currentGroup.Clear();
            }
        }

        if (currentGroup.Count > 0)
            result.Add(currentGroup);

        return result;
    }

    static string GetTextFromRuns(List<Run> runs)
    {
        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            var textElement = run.GetFirstChild<Text>();
            if (textElement != null)
                sb.Append(textElement.Text);
        }
        return sb.ToString();
    }

    static void ReplaceRunTextPreservingFormatting(List<Run> runs, string translatedText)
    {
        if (runs.Count == 0 || string.IsNullOrEmpty(translatedText)) return;

        string originalText = GetTextFromRuns(runs);
        int originalLength = originalText.Length;
        int translatedLength = translatedText.Length;
        int currentIndex = 0;

        // Normalize whitespace in translated text
        translatedText = Regex.Replace(translatedText, @"\s+", " ").Trim();

        foreach (var run in runs)
        {
            var textElement = run.GetFirstChild<Text>();
            if (textElement == null) continue;

            int originalChunkLength = textElement.Text.Length;

            // Calculate proportional chunk length
            int translatedChunkLength = (int)Math.Round((double)originalChunkLength / originalLength * translatedLength);

            // Ensure we don't exceed bounds
            if (currentIndex + translatedChunkLength > translatedText.Length)
                translatedChunkLength = translatedText.Length - currentIndex;

            string chunk = translatedText.Substring(currentIndex, translatedChunkLength);

            // Insert space if needed between chunks
            if (currentIndex > 0 && !char.IsWhiteSpace(translatedText[currentIndex - 1]) && !chunk.StartsWith(" "))
                chunk = " " + chunk;

            textElement.Text = chunk;
            currentIndex += translatedChunkLength;
        }

        // If any remaining text wasn't assigned, append it to the last run
        if (currentIndex < translatedText.Length)
        {
            var lastRun = runs.LastOrDefault();
            var lastText = lastRun?.GetFirstChild<Text>();
            if (lastText != null)
                lastText.Text += translatedText.Substring(currentIndex);
        }
    }


    static async Task<string> TranslateText(string inputText)
    {
        string apiUrl = "http://127.0.0.1:5000/translate";

        var payload = new { text = inputText };
        string json = JsonSerializer.Serialize(payload);
        var content = new StringContent(json, Encoding.UTF8, "application/json");

        using HttpClient client = new HttpClient();
        try
        {
            HttpResponseMessage response = await client.PostAsync(apiUrl, content);
            response.EnsureSuccessStatusCode();

            string responseBody = await response.Content.ReadAsStringAsync();
            using JsonDocument doc = JsonDocument.Parse(responseBody);
            return doc.RootElement.GetProperty("translated").GetString();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️ Translation failed for: \"{inputText}\" — {ex.Message}");
            return $"[Translation Error] {ex.Message}";
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.IO;
using System.Text;

namespace BlazorOffice.Services
{
    public class WordService
    {
        // Extracts text with basic formatting (Bold, Italic, Underline) as HTML
        public async Task<string> ReadDocxAsync(Stream stream)
        {
            using var memStream = new MemoryStream();
            await stream.CopyToAsync(memStream);
            memStream.Position = 0;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memStream, false))
            {
                var body = wordDoc.MainDocumentPart?.Document.Body;
                if (body == null) return string.Empty;

                // Priority 1: Check for AltChunk (HTML content we saved previously)
                var altChunk = body.Elements<AltChunk>().FirstOrDefault();
                if (altChunk != null)
                {
                    var part = wordDoc.MainDocumentPart.GetPartById(altChunk.Id);
                    // Use a stream reader to read the raw HTML content from the chunk part
                    using var streamReader = new StreamReader(part.GetStream());
                    return await streamReader.ReadToEndAsync();
                }

                // Priority 2: Parse native OpenXML Paragraphs (Standard Word Docs)
                StringBuilder htmlBuilder = new StringBuilder();
                bool inList = false;

                // Iterate over Elements to preserve structure (Tables vs Paragraphs)
                foreach (var element in body.Elements())
                {
                    if (element is Paragraph p)
                    {
                        string pHtml = ParseParagraph(p, out bool isListItem);
                        
                        // List State Machine
                        if (isListItem && !inList)
                        {
                            htmlBuilder.Append("<ul>");
                            inList = true;
                        }
                        else if (!isListItem && inList)
                        {
                            htmlBuilder.Append("</ul>");
                            inList = false;
                        }

                        htmlBuilder.Append(pHtml);
                    }
                    else if (element is Table t)
                    {
                        if (inList)
                        {
                            htmlBuilder.Append("</ul>");
                            inList = false;
                        }
                        htmlBuilder.Append(ParseTable(t));
                    }
                }

                if (inList) htmlBuilder.Append("</ul>");

                return htmlBuilder.ToString();
            }
        }

        private string ParseTable(Table table)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<table class=\"table table-bordered\">");

            foreach (var row in table.Elements<TableRow>())
            {
                sb.Append("<tr>");
                foreach (var cell in row.Elements<TableCell>())
                {
                    sb.Append("<td>");
                    // For now, we only grab Paragraphs directly in cells.
                    // This handles 90% of resumes. Recursion would be needed for nested tables.
                    foreach (var p in cell.Elements<Paragraph>())
                    {
                        sb.Append(ParseParagraph(p, out _));
                    }
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }

            sb.Append("</table>");
            return sb.ToString();
        }

        private string ParseParagraph(Paragraph p, out bool isListItem)
        {
            StringBuilder htmlBuilder = new StringBuilder();
            var pPr = p.ParagraphProperties;
            List<string> styles = new List<string>();
            isListItem = false;

            if (pPr != null)
            {
                // Alignment
                if (pPr.Justification?.Val != null)
                {
                    var align = pPr.Justification.Val.Value;
                    if (align == JustificationValues.Center) styles.Add("text-align: center");
                    else if (align == JustificationValues.Right) styles.Add("text-align: right");
                    else if (align == JustificationValues.Both) styles.Add("text-align: justify");
                }

                // Indentation
                if (pPr.Indentation?.Left != null && int.TryParse(pPr.Indentation.Left.Value, out int leftDxa))
                {
                        int pixels = leftDxa / 15; 
                        styles.Add($"padding-left: {pixels}px");
                }

                // Lists
                if (pPr.NumberingProperties != null) isListItem = true;

                // Borders
                if (pPr.ParagraphBorders != null)
                {
                    if (pPr.ParagraphBorders.BottomBorder?.Val != null && pPr.ParagraphBorders.BottomBorder.Val != BorderValues.None)
                    {
                        styles.Add("border-bottom: 1px solid black");
                        styles.Add("padding-bottom: 5px"); 
                    }
                    if (pPr.ParagraphBorders.TopBorder?.Val != null && pPr.ParagraphBorders.TopBorder.Val != BorderValues.None)
                    {
                        styles.Add("border-top: 1px solid black");
                        styles.Add("padding-top: 5px"); 
                    }
                }
            }

            string styleAttr = styles.Count > 0 ? $" style=\"{string.Join("; ", styles)}\"" : "";
            string tag = isListItem ? "li" : "p";
            
            htmlBuilder.Append($"<{tag}{styleAttr}>");

            // Use Descendants<Run> to catch runs inside Hyperlinks, SmartTags, etc.
            foreach (var run in p.Descendants<Run>())
            {
                // Check properties
                var props = run.RunProperties;
                bool isBold = props?.Bold != null && (props.Bold.Val == null || props.Bold.Val.Value);
                bool isItalic = props?.Italic != null && (props.Italic.Val == null || props.Italic.Val.Value);
                bool isUnderline = props?.Underline != null && (props.Underline.Val == null || props.Underline.Val != UnderlineValues.None);

                if (isBold) htmlBuilder.Append("<strong>");
                if (isItalic) htmlBuilder.Append("<em>");
                if (isUnderline) htmlBuilder.Append("<u>");

                // Iterate children to handle Mixed Content (Text + Breaks + Tabs)
                foreach (var child in run.Elements())
                {
                    if (child is Text t)
                    {
                        htmlBuilder.Append(System.Web.HttpUtility.HtmlEncode(t.Text));
                    }
                    else if (child is Break) // Handles <w:br/> (Shift+Enter)
                    {
                        htmlBuilder.Append("<br/>");
                    }
                    else if (child is TabChar) // Handles <w:tab/>
                    {
                        htmlBuilder.Append("&emsp;"); // Tab = ~4 spaces
                    }
                }

                if (isUnderline) htmlBuilder.Append("</u>");
                if (isItalic) htmlBuilder.Append("</em>");
                if (isBold) htmlBuilder.Append("</strong>");
            }
            
            htmlBuilder.Append($"</{tag}>");
            return htmlBuilder.ToString();
        }

        // Creates a native DOCX from HTML content using HtmlToOpenXml
        // This ensures compatibility with Google Docs, LibreOffice, etc. (No AltChunk Frankenstein)
        public async Task<byte[]> CreateDocxFromHtmlAsync(string htmlContent)
        {
            using var memStream = new MemoryStream();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // THE FIX: "True" XML Conversion
                var converter = new HtmlConverter(mainPart);
                var compositeElements = converter.Parse(htmlContent);
                
                body.Append(compositeElements);

                mainPart.Document.Save();
            }

            return memStream.ToArray();
        }

        // Creates a simple DOCX from raw text
        public async Task<byte[]> CreateDocxAsync(string textContent)
        {
            using var memStream = new MemoryStream();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                // Add the main document part
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Split text by newlines to create separate paragraphs
                var paragraphs = textContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                foreach (var line in paragraphs)
                {
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text(line));
                }

                mainPart.Document.Save();
            }

            return memStream.ToArray();
        }
    }
}

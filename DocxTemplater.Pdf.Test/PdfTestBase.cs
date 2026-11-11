using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Pdf.Test
{
    internal abstract class PdfTestBase
    {
        // A4 portrait with 2.5 cm left/right and 2 cm top/bottom margins
        protected const long A4WidthTwips = 11906;
        protected const long A4HeightTwips = 16838;
        protected const long MarginLeftRight = 1417;
        protected const long MarginTopBottom = 1134;

        protected static SectionProperties A4Section(params OpenXmlElement[] extra)
        {
            var section = new SectionProperties();
            foreach (var element in extra)
            {
                section.AddChild(element.CloneNode(true));
            }
            section.AddChild(new PageSize { Width = (uint)A4WidthTwips, Height = (uint)A4HeightTwips });
            section.AddChild(new PageMargin { Top = (int)MarginTopBottom, Right = (uint)MarginLeftRight, Bottom = (int)MarginTopBottom, Left = (uint)MarginLeftRight, Header = 708U, Footer = 708U, Gutter = 0U });
            return section;
        }

        protected static Paragraph P(string text)
        {
            return new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        protected static Paragraph PWithSection(string text, SectionProperties section)
        {
            var paragraph = P(text);
            paragraph.PrependChild(new ParagraphProperties { SectionProperties = section });
            return paragraph;
        }

        protected static long ContentWidthEmu => (A4WidthTwips - (2 * MarginLeftRight)) * 635;

        protected static long ContentHeightEmu => (A4HeightTwips - (2 * MarginTopBottom)) * 635;

        protected static (Body Body, WordprocessingDocument Document) Render(
            Body templateBody,
            object model,
            IPdfRenderer renderer = null,
            PdfFormatterConfiguration configuration = null,
            Action<WordprocessingDocument> setup = null)
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(templateBody);
                setup?.Invoke(wpDocument);
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new PdfFormatter(renderer ?? new FakePdfRenderer(), configuration));
            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            return (document.MainDocumentPart.Document.Body, document);
        }

        /// <summary>
        /// The sections of the body in order, each with a short description of its content and its page size.
        /// Paragraphs are described by their text, page pictures as "[pdf]" (inline) or "[plan]" (anchored),
        /// overlays as "[overlay]", empty paragraphs as "".
        /// </summary>
        protected static List<(List<string> Content, SectionProperties Section)> GetSections(Body body)
        {
            var result = new List<(List<string>, SectionProperties)>();
            var current = new List<string>();
            foreach (var element in body.ChildElements)
            {
                if (element is SectionProperties bodySection)
                {
                    result.Add((current, bodySection));
                    current = new List<string>();
                    continue;
                }
                current.Add(Describe(element));
                if (element is Paragraph { ParagraphProperties.SectionProperties: { } section })
                {
                    result.Add((current, section));
                    current = new List<string>();
                }
            }
            Assert.That(current, Is.Empty, "content after the last section properties");
            return result;
        }

        protected static string Describe(OpenXmlElement element)
        {
            if (element is Table)
            {
                return "[table]";
            }
            var parts = new List<string>();
            foreach (var drawing in element.Descendants<Drawing>())
            {
                if (drawing.Inline != null)
                {
                    parts.Add("[pdf]");
                }
                else if (drawing.Anchor?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().Any() == true)
                {
                    parts.Add("[plan]");
                }
                else
                {
                    parts.Add("[overlay]");
                }
            }
            var text = string.Concat(element.Descendants<Text>().Where(x => x.GetFirstAncestor<Drawing>() == null).Select(x => x.Text));
            if (text.Length > 0)
            {
                parts.Insert(0, text);
            }
            return string.Join(" ", parts);
        }

        protected static string PageSizeName(SectionProperties section)
        {
            var size = section.GetFirstChild<PageSize>();
            var width = (long)size.Width.Value;
            var height = (long)size.Height.Value;
            var name = (width, height) switch
            {
                (11906, 16838) => "A4",
                (16838, 11906) => "A4L",
                (16838, 23811) => "A3",
                (23811, 16838) => "A3L",
                _ => $"{width}x{height}",
            };
            return name;
        }

        protected static List<string> DescribeSections(Body body)
        {
            return GetSections(body).Select(x => $"{PageSizeName(x.Section)}: {string.Join(" | ", x.Content)}").ToList();
        }

        protected static List<DW.Extent> Extents(OpenXmlElement root)
        {
            return root.Descendants<DW.Extent>().ToList();
        }
    }
}

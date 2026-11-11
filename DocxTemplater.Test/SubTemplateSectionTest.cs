using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// Sub-templates (<c>:T</c>) whose document has section breaks of its own, e.g. plan pages of the PDF formatter.
    /// The host and the inserted document both use the relationship id rId9 for different footers.
    /// </summary>
    internal class SubTemplateSectionTest
    {
        [Test]
        public void InsertedDocumentWithSectionBreaksKeepsHostSectionAroundIt()
        {
            var inserted = CreateInsertedDocument(P("Intro", OwnSection()), P("Plan", PlanSection()), P("After plans"));

            var sections = Render(new Body(P("Cover"), P("{{ds}:T('ds.Anejo')}"), P("", HostSection()), P("Back cover"), HostBodySection()), inserted);

            Assert.That(sections, Is.EqualTo([
                "A4 titlePg start=1 [first=host first, default=host default]: Cover | Intro",
                "A3L [default=]: Plan",
                "A4 [first=host first, default=host default]: After plans | ",
                "A4 [default=host default]: Back cover",
            ]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void InsertedDocumentStartingWithAPlan()
        {
            var inserted = CreateInsertedDocument(P("Plan", PlanSection()), P("After plans"));

            var sections = Render(new Body(P("Cover"), P("{{ds}:T('ds.Anejo')}"), P("", HostSection()), P("Back cover"), HostBodySection()), inserted);

            Assert.That(sections, Is.EqualTo([
                "A4 titlePg start=1 [first=host first, default=host default]: Cover | ",
                "A3L [default=]: Plan",
                "A4 [first=host first, default=host default]: After plans | ",
                "A4 [default=host default]: Back cover",
            ]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void InsertedDocumentStartingWithAPlanAfterTextInThePlaceholderParagraph()
        {
            var inserted = CreateInsertedDocument(P("Plan", PlanSection()), P("After plans"));

            var sections = Render(new Body(P("Cover {{ds}:T('ds.Anejo')}"), P("", HostSection()), P("Back cover"), HostBodySection()), inserted);

            Assert.That(sections, Is.EqualTo([
                "A4 titlePg start=1 [first=host first, default=host default]: Cover ",
                "A3L [default=]: Plan",
                "A4 [first=host first, default=host default]: After plans | ",
                "A4 [default=host default]: Back cover",
            ]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void PlaceholderParagraphEndsTheHostSection()
        {
            var inserted = CreateInsertedDocument(P("Intro", OwnSection()), P("Plan", PlanSection()), P("After plans"));

            var sections = Render(new Body(P("Cover"), P("{{ds}:T('ds.Anejo')}", HostSection()), P("Back cover"), HostBodySection()), inserted);

            Assert.That(sections, Is.EqualTo([
                "A4 titlePg start=1 [first=host first, default=host default]: Cover | Intro",
                "A3L [default=]: Plan",
                "A4 [first=host first, default=host default]: After plans",
                "A4 [default=host default]: Back cover",
            ]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void PlaceholderParagraphEndingTheSectionWithoutBreaksInTheInsertedDocument()
        {
            var inserted = CreateInsertedDocument(P("Intro"), P("More"));

            var sections = Render(new Body(P("Cover"), P("{{ds}:T('ds.Anejo')}", HostSection()), P("Back cover"), HostBodySection()), inserted);

            Assert.That(sections, Is.EqualTo([
                "A4 titlePg start=1 [first=host first, default=host default]: Cover | Intro | More",
                "A4 [default=host default]: Back cover",
            ]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void InsertedDocumentWithoutSectionBreaksTakesTheHostSection()
        {
            var inserted = CreateInsertedDocument(P("Intro"), P("More"));

            var sections = Render(new Body(P("Cover"), P("{{ds}:T('ds.Anejo')}"), P("", HostSection()), P("Back cover"), HostBodySection()), inserted);

            Assert.That(sections, Is.EqualTo([
                "A4 titlePg start=1 [first=host first, default=host default]: Cover | Intro | More | ",
                "A4 [default=host default]: Back cover",
            ]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void SectionBreaksAreDroppedInATableCell()
        {
            var inserted = CreateInsertedDocument(P("Intro", OwnSection()), P("Plan", PlanSection()), P("After plans"));

            var sections = Render(new Body(P("Cover"), new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "5000" }), new TableRow(new TableCell(P("{{ds}:T('ds.Anejo')}")))), P("Back cover"), HostBodySection()), inserted);

            Assert.That(sections, Is.EqualTo(["A4 [default=host default]: Cover | [table] | Back cover"]), string.Join(Environment.NewLine, sections));
        }

        private static List<string> Render(Body hostBody, byte[] inserted)
        {
            using var memStream = new MemoryStream();
            using (var host = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = host.AddMainDocumentPart();
                mainPart.Document = new Document(hostBody);
                mainPart.AddNewPart<FooterPart>("rId9").Footer = new Footer(P("host first"));
                mainPart.AddNewPart<FooterPart>("rId10").Footer = new Footer(P("host default"));
            }
            memStream.Position = 0;
            var template = new DocxTemplate(memStream);
            template.BindModel("ds", new { Anejo = inserted });
            var result = template.Process();
            template.Validate();
            result.Position = 0;
            using var document = WordprocessingDocument.Open(result, false);
            return DescribeSections(document.MainDocumentPart);
        }

        private static byte[] CreateInsertedDocument(params Paragraph[] paragraphs)
        {
            using var memStream = new MemoryStream();
            using (var document = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(paragraphs.Append<OpenXmlElement>(OwnSection())));
                mainPart.AddNewPart<FooterPart>("rId9").Footer = new Footer(P("anejo"));
                mainPart.AddNewPart<FooterPart>("rId12").Footer = new Footer(new Paragraph());
            }
            return memStream.ToArray();
        }

        private static List<string> DescribeSections(MainDocumentPart mainPart)
        {
            var result = new List<string>();
            var content = new List<string>();
            foreach (var element in mainPart.Document.Body.ChildElements)
            {
                if (element is Paragraph or Table)
                {
                    content.Add(element is Table ? "[table]" : element.InnerText);
                }
                var section = element as SectionProperties ?? (element as Paragraph)?.ParagraphProperties?.SectionProperties;
                if (section != null)
                {
                    result.Add($"{Describe(section, mainPart)}: {string.Join(" | ", content)}");
                    content.Clear();
                }
            }
            if (mainPart.Document.Body.Descendants<SectionProperties>().Count() != result.Count)
            {
                result.Add("section break below the body level");
            }
            return result;
        }

        private static string Describe(SectionProperties section, MainDocumentPart mainPart)
        {
            var size = section.GetFirstChild<PageSize>();
            var name = size.Width.Value > size.Height.Value ? "A3L" : "A4";
            var titlePage = section.GetFirstChild<TitlePage>() != null ? " titlePg" : string.Empty;
            var start = section.GetFirstChild<PageNumberType>()?.Start?.Value is { } value ? $" start={value}" : string.Empty;
            var footers = section.Elements<FooterReference>()
                .Select(x => $"{x.Type.InnerText}={((FooterPart)mainPart.GetPartById(x.Id)).Footer.InnerText}");
            return $"{name}{titlePage}{start} [{string.Join(", ", footers)}]";
        }

        private static Paragraph P(string text, SectionProperties section = null)
        {
            var paragraph = new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            if (section != null)
            {
                paragraph.PrependChild(new ParagraphProperties(section));
            }
            return paragraph;
        }

        private static SectionProperties A4(params OpenXmlElement[] children)
        {
            var section = new SectionProperties(children);
            section.AddChild(new PageSize { Width = 11906U, Height = 16838U });
            section.AddChild(new PageMargin { Top = 1417, Right = 1701U, Bottom = 1417, Left = 1701U, Header = 708U, Footer = 708U, Gutter = 0U });
            return section;
        }

        // the host section S holding the placeholder: different first page footer, page numbering starts at 1
        private static SectionProperties HostSection()
        {
            var section = A4(
                new FooterReference { Type = HeaderFooterValues.First, Id = "rId9" },
                new FooterReference { Type = HeaderFooterValues.Default, Id = "rId10" });
            section.AddChild(new PageNumberType { Start = 1 });
            section.AddChild(new TitlePage());
            return section;
        }

        private static SectionProperties HostBodySection()
        {
            return A4(new FooterReference { Type = HeaderFooterValues.Default, Id = "rId10" });
        }

        // the inserted document's own page setup - as the closing copies of fit:section are
        private static SectionProperties OwnSection()
        {
            return A4(new FooterReference { Type = HeaderFooterValues.Default, Id = "rId9" });
        }

        private static SectionProperties PlanSection()
        {
            var section = new SectionProperties(new FooterReference { Type = HeaderFooterValues.Default, Id = "rId12" });
            section.AddChild(new PageSize { Width = 23811U, Height = 16838U, Orient = PageOrientationValues.Landscape });
            section.AddChild(new PageMargin { Top = 0, Right = 0U, Bottom = 0, Left = 0U, Header = 0U, Footer = 0U, Gutter = 0U });
            return section;
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Pdf.Test
{
    internal class PdfFormatterBlankTest : PdfTestBase
    {
        private const string Marker = "[[PDF|%key%|%page%|%pages%]]";

        private static readonly byte[] A3Landscape = FakePdfRenderer.Pdf((FakePdfRenderer.A3Height, FakePdfRenderer.A3Width));

        private static byte[] A3LandscapePages(int count)
        {
            return FakePdfRenderer.Pdf(Enumerable.Repeat((FakePdfRenderer.A3Height, FakePdfRenderer.A3Width), count).ToArray());
        }

        private static List<Run> MarkerRuns(OpenXmlElement root)
        {
            return root.Descendants<Run>().Where(x => x.RunProperties?.Color?.Val?.Value == "FFFFFF").ToList();
        }

        [Test]
        public void NothingIsRendered()
        {
            var renderer = new FakePdfRenderer();
            var (body, document) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(content:blank)}"), P("End"), A4Section()),
                new { Pdf = A3LandscapePages(3) }, renderer, new PdfFormatterConfiguration { DisableImageCompression = true });

            Assert.That(renderer.Requests, Is.Empty);
            Assert.That(renderer.GetPageSizesCalls, Is.EqualTo(1));
            Assert.That(document.MainDocumentPart.ImageParts, Is.Empty);
            Assert.That(body.Descendants<Drawing>(), Is.Empty);
            Assert.That(document.MainDocumentPart.DocumentSettingsPart?.Settings?.GetFirstChild<DoNotAutoCompressPictures>(), Is.Null, "no pictures to compress");
        }

        [Test]
        public void OnePagePerPdfPage()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(content:blank)}"), P("End"), A4Section()), new { Pdf = A3LandscapePages(3) });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro", "A3L:  |  | ", "A4: End"]));
            var pages = body.Elements<Paragraph>().Skip(1).Take(3).ToList();
            Assert.That(pages.Select(x => x.ParagraphProperties.PageBreakBefore != null), Is.EqualTo([false, true, true]));
            Assert.That(pages.Select(x => x.ParagraphProperties.SpacingBetweenLines.Line.Value), Has.All.EqualTo("20"));
        }

        [Test]
        public void PageSizesAndOrientationOfThePdfPages()
        {
            var pdf = FakePdfRenderer.Pdf((FakePdfRenderer.A4Width, FakePdfRenderer.A4Height), (FakePdfRenderer.A3Height, FakePdfRenderer.A3Width));
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(content:blank)}"), P("End"), A4Section()), new { Pdf = pdf });

            // the A4 page gets a section of its own too (same is ignored), so it can't merge with the text
            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro", "A4: ", "A3L: ", "A4: End"]));
            var sections = GetSections(body);
            Assert.That(sections[1].Section.GetFirstChild<PageMargin>().Left.Value, Is.Zero);
            Assert.That(sections[2].Section.GetFirstChild<PageSize>().Orient.Value, Is.EqualTo(PageOrientationValues.Landscape));
        }

        [Test]
        public void FitIsIgnored()
        {
            var renderer = new FakePdfRenderer();
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(content:blank, fit:page, same:inline)}"), P("End"), A4Section()),
                new { Pdf = FakePdfRenderer.A4Pages(2) }, renderer);

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro", "A4:  | ", "A4: End"]));
            Assert.That(renderer.Requests, Is.Empty);
        }

        [Test]
        public void PageSelectionAndMarker()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(content:blank, pages:'2-3', marker:'" + Marker + "')}"), P("End"), A4Section()),
                new { Pdf = new PdfFile(A3LandscapePages(5), "Data sheet", "sheet3") });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro", "A3L: [[PDF|sheet3|2|5]] | [[PDF|sheet3|3|5]]", "A4: End"]));
        }

        [Test]
        public void MarkerIsWhiteTinyTextOnOneLine()
        {
            var configuration = new PdfFormatterConfiguration { MarkerText = "PDF {name}\r\npage {page}\n{pages}" };
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf(content:blank)}"), P("End"), A4Section()), new { Pdf = new PdfFile(A3Landscape, "Sheet") },
                configuration: configuration);

            var run = MarkerRuns(body).Single();
            Assert.That(run.InnerText, Is.EqualTo("PDF Sheet page 1 1"));
            Assert.That(run.RunProperties.FontSize.Val.Value, Is.EqualTo("2"));
            Assert.That(run.RunProperties.Vanish, Is.Null, "hidden text is not exported to PDF");
            Assert.That(run.GetFirstChild<Text>().Space.Value, Is.EqualTo(SpaceProcessingModeValues.Preserve));
        }

        [Test]
        public void NoMarkerByDefault()
        {
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf(content:blank)}"), P("End"), A4Section()), new { Pdf = A3Landscape });
            Assert.That(body.Descendants<Paragraph>().First().Descendants<Run>(), Is.Empty);
        }

        [Test]
        public void KeyFromTheFileTheTemplateOrTheName()
        {
            static string RenderMarker(string arguments, object pdf)
            {
                var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf(content:blank, marker:'%key%'" + arguments + ")}"), P("End"), A4Section()), new { Pdf = pdf });
                return MarkerRuns(body).Single().InnerText;
            }

            Assert.That(RenderMarker("", new PdfFile(A3Landscape, "Name", "file-key")), Is.EqualTo("file-key"));
            Assert.That(RenderMarker(", key:'template-key'", new PdfFile(A3Landscape, "Name", "file-key")), Is.EqualTo("template-key"));
            Assert.That(RenderMarker("", new PdfFile(A3Landscape, "Name")), Is.EqualTo("Name"));
            Assert.That(RenderMarker(", label:'Label'", A3Landscape), Is.EqualTo("Label"));
        }

        [Test]
        public void OverlayOnBlankPages()
        {
            var (always, _) = Render(new Body(P("{{ds.Pdf}:pdf(content:blank, overlay:always)}"), P("End"), A4Section()),
                new { Pdf = new PdfFile(A3Landscape, "Sheet") });
            var overlay = always.Descendants<DW.Anchor>().Single();
            Assert.That(overlay.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>(), Is.Empty);
            Assert.That(overlay.Descendants<Paragraph>().Select(x => x.InnerText).Last(), Is.EqualTo("Plan: Sheet, page 1 of 1"));

            var (auto, _) = Render(new Body(P("{{ds.Pdf}:pdf(content:blank)}"), P("End"), A4Section()), new { Pdf = A3Landscape });
            Assert.That(auto.Descendants<DW.Anchor>(), Is.Empty, "true size: no overlay with overlay:auto");

            var a1Landscape = FakePdfRenderer.Pdf((2383.94, 1683.78));
            var (oversize, _) = Render(new Body(P("{{ds.Pdf}:pdf(content:blank)}"), P("End"), A4Section()), new { Pdf = a1Landscape });
            Assert.That(GetSections(oversize)[0].Section.GetFirstChild<PageSize>().Width.Value, Is.EqualTo(31680));
            Assert.That(oversize.Descendants<DW.Anchor>().Count(), Is.EqualTo(1), "scaled pages get the overlay");
        }

        [Test]
        public void BlankPagesInALoopWithTitles()
        {
            var (body, _) = Render(
                new Body(P("Sheets"), P("{{#ds.Sheets}}"), P("Sheet {{.Name}}"), P("{{.Pdf}:pdf(content:blank, marker:'%key%-%page%')}"), P("{{/ds.Sheets}}"), A4Section()),
                new
                {
                    Sheets = new[]
                    {
                        new { Name = "A", Pdf = new PdfFile(A3LandscapePages(2), "A", "a") },
                        new { Name = "B", Pdf = new PdfFile(FakePdfRenderer.A4Pages(1), "B", "b") },
                    },
                });

            var sections = DescribeSections(body);
            Assert.That(sections, Is.EqualTo(["A4: Sheets | Sheet A", "A3L: a-1 | a-2", "A4: Sheet B", "A4: b-1"]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void BlankPagesInAnInlineLoopFollowedByText()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{#ds.Pdfs}}{{.}:pdf(content:blank, marker:'x')}{{/ds.Pdfs}}"), P("End"), A4Section()),
                new { Pdfs = new[] { A3Landscape, A3Landscape, A3Landscape } });

            var sections = DescribeSections(body);
            Assert.That(sections, Is.EqualTo(["A4: Intro | ", "A3L: x | ", "A3L: x | ", "A3L: x", "A4:  | End"]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void BlankPagesAtTheEndOfTheDocumentAddNoEmptyPage()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(content:blank)}"), A4Section()), new { Pdf = A3LandscapePages(2) });
            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro", "A3L:  | "]));
        }

        [Test]
        public void MarkerOnRenderedSectionPagesOnly()
        {
            var pdf = FakePdfRenderer.Pdf((FakePdfRenderer.A4Width, FakePdfRenderer.A4Height), (FakePdfRenderer.A3Height, FakePdfRenderer.A3Width));
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(fit:section, marker:'m%page%')}"), P("End"), A4Section()), new { Pdf = pdf });

            // the A4 page stays in the text flow (same:inline) and has no marker
            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro | [pdf]", "A3L: m2 [plan]", "A4: End"]));
        }

        [Test]
        public void NotPossibleInTables()
        {
            var table = new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "2000" }), new TableRow(new TableCell(P("{{ds.Pdf}:pdf(content:blank)}"))));
            var renderer = new FakePdfRenderer();
            var exception = Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(table, P(""), A4Section()), new { Pdf = A3Landscape }, renderer));
            Assert.That(exception.InnerException.Message, Does.Contain("content:blank"));
            Assert.That(renderer.GetPageSizesCalls, Is.Zero, "fails before reading the PDF");
        }

        [Test]
        public void ContentArgument()
        {
            Assert.That(PdfArguments.Apply(PdfFormatterConfiguration.Default, ["content:blank"]).Content, Is.EqualTo(PdfContent.Blank));
            Assert.That(PdfArguments.Apply(PdfFormatterConfiguration.Default, ["content:Render"]).Content, Is.EqualTo(PdfContent.Render));
            var exception = Assert.Throws<OpenXmlTemplateException>(() => PdfArguments.Apply(PdfFormatterConfiguration.Default, ["content:none"]));
            Assert.That(exception.Message, Does.Contain("expected render, blank"));

            var configuration = PdfArguments.Apply(PdfFormatterConfiguration.Default, ["marker:'[[PDF|%key%|%page%]] 100%'", "key:'k1'"]);
            Assert.That(configuration.MarkerText, Is.EqualTo("[[PDF|{key}|{page}]] 100%"));
            Assert.That(configuration.Key, Is.EqualTo("k1"));
        }
    }
}

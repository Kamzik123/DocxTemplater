using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Model;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Pdf.Test
{
    internal class PdfFormatterFlowTest : PdfTestBase
    {
        [Test]
        public void SinglePageIsScaledIntoTheTextArea()
        {
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(1) });

            var paragraph = body.Elements<Paragraph>().Single();
            var extent = paragraph.Descendants<DW.Extent>().Single();
            Assert.That(extent.Cx.Value, Is.LessThanOrEqualTo(ContentWidthEmu));
            Assert.That(extent.Cy.Value, Is.LessThanOrEqualTo((long)(ContentHeightEmu * 0.97) + 1));
            Assert.That(extent.Cx.Value == ContentWidthEmu || Math.Abs(extent.Cy.Value - (long)(ContentHeightEmu * 0.97)) <= 1, "one side fills the area");
            Assert.That((double)extent.Cx.Value / extent.Cy.Value, Is.EqualTo(FakePdfRenderer.A4Width / FakePdfRenderer.A4Height).Within(0.001));
            Assert.That(paragraph.Descendants<Break>(), Is.Empty, "no page break before the first page");
        }

        [Test]
        public void PagesAfterTheFirstStartWithAPageBreakInsideTheirParagraph()
        {
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(3) });

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs, Has.Count.EqualTo(3));
            Assert.That(paragraphs[0].Descendants<Break>(), Is.Empty);
            foreach (var paragraph in paragraphs.Skip(1))
            {
                var runs = paragraph.Elements<Run>().ToList();
                Assert.That(runs, Has.Count.EqualTo(2));
                Assert.That(runs[0].GetFirstChild<Break>().Type.Value, Is.EqualTo(BreakValues.Page), "break before the picture");
                Assert.That(runs[1].GetFirstChild<Drawing>(), Is.Not.Null);
            }
            var spacing = paragraphs[0].ParagraphProperties.SpacingBetweenLines;
            Assert.That(spacing.LineRule.Value, Is.EqualTo(LineSpacingRuleValues.Auto));
            Assert.That(spacing.Before.Value, Is.EqualTo("0"));
            Assert.That(spacing.After.Value, Is.EqualTo("0"));
        }

        [Test]
        public void PageBreakBeforeTheFirstPage()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Pdf}:pdf(pb:true)}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(1) });

            Assert.That(body.Elements<Paragraph>().Last().Descendants<Break>().Single().Type.Value, Is.EqualTo(BreakValues.Page));
        }

        [Test]
        public void ExactLineSpacingOfTheTemplateIsReplaced()
        {
            var template = P("{{ds.Pdf}:pdf}");
            template.PrependChild(new ParagraphProperties(
                new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Exact },
                new Indentation { Left = "720" },
                new Justification { Val = JustificationValues.Center }));
            var (body, _) = Render(new Body(template, A4Section()), new { Pdf = FakePdfRenderer.A4Pages(1) });

            var properties = body.Elements<Paragraph>().Single().ParagraphProperties;
            Assert.That(properties.SpacingBetweenLines.LineRule.Value, Is.EqualTo(LineSpacingRuleValues.Auto));
            Assert.That(properties.Indentation.Left.Value, Is.EqualTo("0"));
            Assert.That(properties.Justification.Val.Value, Is.EqualTo(JustificationValues.Center), "alignment is kept");
        }

        [Test]
        public void PageSelectionWithCommasInQuotes()
        {
            var renderer = new FakePdfRenderer();
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf(pages:'1-2,4', dpi:200)}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(5) }, renderer);

            Assert.That(body.Descendants<Drawing>().Count(), Is.EqualTo(3));
            Assert.That(renderer.Requests.Single().PageIndices, Is.EqualTo([0, 1, 3]));
            Assert.That(renderer.Requests.Single().Dpi, Is.EqualTo(200));
        }

        [Test]
        public void OpenEndedPageRange()
        {
            var renderer = new FakePdfRenderer();
            Render(new Body(P("{{ds.Pdf}:pdf(pages:'3-')}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(5) }, renderer);

            Assert.That(renderer.Requests.Single().PageIndices, Is.EqualTo([2, 3, 4]));
        }

        [Test]
        public void PageOutOfRangeThrows()
        {
            var exception = Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf(pages:'2-4')}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(3) }));
            Assert.That(exception.InnerException.Message, Does.Contain("3 page(s)"));
        }

        [Test]
        public void DpiDoesNotChangeTheDisplayedSize()
        {
            var pdf = FakePdfRenderer.A4Pages(1);
            var extents = new[] { 72, 150, 300 }
                .Select(dpi => Render(new Body(P($"{{{{ds.Pdf}}:pdf(dpi:{dpi})}}"), A4Section()), new { Pdf = pdf }).Body.Descendants<DW.Extent>().Single())
                .Select(x => (x.Cx.Value, x.Cy.Value))
                .Distinct()
                .ToList();
            Assert.That(extents, Has.Count.EqualTo(1));
        }

        [Test]
        public void TextBeforeAndAfterThePlaceholderStaysInOwnParagraphs()
        {
            var (body, _) = Render(new Body(P("Before {{ds.Pdf}:pdf} After"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(2) });

            Assert.That(body.Elements<Paragraph>().Select(Describe), Is.EqualTo(["Before ", "[pdf]", "[pdf]", " After"]));
        }

        [Test]
        public void FitWidthAndExplicitWidth()
        {
            var landscape = FakePdfRenderer.Pdf((FakePdfRenderer.A4Height, FakePdfRenderer.A4Width));
            var width = Render(new Body(P("{{ds.Pdf}:pdf(fit:width)}"), A4Section()), new { Pdf = landscape }).Body.Descendants<DW.Extent>().Single();
            Assert.That(width.Cx.Value, Is.EqualTo(ContentWidthEmu));

            var explicitWidth = Render(new Body(P("{{ds.Pdf}:pdf(w:5cm)}"), A4Section()), new { Pdf = landscape }).Body.Descendants<DW.Extent>().Single();
            Assert.That(explicitWidth.Cx.Value, Is.EqualTo(1800000));

            var box = Render(new Body(P("{{ds.Pdf}:pdf(w:10cm, h:2cm)}"), A4Section()), new { Pdf = landscape }).Body.Descendants<DW.Extent>().Single();
            Assert.That(box.Cy.Value, Is.EqualTo(720000), "fits into the box");
        }

        [Test]
        public void FitNoneKeepsTheRealSizeIfItFits()
        {
            var small = FakePdfRenderer.Pdf((288, 144)); // 4 x 2 inch
            var extent = Render(new Body(P("{{ds.Pdf}:pdf(fit:none)}"), A4Section()), new { Pdf = small }).Body.Descendants<DW.Extent>().Single();
            Assert.That(extent.Cx.Value, Is.EqualTo(4 * 914400));
            Assert.That(extent.Cy.Value, Is.EqualTo(2 * 914400));
        }

        [Test]
        public void SizeComesFromTheSectionOfThePlaceholder()
        {
            // first section is A4 landscape, the last one A4 portrait
            var landscapeSection = new SectionProperties(
                new PageSize { Width = (uint)A4HeightTwips, Height = (uint)A4WidthTwips, Orient = PageOrientationValues.Landscape },
                new PageMargin { Top = 1134, Right = 1417U, Bottom = 1134, Left = 1417U, Header = 708U, Footer = 708U, Gutter = 0U });
            var landscapePdf = FakePdfRenderer.Pdf((FakePdfRenderer.A4Height, FakePdfRenderer.A4Width));
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf(fit:width)}"), PWithSection("end of landscape", landscapeSection), P("portrait"), A4Section()), new { Pdf = landscapePdf });

            Assert.That(body.Descendants<DW.Extent>().Single().Cx.Value, Is.EqualTo((A4HeightTwips - (2 * MarginLeftRight)) * 635));
        }

        [Test]
        public void ColumnsReduceTheWidth()
        {
            var section = A4Section(new Columns { ColumnCount = 2, Space = "720" });
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf(fit:width)}"), section), new { Pdf = FakePdfRenderer.A4Pages(1) });

            Assert.That(body.Descendants<DW.Extent>().Single().Cx.Value, Is.EqualTo((A4WidthTwips - (2 * MarginLeftRight) - 720) / 2 * 635));
        }

        [Test]
        public void PlaceholderParagraphWithSectionBreakPassesItToTheLastPage()
        {
            var firstSection = A4Section();
            var (body, _) = Render(new Body(PWithSection("{{ds.Pdf}:pdf}", firstSection), P("next"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(2) });

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(Describe), Is.EqualTo(["[pdf]", "[pdf]", "next"]));
            Assert.That(paragraphs[0].ParagraphProperties.SectionProperties, Is.Null);
            Assert.That(paragraphs[1].ParagraphProperties.SectionProperties, Is.Not.Null);
        }

        [Test]
        public void InTableCellPagesHaveTheCellWidthAndNoPageBreaks()
        {
            var table = new Table(
                new TableProperties(new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" }),
                new TableGrid(new GridColumn { Width = "3000" }, new GridColumn { Width = "4000" }),
                new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }), P("a")),
                    new TableCell(P("{{ds.Pdf}:pdf}"))));
            var (body, _) = Render(new Body(table, P(""), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(2) });

            var cell = body.Descendants<TableCell>().Last();
            Assert.That(cell.Elements<Paragraph>().Count(), Is.EqualTo(2));
            Assert.That(cell.Descendants<Break>(), Is.Empty);
            Assert.That(cell.Descendants<DW.Extent>().First().Cx.Value, Is.EqualTo((4000 - 216) * 635), "grid column minus the default cell margins");
        }

        [Test]
        public void InHeaderTheImageIsStoredInTheHeaderPart()
        {
            WordprocessingDocument result = null;
            (_, result) = Render(new Body(P("text"), A4Section(new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdHeader" })),
                new { Pdf = FakePdfRenderer.A4Pages(2) },
                setup: document =>
                {
                    var headerPart = document.MainDocumentPart.AddNewPart<HeaderPart>("rIdHeader");
                    headerPart.Header = new Header(P("{{ds.Pdf}:pdf(pages:'1', h:1cm)}"));
                });

            var header = result.MainDocumentPart.HeaderParts.Single();
            var blip = header.Header.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single();
            Assert.That(header.GetPartById(blip.Embed.Value), Is.InstanceOf<ImagePart>());
            Assert.That(header.Header.Descendants<DW.Extent>().Single().Cy.Value, Is.EqualTo(360000));
            Assert.That(result.MainDocumentPart.ImageParts, Is.Empty);
        }

        [Test]
        public void AltTextUsesTheNameOfThePdfFile()
        {
            var (body, _) = Render(new Body(P("{{ds.Plan}:pdf}"), P("{{ds.Plan}:pdf(label:'Override: B')}"), A4Section()),
                new { Plan = new PdfFile(FakePdfRenderer.A4Pages(2), "Ground floor") });

            var descriptions = body.Descendants<DW.DocProperties>().Select(x => x.Description.Value).ToList();
            Assert.That(descriptions, Is.EqualTo(["Ground floor, page 1 of 2", "Ground floor, page 2 of 2", "Override: B, page 1 of 2", "Override: B, page 2 of 2"]));
        }

        [Test]
        public void DefaultLabelIsUsedForPlainBytes()
        {
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(1) },
                configuration: new PdfFormatterConfiguration { DefaultLabel = "Anhang", AltText = "{name} - Seite {page} von {pages} {unknown}" });

            Assert.That(body.Descendants<DW.DocProperties>().Single().Description.Value, Is.EqualTo("Anhang - Seite 1 von 1 {unknown}"));
        }

        [Test]
        public void SamePdfInALoopIsRenderedAndStoredOnce()
        {
            var renderer = new FakePdfRenderer();
            var pdf = FakePdfRenderer.A4Pages(2);
            var (body, document) = Render(new Body(P("{{#ds.Items}}"), P("{{.}:pdf}"), P("{{/ds.Items}}"), A4Section()),
                new { Items = new[] { pdf, pdf, pdf } }, renderer);

            Assert.That(body.Descendants<Drawing>().Count(), Is.EqualTo(6));
            Assert.That(renderer.Requests, Has.Count.EqualTo(1));
            Assert.That(document.MainDocumentPart.ImageParts.Count(), Is.EqualTo(2));
            var ids = body.Descendants<DW.DocProperties>().Select(x => x.Id.Value).ToList();
            Assert.That(ids.Distinct().Count(), Is.EqualTo(ids.Count), "unique drawing ids");
        }

        [Test]
        public void EmptyBytesRemoveThePlaceholder()
        {
            var renderer = new FakePdfRenderer();
            var (body, _) = Render(new Body(P("a{{ds.Pdf}:pdf}b"), A4Section()), new { Pdf = Array.Empty<byte>() }, renderer);

            Assert.That(body.Elements<Paragraph>().Single().InnerText, Is.EqualTo("ab"));
            Assert.That(renderer.GetPageSizesCalls, Is.Zero);
        }

        [Test]
        public void InvalidArgumentsThrow()
        {
            var pdf = new { Pdf = FakePdfRenderer.A4Pages(1) };
            Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf(foo:bar)}"), A4Section()), pdf));
            Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf(mode:attach)}"), A4Section()), pdf));
            Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf(dpi:5000)}"), A4Section()), pdf));
            Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf(fit:huge)}"), A4Section()), pdf));
        }

        [Test]
        public void TooManyPagesThrow()
        {
            var exception = Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()),
                new { Pdf = FakePdfRenderer.A4Pages(5) }, configuration: new PdfFormatterConfiguration { MaxPages = 4 }));
            Assert.That(exception.InnerException.Message, Does.Contain("MaxPages"));
        }

        [Test]
        public void HugePagesAreRenderedWithLowerDpi()
        {
            var renderer = new FakePdfRenderer();
            // 2 x 2 m at 150 dpi would be 139 megapixels
            Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()), new { Pdf = FakePdfRenderer.Pdf((5669, 5669)) }, renderer);
            Assert.That(renderer.Requests.Single().Dpi, Is.LessThan(150));

            Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()), new { Pdf = FakePdfRenderer.Pdf((5669, 5669)) },
                configuration: new PdfFormatterConfiguration { ReduceDpiForLargePages = false }));
        }

        [Test]
        public void RendererErrorsAreReportedWithTheName()
        {
            var exception = Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()),
                new { Pdf = new PdfFile([1, 2, 3], "Broken plan") }));
            Assert.That(exception.InnerException.Message, Does.Contain("Broken plan"));
            Assert.That(exception.InnerException.InnerException, Is.InstanceOf<InvalidDataException>());
        }

        [Test]
        public void PasswordComesFromTheProvider()
        {
            var renderer = new FakePdfRenderer();
            Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()), new { Pdf = new PdfFile(FakePdfRenderer.A4Pages(1), "Secret") }, renderer,
                new PdfFormatterConfiguration { PasswordProvider = name => name + "-pwd" });
            Assert.That(renderer.Requests.Single().Password, Is.EqualTo("Secret-pwd"));
        }

        [Test]
        public void JpegFormatAndBorder()
        {
            var renderer = new FakePdfRenderer();
            var (body, _) = Render(new Body(P("{{ds.Pdf}:pdf(fmt:jpg, q:70, border:true)}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(1) }, renderer);
            Assert.That(renderer.Requests.Single().Format, Is.EqualTo(PdfImageFormat.Jpeg));
            Assert.That(renderer.Requests.Single().JpegQuality, Is.EqualTo(70));
            Assert.That(body.Descendants<DocumentFormat.OpenXml.Drawing.Outline>().Count(), Is.EqualTo(1));
        }

        [Test]
        public void DisableImageCompressionAddsTheSetting()
        {
            var (_, document) = Render(new Body(P("{{ds.Pdf}:pdf}"), A4Section()), new { Pdf = FakePdfRenderer.A4Pages(1) },
                configuration: new PdfFormatterConfiguration { DisableImageCompression = true },
                setup: d => d.MainDocumentPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings(new Zoom { Percent = "100" }, new DefaultTabStop { Val = 708 }, new CharacterSpacingControl { Val = CharacterSpacingValues.DoNotCompress }));

            var settings = document.MainDocumentPart.DocumentSettingsPart.Settings;
            Assert.That(settings.GetFirstChild<DoNotAutoCompressPictures>(), Is.Not.Null);
        }

        [Test]
        public void CanHandle()
        {
            var formatter = new PdfFormatter(new FakePdfRenderer());
            Assert.That(formatter.CanHandle(typeof(byte[]), "PDF"), Is.True);
            Assert.That(formatter.CanHandle(typeof(PdfFile), "pdf"), Is.True);
            Assert.That(formatter.CanHandle(typeof(byte[]), "img"), Is.False);
            Assert.That(formatter.CanHandle(typeof(string), "pdf"), Is.False);
        }

        [Test]
        public void ArgumentsAreJoinedAgainAfterTheCoreSplitThemAtCommas()
        {
            var arguments = PdfArguments.Parse(["pages:'1-3", "7'", "label:'Plan: A", "B'", "dpi:200"]);
            Assert.That(arguments["pages"], Is.EqualTo("1-3,7"));
            Assert.That(arguments["label"], Is.EqualTo("Plan: A,B"));
            Assert.That(arguments["dpi"], Is.EqualTo("200"));
            Assert.That(PdfArguments.ParsePageSelection("2;4-5 1", 5), Is.EqualTo([1, 3, 4, 0]));
            Assert.That(PdfArguments.ParseLengthEmu("w", "2.5cm"), Is.EqualTo(900000));
        }

        [Test]
        public void FormatterContextCultureIsUsedForNumbers()
        {
            var tokens = PageTexts.CreateTokens(new PdfFormatterConfiguration(), new PageDescription("x", null, 1, 2, 2383.94, 1683.78, 0.664), CultureInfo.GetCultureInfo("de-CH"));
            Assert.That(tokens["width"], Is.EqualTo("841"));
            Assert.That(tokens["height"], Is.EqualTo("594"));
            Assert.That(tokens["sheet"], Is.EqualTo("A1"));
            Assert.That(tokens["orientation"], Is.EqualTo("landscape"));
            Assert.That(tokens["scale"], Does.StartWith("66"));
        }

        [Test]
        public void SheetNames()
        {
            Assert.That(PageTexts.FindSheetName(FakePdfRenderer.A4Width, FakePdfRenderer.A4Height), Is.EqualTo("A4"));
            Assert.That(PageTexts.FindSheetName(FakePdfRenderer.A3Height, FakePdfRenderer.A3Width), Is.EqualTo("A3"));
            Assert.That(PageTexts.FindSheetName(612, 792), Is.EqualTo("Letter"));
            Assert.That(PageTexts.FindSheetName(500, 500), Is.Null);
        }

        [Test]
        public void ModelPropertyAttributeSelectsThePdfFormatterWithoutPrefix()
        {
            var (body, _) = Render(new Body(P("{{ds.Pdf}}"), A4Section()),
                new ModelWithDefaultFormatter { Pdf = FakePdfRenderer.A4Pages(1) });
            Assert.That(body.Descendants<Drawing>().Count(), Is.EqualTo(1));
        }

        private sealed class ModelWithDefaultFormatter
        {
            [ModelProperty(DefaultFormatter = "pdf(fit:width)")]
            public byte[] Pdf { get; init; }
        }
    }
}

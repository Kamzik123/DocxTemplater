using System.Buffers.Binary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Pdf.Pdfium.Test
{
    internal class PdfiumPdfRendererTest
    {
        [Test]
        public void PageSizesAreInPointsWithTheRotationApplied()
        {
            var pdf = TestPdf.Create((TestPdf.A4Width, TestPdf.A4Height, 0), (TestPdf.A3Height, TestPdf.A3Width, 0), (TestPdf.A4Width, TestPdf.A4Height, 90));
            var sizes = new PdfiumPdfRenderer().GetPageSizes(pdf, null);

            Assert.That(sizes, Has.Count.EqualTo(3));
            Assert.That(sizes[0].WidthPoints, Is.EqualTo(TestPdf.A4Width).Within(0.01));
            Assert.That(sizes[1].WidthPoints, Is.EqualTo(TestPdf.A3Height).Within(0.01));
            Assert.That(sizes[2].WidthPoints, Is.EqualTo(TestPdf.A4Height).Within(0.01), "rotated page is landscape");
            Assert.That(sizes[2].HeightPoints, Is.EqualTo(TestPdf.A4Width).Within(0.01));
        }

        [Test]
        public void RendersPngAtTheRequestedResolution()
        {
            var pdf = TestPdf.Create((TestPdf.A4Width, TestPdf.A4Height, 0), (TestPdf.A3Height, TestPdf.A3Width, 0));
            var pages = new PdfiumPdfRenderer().RenderPages(pdf, new PdfRenderRequest([1, 0], 150, PdfImageFormat.Png, 85, true, true, null)).ToList();

            Assert.That(pages.Select(x => x.PageIndex), Is.EqualTo([1, 0]));
            var (width, height) = ReadPngSize(pages[0].ImageBytes);
            Assert.That(width, Is.EqualTo(2480).Within(2), "A3 landscape at 150 dpi");
            Assert.That(height, Is.EqualTo(1754).Within(2));
            Assert.That(pages[0].WidthPoints, Is.EqualTo(TestPdf.A3Height).Within(0.01));
            Assert.That(pages[1].Format, Is.EqualTo(PdfImageFormat.Png));
        }

        [Test]
        public void RendersJpeg()
        {
            var pdf = TestPdf.Create((TestPdf.A4Width, TestPdf.A4Height, 0));
            var page = new PdfiumPdfRenderer().RenderPages(pdf, new PdfRenderRequest([0], 72, PdfImageFormat.Jpeg, 80, true, true, null)).Single();

            Assert.That(page.ImageBytes[0], Is.EqualTo(0xFF));
            Assert.That(page.ImageBytes[1], Is.EqualTo(0xD8));
        }

        [Test]
        public void CorruptPdfGivesATemplateExceptionWithTheName()
        {
            var exception = Assert.Throws<OpenXmlTemplateException>(() => Render(
                new Body(new Paragraph(new Run(new Text("{{ds.Plan}:pdf}")))),
                new { Plan = new PdfFile([1, 2, 3, 4], "Broken.pdf") }));
            Assert.That(exception.InnerException.Message, Does.Contain("Broken.pdf"));
        }

        [Test]
        public void EndToEndWithPagesAndPlans()
        {
            var appendix = TestPdf.Create((TestPdf.A4Width, TestPdf.A4Height, 0), (TestPdf.A4Height, TestPdf.A4Width, 0));
            var plan = TestPdf.Create((TestPdf.A3Height, TestPdf.A3Width, 0));
            var body = new Body(
                new Paragraph(new Run(new Text("Appendix A"))),
                new Paragraph(new Run(new Text("{{ds.Appendix}:pdf}"))),
                new Paragraph(new Run(new Text("Appendix B"))),
                new Paragraph(new Run(new Text("{{ds.Plan}:pdf(fit:section)}"))),
                new Paragraph(new Run(new Text("End"))),
                new SectionProperties(new PageSize { Width = 11906U, Height = 16838U }, new PageMargin { Top = 1134, Right = 1417U, Bottom = 1134, Left = 1417U, Header = 708U, Footer = 708U, Gutter = 0U }));

            var document = Render(body, new { Appendix = appendix, Plan = new PdfFile(plan, "Ground floor") });

            var result = document.MainDocumentPart.Document.Body;
            Assert.That(result.Descendants<DW.Inline>().Count(), Is.EqualTo(2));
            Assert.That(result.Descendants<DW.Anchor>().Count(), Is.EqualTo(1));
            Assert.That(document.MainDocumentPart.ImageParts.Select(x => x.ContentType), Has.All.EqualTo("image/png"));
            var planSize = result.Descendants<SectionProperties>().Select(x => x.GetFirstChild<PageSize>()).First(x => x.Width.Value > x.Height.Value);
            Assert.That(planSize.Width.Value, Is.EqualTo(23811));
        }

        /// <summary>
        /// Writes a sample document with all layouts to the path in DOCX_TEMPLATER_PDF_SAMPLE (default: temp folder),
        /// plus the plan PDF used in it - for checks in Word: no blank pages, plans edge to edge, "Save as PDF",
        /// replacing the plan pages with the originals, measuring the 10 cm scale bar in a print.
        /// </summary>
        [Test]
        [Explicit("writes a sample document for a manual check in Word")]
        public void SampleDocument()
        {
            var path = Environment.GetEnvironmentVariable("DOCX_TEMPLATER_PDF_SAMPLE") ?? Path.Combine(Path.GetTempPath(), "DocxTemplater.Pdf.Sample.docx");
            var appendix = TestPdf.Create((TestPdf.A4Width, TestPdf.A4Height, 0), (TestPdf.A4Height, TestPdf.A4Width, 0), (TestPdf.A4Width, TestPdf.A4Height, 0));
            var plans = TestPdf.Create((TestPdf.A3Height, TestPdf.A3Width, 0), (TestPdf.A3Height, TestPdf.A3Width, 0), (TestPdf.A4Width, TestPdf.A4Height, 0), (TestPdf.A1Height, TestPdf.A1Width, 0));
            File.WriteAllBytes(Path.ChangeExtension(path, ".plans.pdf"), plans);

            var body = new Body(
                Heading("Report"),
                Text("The appendix is scaled into the text area (fit:page):"),
                Text("{{ds.Appendix}:pdf}"),
                Heading("Plans"),
                Text("Plans at true size on pages of their own (fit:section): 2 × A3 landscape, A4, A1 landscape (scaled, with overlay)."),
                Text("{{ds.Plans}:pdf(fit:section)}"),
                Heading("Plans in a loop"),
                Text("{{#ds.Loop}}"),
                Text("Plan {{.Name}}"),
                Text("{{.Pdf}:pdf(fit:section, overlay:always)}"),
                Text("{{/ds.Loop}}"),
                Heading("End"),
                Text("Page numbering continues through the plans, the header is back."),
                new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdHeader" },
                    new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdFooter" },
                    new PageSize { Width = 11906U, Height = 16838U },
                    new PageMargin { Top = 1417, Right = 1417U, Bottom = 1134, Left = 1417U, Header = 708U, Footer = 708U, Gutter = 0U }));
            var a3 = TestPdf.Create((TestPdf.A3Height, TestPdf.A3Width, 0));
            var document = Render(body, new
            {
                Appendix = new PdfFile(appendix, "Appendix"),
                Plans = new PdfFile(plans, "Plan set 102"),
                Loop = new[] { new { Name = "A", Pdf = new PdfFile(a3, "Plan A") }, new { Name = "B", Pdf = new PdfFile(a3, "Plan B") } },
            }, setup: d =>
            {
                d.MainDocumentPart.AddNewPart<HeaderPart>("rIdHeader").Header = new Header(new Paragraph(new Run(new Text("Header of the document"))));
                d.MainDocumentPart.AddNewPart<FooterPart>("rIdFooter").Footer = new Footer(new Paragraph(new Run(new Text("Page ") { Space = SpaceProcessingModeValues.Preserve }), new SimpleField(new Run(new Text("1"))) { Instruction = " PAGE " }));
            });
            using (var file = File.Create(path))
            {
                using (document.Clone(file))
                {
                }
            }
            TestContext.Out.WriteLine(path);

            static Paragraph Heading(string text)
            {
                return new Paragraph(new Run(new RunProperties(new Bold(), new FontSize { Val = "32" }), new Text(text)));
            }

            static Paragraph Text(string text)
            {
                return new Paragraph(new Run(new Text(text)));
            }
        }
        private static WordprocessingDocument Render(Body templateBody, object model, Action<WordprocessingDocument> setup = null)
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                wpDocument.AddMainDocumentPart().Document = new Document(templateBody);
                setup?.Invoke(wpDocument);
            }
            memStream.Position = 0;
            var template = new DocxTemplate(memStream);
            template.RegisterFormatter(new PdfFormatter(new PdfiumPdfRenderer()));
            template.BindModel("ds", model);
            var result = template.Process();
            template.Validate();
            result.Position = 0;
            return WordprocessingDocument.Open(result, false);
        }

        private static (int Width, int Height) ReadPngSize(byte[] png)
        {
            // IHDR: width and height are the first fields after the 8 byte signature and the chunk header
            return (BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(16, 4)), BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(20, 4)));
        }
    }
}

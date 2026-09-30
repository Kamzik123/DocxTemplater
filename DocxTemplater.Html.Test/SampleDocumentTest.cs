using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Images;

namespace DocxTemplater.Html.Test
{
    internal class SampleDocumentTest
    {
        /// <summary>
        /// Renders a document with most supported HTML features for a visual check in Word / LibreOffice.
        /// Set DOCX_TEMPLATER_HTML_SAMPLE to the output path to run it.
        /// </summary>
        [Test]
        [Explicit]
        public void RenderSampleDocument()
        {
            var outputPath = Environment.GetEnvironmentVariable("DOCX_TEMPLATER_HTML_SAMPLE");
            if (string.IsNullOrEmpty(outputPath))
            {
                Assert.Ignore("DOCX_TEMPLATER_HTML_SAMPLE not set");
            }

            const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
            var html = $"""
                        <h1>Heading 1</h1>
                        <p>Plain paragraph with <b>bold</b>, <i>italic</i>, <u>underline</u>, <s>strike</s>, H<sub>2</sub>O, x<sup>2</sup>,
                        <code>code</code>, <mark>marked</mark> and <span style="color:#c00;background-color:#ffd">colored</span> text.</p>
                        <h2>Lists</h2>
                        <ul><li>Bullet one</li><li>Bullet two<ol type="a"><li>nested a</li><li>nested b</li></ol></li></ul>
                        <ol start="5"><li>five</li><li>six</li></ol>
                        <h2>Table</h2>
                        <table>
                          <thead><tr><th>Name</th><th>Qty</th><th>Price</th></tr></thead>
                          <tr><td rowspan="2">Merged</td><td>1</td><td style="text-align:right">10.00</td></tr>
                          <tr><td>2</td><td style="text-align:right;background-color:#eef">20.00</td></tr>
                          <tr><td colspan="2"><b>Total</b></td><td style="text-align:right">30.00</td></tr>
                        </table>
                        <h2>Other blocks</h2>
                        <blockquote>A quote from someone.</blockquote>
                        <pre>preformatted
                            indented line</pre>
                        <hr>
                        <p style="text-align:center">Centered with a <a href="https://github.com/Amberg/DocxTemplater">link</a>
                        and an image <img src="data:image/png;base64,{png}" width="40" height="20"></p>
                        """;

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("Before the html placeholder"))),
                    new Paragraph(new Run(new Text("{{ds}:html}"))),
                    new Paragraph(new Run(new Text("After the html placeholder")))));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.RegisterFormatter(new HtmlFormatter());
            docTemplate.BindModel("ds", html);
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var file = File.Create(outputPath);
            result.Position = 0;
            result.CopyTo(file);
        }
    }
}

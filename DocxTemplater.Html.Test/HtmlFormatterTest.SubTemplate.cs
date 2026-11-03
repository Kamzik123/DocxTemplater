using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dml = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxTemplater.Html.Test
{
    internal partial class HtmlFormatterTest
    {
        // Regression test: a {{...}:T(...)} placeholder inside the HTML is resolved while the converted content
        // is still detached from the document. The sub-template formatter must still import the parts and styles
        // the inserted document references - otherwise images dangle and paragraphs point at undefined styles.
        [Test]
        public void SubTemplateInsideHtml_ImportsImagesAndStyles()
        {
            using var subDocStream = new MemoryStream();
            using (var subDocument = WordprocessingDocument.Create(subDocStream, WordprocessingDocumentType.Document))
            {
                var subMainPart = subDocument.AddMainDocumentPart();
                subMainPart.Document = new Document(new Body());
                var stylesPart = subMainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new Style(new StyleName { Val = "Sub Title" }, new StyleRunProperties(new Bold()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "SubTitle"
                });
                var imagePart = subMainPart.AddImagePart(ImagePartType.Png);
                using (var imageStream = new MemoryStream(Convert.FromBase64String(PngBase64)))
                {
                    imagePart.FeedData(imageStream);
                }
                subMainPart.Document.Body.Append(
                    new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "SubTitle" }), new Run(new Text("Inserted title"))),
                    new Paragraph(new Run(CreateInlineImage(subMainPart.GetIdOfPart(imagePart)))));
                subMainPart.Document.Save();
            }

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds.Html}:html}")))));
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new HtmlFormatter());
            docTemplate.BindModel("ds", new
            {
                Html = "<p>before</p><p>{{ds}:T('ds.SubDocument')}</p><p>after</p>",
                SubDocument = subDocStream.ToArray()
            });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var document = WordprocessingDocument.Open(result, false);
            var mainDocumentPart = document.MainDocumentPart;
            var body = mainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Does.Contain("Inserted title"));

            var blip = body.Descendants<Dml.Blip>().Single();
            Assert.That(mainDocumentPart.GetPartById(blip.Embed.Value), Is.InstanceOf<ImagePart>(), "the image must be copied into the target");

            var styleId = body.Descendants<ParagraphStyleId>().Single(x => x.Val == "SubTitle");
            Assert.That(mainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().Any(x => x.StyleId == styleId.Val.Value), Is.True, "the style must be copied into the target");
        }

        private static Drawing CreateInlineImage(string relationshipId)
        {
            const long size = 9525L;
            return new Drawing(new Wp.Inline(
                new Wp.Extent { Cx = size, Cy = size },
                new Wp.DocProperties { Id = 1U, Name = "Picture 1" },
                new Dml.Graphic(new Dml.GraphicData(
                    new Pic.Picture(
                        new Pic.NonVisualPictureProperties(new Pic.NonVisualDrawingProperties { Id = 0U, Name = "image.png" }, new Pic.NonVisualPictureDrawingProperties()),
                        new Pic.BlipFill(new Dml.Blip { Embed = relationshipId }, new Dml.Stretch(new Dml.FillRectangle())),
                        new Pic.ShapeProperties(new Dml.Transform2D(new Dml.Offset { X = 0L, Y = 0L }, new Dml.Extents { Cx = size, Cy = size }), new Dml.PresetGeometry { Preset = Dml.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    [TestFixture]
    internal class SubTemplateNumberingImportTest
    {
        // Regression test: list numbering of an inserted document must be copied. Host and sub document both
        // define numId 1 - without the import, the inserted list silently continues the host's list (wrong format).
        [Test]
        public void SubTemplateInsert_ImportsNumberingWithNewIds()
        {
            using var subDocStream = new MemoryStream();
            using (var subDocument = WordprocessingDocument.Create(subDocStream, WordprocessingDocumentType.Document))
            {
                var subMainPart = subDocument.AddMainDocumentPart();
                AddNumbering(subMainPart, NumberFormatValues.LowerLetter);
                subMainPart.Document = new Document(new Body(
                    new Paragraph(new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })), new Run(new Text("sub item")))));
                subMainPart.Document.Save();
            }

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                AddNumbering(mainPart, NumberFormatValues.Decimal);
                mainPart.Document = new Document(new Body(
                    new Paragraph(new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })), new Run(new Text("host item"))),
                    new Paragraph(new Run(new Text("{{ds}:template('ds.SubDocument')}")))));
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { SubDocument = subDocStream.ToArray() });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var numbering = document.MainDocumentPart.NumberingDefinitionsPart.Numbering;

            int NumId(string text)
            {
                return body.Elements<Paragraph>().Single(x => x.InnerText == text).ParagraphProperties.NumberingProperties.NumberingId.Val.Value;
            }
            NumberFormatValues Format(int numId)
            {
                var abstractId = numbering.Elements<NumberingInstance>().Single(x => x.NumberID == numId).AbstractNumId.Val.Value;
                return numbering.Elements<AbstractNum>().Single(x => x.AbstractNumberId == abstractId).Elements<Level>().First().NumberingFormat.Val.Value;
            }

            Assert.That(NumId("host item"), Is.EqualTo(1), "the host list is unchanged");
            Assert.That(NumId("sub item"), Is.Not.EqualTo(1), "the inserted list gets its own numId");
            Assert.That(Format(NumId("host item")), Is.EqualTo(NumberFormatValues.Decimal));
            Assert.That(Format(NumId("sub item")), Is.EqualTo(NumberFormatValues.LowerLetter), "the inserted list keeps its own definition");
        }

        private static void AddNumbering(MainDocumentPart mainPart, NumberFormatValues format)
        {
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(
                new AbstractNum(new Level(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = format }, new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
        }
    }
}

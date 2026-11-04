using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html.Test
{
    internal partial class HtmlFormatterTest
    {
        private static Style ParagraphStyle(string name, string id, bool outline = false, bool isDefault = false)
        {
            var style = new Style(new StyleName { Val = name }) { Type = StyleValues.Paragraph, StyleId = id };
            if (isDefault)
            {
                style.Default = true;
            }
            if (outline)
            {
                style.AppendChild(new StyleParagraphProperties(new OutlineLevel { Val = 1 }));
            }
            return style;
        }

        private static Styles HeadingTestStyles()
        {
            return new Styles(
                ParagraphStyle("heading 2", "Heading2"),
                ParagraphStyle("heading 3", "Heading3"),
                ParagraphStyle("Custom Heading", "CustomHeading", outline: true));
        }

        private static string StyleOf(Paragraph paragraph)
        {
            return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        }

        private static (Body Body, WordprocessingDocument Document) RenderInBodyStyle(string html, bool fromTemplate, params Style[] extraStyles)
        {
            var styles = new Styles(ParagraphStyle("Normal", "Normal", isDefault: true), ParagraphStyle("Body", "Body"));
            foreach (var style in extraStyles)
            {
                styles.AppendChild(style);
            }
            var template = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Body" }),
                new Run(new Text("{{ds}:html}")));
            return RenderDocument(new Body(template), html, styles: styles,
                configuration: new HtmlFormatterConfiguration { ParagraphStyleFromTemplate = fromTemplate });
        }

        [TestCase("Custom Heading", TestName = "HeadingStyles_UsesConfiguredStyleByName")]
        [TestCase("CustomHeading", TestName = "HeadingStyles_UsesConfiguredStyleById")]
        public void HeadingStyles_UsesConfiguredStyle(string configured)
        {
            var body = Render("{{ds}:html}", "<h2>x</h2>", styles: HeadingTestStyles(),
                configuration: new HtmlFormatterConfiguration { HeadingStyles = { [2] = configured } });

            Assert.That(StyleOf(body.Elements<Paragraph>().Single()), Is.EqualTo("CustomHeading"));
        }

        [TestCase("Does not exist")]
        [TestCase("  ")]
        public void HeadingStyles_MissingStyleFallsBackToBuiltIn(string configured)
        {
            var body = Render("{{ds}:html}", "<h2>x</h2>", styles: HeadingTestStyles(),
                configuration: new HtmlFormatterConfiguration { HeadingStyles = { [2] = configured } });

            Assert.That(StyleOf(body.Elements<Paragraph>().Single()), Is.EqualTo("Heading2"));
        }

        [Test]
        public void HeadingStyles_OtherLevelsUnaffected()
        {
            var body = Render("{{ds}:html}", "<h3>x</h3>", styles: HeadingTestStyles(),
                configuration: new HtmlFormatterConfiguration { HeadingStyles = { [2] = "Custom Heading" } });

            Assert.That(StyleOf(body.Elements<Paragraph>().Single()), Is.EqualTo("Heading3"));
        }

        [Test]
        public void HeadingStyles_MissingEverywhereUsesDirectFormatting()
        {
            var body = Render("{{ds}:html}", "<h2>x</h2>",
                configuration: new HtmlFormatterConfiguration { HeadingStyles = { [2] = "Does not exist" } });

            var paragraph = body.Elements<Paragraph>().Single();
            Assert.That(StyleOf(paragraph), Is.Null);
            Assert.That(paragraph.Descendants<Run>().Single().RunProperties.Bold, Is.Not.Null);
        }

        [Test]
        public void HeadingWithStyleDoesNotInheritFormattingOfAncestors()
        {
            var body = Render("{{ds}:html}", "<div style=\"font-family: Arial; font-size: 20pt\"><h2>x</h2><p>y</p></div>", styles: HeadingTestStyles(),
                configuration: new HtmlFormatterConfiguration { HeadingStyles = { [2] = "Custom Heading" } });

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs[0].Descendants<Run>().Single().RunProperties, Is.Null);
            Assert.That(paragraphs[1].Descendants<Run>().Single().RunProperties.RunFonts.Ascii.Value, Is.EqualTo("Arial"), "other paragraphs still inherit it");
        }

        [Test]
        public void HeadingWithStyleKeepsItsOwnCss()
        {
            var body = Render("{{ds}:html}", "<div style=\"font-size: 20pt\"><h2 style=\"color: red\">x</h2></div>", styles: HeadingTestStyles(),
                configuration: new HtmlFormatterConfiguration { HeadingStyles = { [2] = "Custom Heading" } });

            var properties = body.Elements<Paragraph>().Single().Descendants<Run>().Single().RunProperties;
            Assert.That(properties.Color.Val.Value, Is.EqualTo("FF0000"));
            Assert.That(properties.FontSize, Is.Null);
        }

        [Test]
        public void HeadingWithoutStyleKeepsFormattingOfAncestors()
        {
            var body = Render("{{ds}:html}", "<div style=\"font-family: Arial\"><h2>x</h2></div>");

            var properties = body.Elements<Paragraph>().Single().Descendants<Run>().Single().RunProperties;
            Assert.That(properties.RunFonts.Ascii.Value, Is.EqualTo("Arial"));
            Assert.That(properties.Bold, Is.Not.Null);
        }

        [Test]
        public void Clone_CopiesHeadingStylesAndParagraphStyleFromTemplate()
        {
            var original = new HtmlFormatterConfiguration { ParagraphStyleFromTemplate = true, HeadingStyles = { [1] = "One" } };

            var clone = original.Clone();
            clone.HeadingStyles[1] = "Changed";
            clone.HeadingStyles[2] = "Two";

            Assert.That(clone.ParagraphStyleFromTemplate, Is.True);
            Assert.That(original.HeadingStyles, Has.Count.EqualTo(1));
            Assert.That(original.HeadingStyles[1], Is.EqualTo("One"));
        }

        [Test]
        public void ParagraphStyleFromTemplate_AlignedParagraphUsesTemplateStyle()
        {
            var (body, _) = RenderInBodyStyle("<p>a</p><p style=\"text-align: center\">b</p><p>c</p>", fromTemplate: true);

            var paragraph = body.Elements<Paragraph>().ElementAt(1);
            Assert.That(StyleOf(paragraph), Is.EqualTo("Body"));
            Assert.That(paragraph.ParagraphProperties.Justification.Val.Value, Is.EqualTo(JustificationValues.Center));
        }

        [Test]
        public void ParagraphStyleFromTemplate_GeneratedStyleIsBasedOnTemplateStyle()
        {
            var (body, document) = RenderInBodyStyle("<p>a</p><p style=\"text-align: center; color: red\">b</p><p>c</p>", fromTemplate: true);

            var styleId = StyleOf(body.Elements<Paragraph>().ElementAt(1));
            Assert.That(styleId, Does.StartWith("html"));
            var style = document.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().Single(x => x.StyleId == styleId);
            Assert.That(style.BasedOn.Val.Value, Is.EqualTo("Body"));
        }

        [Test]
        public void ParagraphStyleFromTemplate_Off_UsesDefault()
        {
            var (body, document) = RenderInBodyStyle("<p>a</p><p style=\"text-align: center; color: red\">b</p><p style=\"text-align: center\">c</p><p>d</p>", fromTemplate: false);

            var paragraphs = body.Elements<Paragraph>().ToList();
            var style = document.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().Single(x => x.StyleId == StyleOf(paragraphs[1]));
            Assert.That(style.BasedOn.Val.Value, Is.EqualTo("Normal"));
            Assert.That(StyleOf(paragraphs[2]), Is.Null);
        }

        [Test]
        public void ParagraphStyleFromTemplate_ListAndHeadingKeepOwnStyles()
        {
            var (body, _) = RenderInBodyStyle("<p>a</p><h2>b</h2><ul><li>c</li></ul><blockquote>d</blockquote><p>e</p>", fromTemplate: true,
                ParagraphStyle("heading 2", "Heading2"), ParagraphStyle("List Paragraph", "ListParagraph"), ParagraphStyle("Quote", "Quote"));

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(StyleOf(paragraphs[1]), Is.EqualTo("Heading2"));
            Assert.That(StyleOf(paragraphs[2]), Is.EqualTo("ListParagraph"));
            Assert.That(StyleOf(paragraphs[3]), Is.EqualTo("Quote"));
        }

        [Test]
        public void ParagraphStyleFromTemplate_EmulatedHeadingAndQuoteDoNotUseTemplateStyle()
        {
            var (body, _) = RenderInBodyStyle("<p>a</p><h2>b</h2><blockquote>c</blockquote><p>d</p>", fromTemplate: true);

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(StyleOf(paragraphs[1]), Is.Null);
            Assert.That(StyleOf(paragraphs[2]), Is.Null);
        }

        [Test]
        public void ParagraphStyleFromTemplate_PreformattedTextIsNotMergedIntoTemplateText()
        {
            var styles = new Styles(ParagraphStyle("Normal", "Normal", isDefault: true), ParagraphStyle("Body", "Body"));
            var template = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Body" }),
                new Run(new Text("Intro: {{ds}:html} end") { Space = SpaceProcessingModeValues.Preserve }));

            var (body, _) = RenderDocument(new Body(template), "<p>a</p><pre>code</pre>", styles: styles,
                configuration: new HtmlFormatterConfiguration { ParagraphStyleFromTemplate = true });

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo((string[])["Intro: a", "code", " end"]));
        }

        [Test]
        public void ParagraphStyleFromTemplate_TableCellsUnaffected()
        {
            var (body, _) = RenderInBodyStyle("<p style=\"text-align: center\">a</p><table><caption>t</caption><tr><td style=\"text-align: center\">b</td></tr></table><p>c</p>", fromTemplate: true);

            Assert.That(StyleOf(body.Elements<Paragraph>().First()), Is.EqualTo("Body"));
            var table = body.Elements<Table>().Single();
            Assert.That(StyleOf(table.Descendants<Paragraph>().Single()), Is.Null);
        }
    }
}

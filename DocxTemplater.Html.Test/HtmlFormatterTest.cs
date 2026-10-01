using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Images;

namespace DocxTemplater.Html.Test
{
    internal partial class HtmlFormatterTest
    {
        // 1x1 pixel png
        private const string PngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

        [Test]
        public void InlineHtmlIsMergedIntoTemplateParagraph()
        {
            var body = Render("Hello {{ds}:html} World", "<b>bold</b> and <i>italic</i>");

            var paragraph = body.Elements<Paragraph>().Single();
            Assert.That(paragraph.InnerText, Is.EqualTo("Hello bold and italic World"));
            var boldRun = paragraph.Elements<Run>().Single(r => r.InnerText == "bold");
            Assert.That(boldRun.RunProperties?.Bold, Is.Not.Null);
            var italicRun = paragraph.Elements<Run>().Single(r => r.InnerText == "italic");
            Assert.That(italicRun.RunProperties?.Italic, Is.Not.Null);
        }

        [Test]
        public void TemplateRunFormattingIsKept()
        {
            var body = Render(new Paragraph(new Run(new RunProperties(new Color { Val = "FF0000" }), new Text("{{ds}:html}"))), "<b>x</b>");

            var run = body.Descendants<Run>().Single();
            Assert.That(run.RunProperties.Color.Val.Value, Is.EqualTo("FF0000"));
            Assert.That(run.RunProperties.Bold, Is.Not.Null);
        }

        [Test]
        public void ParagraphsAndHeadingWithoutStylesUseDirectFormatting()
        {
            var body = Render("{{ds}:html}", "<h1>Title</h1><p>One</p><p>Two</p>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo((string[])["Title", "One", "Two"]));
            var titleRun = paragraphs[0].Descendants<Run>().Single();
            Assert.That(titleRun.RunProperties.Bold, Is.Not.Null);
            Assert.That(titleRun.RunProperties.FontSize.Val.Value, Is.EqualTo("40"));
        }

        [Test]
        public void HeadingStyleIsFoundByNameInLocalizedTemplate()
        {
            var body = Render("{{ds}:html}", "<h1>Nadpis</h1><p>Text</p>", styles: new Styles(
                new Style(new StyleName { Val = "heading 1" }) { Type = StyleValues.Paragraph, StyleId = "Nadpis1" }));

            var heading = body.Elements<Paragraph>().First();
            Assert.That(heading.ParagraphProperties.ParagraphStyleId.Val.Value, Is.EqualTo("Nadpis1"));
            Assert.That(heading.Descendants<Run>().Single().RunProperties, Is.Null, "a style was found - no direct formatting");
        }

        [Test]
        public void HeadingIsNotMergedIntoTemplateText()
        {
            var body = Render("Intro: {{ds}:html} end", "<h2>Heading</h2>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo((string[])["Intro: ", "Heading", " end"]));
        }

        [Test]
        public void TextAfterPlaceholderContinuesLastParagraph()
        {
            var body = Render("{{ds}:html} end", "<p>first</p><p>last</p>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo((string[])["first", "last end"]));
        }

        [Test]
        public void WhitespaceIsCollapsedLikeInABrowser()
        {
            var body = Render("{{ds}:html}", "<p>\n   a    b\n\t<b> c </b> d&nbsp;&amp;   </p>\n\n  <p>  next</p>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo((string[])["a b c d &", "next"]));
        }

        [Test]
        public void LineBreaks()
        {
            var body = Render("{{ds}:html}", "<p>a<br>b<br></p><p><br></p><p>c</p>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs, Has.Count.EqualTo(3));
            Assert.That(paragraphs[0].Descendants<Break>().Count(), Is.EqualTo(1), "trailing <br> is dropped");
            Assert.That(paragraphs[1].InnerText, Is.Empty, "<p><br></p> is an empty line");
            Assert.That(paragraphs[2].InnerText, Is.EqualTo("c"));
        }

        [Test]
        public void InlineCss()
        {
            var body = Render("{{ds}:html}",
                "<span style=\"color: #3c3; background-color: rgb(255, 0, 0); font-size: 16px; font-family: 'Arial', sans-serif; text-decoration: underline line-through\">x</span>");

            var properties = body.Descendants<Run>().Single().RunProperties;
            Assert.That(properties.Color.Val.Value, Is.EqualTo("33CC33"));
            Assert.That(properties.Shading.Fill.Value, Is.EqualTo("FF0000"));
            Assert.That(properties.FontSize.Val.Value, Is.EqualTo("24"));
            Assert.That(properties.RunFonts.Ascii.Value, Is.EqualTo("Arial"));
            Assert.That(properties.Underline.Val.Value, Is.EqualTo(UnderlineValues.Single));
            Assert.That(properties.Strike, Is.Not.Null);
        }

        [Test]
        public void ParagraphAlignment()
        {
            var body = Render("{{ds}:html}", "<p style=\"text-align:center\">c</p><p align=\"right\">r</p>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs[0].ParagraphProperties.Justification.Val.Value, Is.EqualTo(JustificationValues.Center));
            Assert.That(paragraphs[1].ParagraphProperties.Justification.Val.Value, Is.EqualTo(JustificationValues.Right));
        }

        [Test]
        public void NestedLists()
        {
            var html = "<ul><li>a<ul><li>a1</li></ul>after nested</li><li>b</li></ul><ol start=\"3\"><li>three</li><li>four</li></ol><ol><li>one</li></ol>";
            var (body, document) = RenderDocument("{{ds}:html}", html);

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo((string[])["a", "a1", "after nested", "b", "three", "four", "one"]));


            Assert.That(Level(paragraphs[0]), Is.EqualTo(0));
            Assert.That(Level(paragraphs[1]), Is.EqualTo(1));
            Assert.That(NumId(paragraphs[2]), Is.Null, "continuation of a list item has no bullet");
            Assert.That(paragraphs[2].ParagraphProperties.Indentation.Left.Value, Is.EqualTo("720"));
            Assert.That(NumId(paragraphs[3]), Is.EqualTo(NumId(paragraphs[0])));

            Assert.That(NumId(paragraphs[4]), Is.EqualTo(NumId(paragraphs[5])));
            Assert.That(NumId(paragraphs[6]), Is.Not.EqualTo(NumId(paragraphs[4])), "every ordered list restarts its numbering");

            var numbering = document.MainDocumentPart.NumberingDefinitionsPart.Numbering;
            var instance = numbering.Elements<NumberingInstance>().Single(x => x.NumberID.Value == NumId(paragraphs[4]));
            Assert.That(instance.Descendants<StartOverrideNumberingValue>().Single().Val.Value, Is.EqualTo(3));
        }

        [Test]
        public void ListDefinitionsAreReusedAcrossPlaceholders()
        {
            var html = "<ul><li>x</li></ul><ol><li>y</li></ol>";
            var (_, document) = RenderDocument(new Body(
                new Paragraph(new Run(new Text("{{#ds.Items}}"))),
                new Paragraph(new Run(new Text("{{.}:html}"))),
                new Paragraph(new Run(new Text("{{/ds.Items}}")))), new { Items = new[] { html, html, html } });

            var numbering = document.MainDocumentPart.NumberingDefinitionsPart.Numbering;
            Assert.That(numbering.Elements<AbstractNum>().Count(), Is.EqualTo(2));
        }

        [Test]
        public void TableWithColspanAndRowspan()
        {
            var html = """
                       <table>
                         <thead><tr><th>H1</th><th>H2</th><th>H3</th></tr></thead>
                         <tbody>
                           <tr><td rowspan="2">R</td><td colspan="2">C</td></tr>
                           <tr><td>x</td><td style="background-color:#ff0">y</td></tr>
                           <tr><td>short</td></tr>
                         </tbody>
                       </table>
                       """;
            var body = Render("{{ds}:html}", html);

            var table = body.Elements<Table>().Single();
            Assert.That(table.GetFirstChild<TableGrid>().Elements<GridColumn>().Count(), Is.EqualTo(3));
            var rows = table.Elements<TableRow>().ToList();
            Assert.That(rows[0].TableRowProperties.GetFirstChild<TableHeader>(), Is.Not.Null);
            Assert.That(rows[0].Descendants<Run>().First().RunProperties.Bold, Is.Not.Null);

            var row1 = rows[1].Elements<TableCell>().ToList();
            Assert.That(row1, Has.Count.EqualTo(2));
            Assert.That(row1[0].TableCellProperties.VerticalMerge.Val.Value, Is.EqualTo(MergedCellValues.Restart));
            Assert.That(row1[1].TableCellProperties.GridSpan.Val.Value, Is.EqualTo(2));

            var row2 = rows[2].Elements<TableCell>().ToList();
            Assert.That(row2, Has.Count.EqualTo(3));
            Assert.That(row2[0].TableCellProperties.VerticalMerge, Is.Not.Null);
            Assert.That(row2[0].TableCellProperties.VerticalMerge.Val, Is.Null, "continuation of the merged cell");
            Assert.That(row2[2].TableCellProperties.Shading.Fill.Value, Is.EqualTo("FFFF00"));

            Assert.That(rows[3].Elements<TableCell>().Count(), Is.EqualTo(3), "short rows are padded");
            Assert.That(table.Descendants<TableCell>().All(c => c.LastChild is Paragraph));
        }

        [Test]
        public void TableInsideTemplateTableCell()
        {
            var body = Render(new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn { Width = "5000" }),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("{{ds}:html}")))))),
                "<table><tr><td>inner</td></tr></table>");

            var cell = body.Elements<Table>().Single().Elements<TableRow>().Single().Elements<TableCell>().Single();
            Assert.That(cell.Elements<Table>().Count(), Is.EqualTo(1));
            Assert.That(cell.LastChild, Is.TypeOf<Paragraph>());
        }

        [Test]
        public void Hyperlinks()
        {
            var html = "<a href=\"https://example.com/a?b=1\">external</a> <a href=\"#anchor\">internal</a> <a href=\"javascript:alert(1)\">script</a>";
            var (body, document) = RenderDocument("{{ds}:html}", html);

            var links = body.Descendants<Hyperlink>().ToList();
            Assert.That(links, Has.Count.EqualTo(2));
            Assert.That(links[0].InnerText, Is.EqualTo("external"));
            var relationship = document.MainDocumentPart.HyperlinkRelationships.Single(r => r.Id == links[0].Id);
            Assert.That(relationship.Uri.ToString(), Is.EqualTo("https://example.com/a?b=1"));
            Assert.That(links[1].Anchor.Value, Is.EqualTo("anchor"));
            Assert.That(body.InnerText, Does.Contain("script"), "a javascript link is rendered as plain text");
        }

        [Test]
        public void Base64Images()
        {
            var html = $"<p><img src=\"data:image/png;base64,{PngBase64}\" width=\"100\"> <img src=\"data:image/png;base64,{PngBase64}\" alt=\"second\"></p><img src=\"https://example.com/x.png\" alt=\"remote\">";
            var body = Render("{{ds}:html}", html, registerImages: true);

            var drawings = body.Descendants<Drawing>().ToList();
            Assert.That(drawings, Has.Count.EqualTo(2));
            var extent = drawings[0].Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().Single();
            Assert.That(extent.Cx.Value, Is.EqualTo(100 * 9525));
            Assert.That(extent.Cy.Value, Is.EqualTo(100 * 9525), "height follows the aspect ratio");
            var ids = drawings.Select(d => d.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().Single().Id.Value).ToList();
            Assert.That(ids.Distinct().Count(), Is.EqualTo(2), "every image needs a unique drawing id");
            Assert.That(body.InnerText, Does.Contain("remote"), "external images are not downloaded - alt text is used");
        }

        [Test]
        public void PlaceholdersInsideHtmlAreReplaced()
        {
            var body = Render("{{ds.Text}:html}", new { Text = "<p>Hello <b>{{ds.Name}}</b></p>", Name = "World" });

            Assert.That(body.InnerText, Is.EqualTo("Hello World"));
        }

        [Test]
        public void PlaceholdersInsideHtmlCanBeDisabled()
        {
            var body = Render("{{ds.Text}:html}", new { Text = "<p>Hello {{ds.Name}}</p>", Name = "World" },
                configuration: new HtmlFormatterConfiguration { ReplacePlaceholdersInHtml = false });

            Assert.That(body.InnerText, Is.EqualTo("Hello {{ds.Name}}"));
        }

        [Test]
        public void ReplacesBuiltInAltChunkFormatter()
        {
            var withModule = Render("{{ds}:html}", "<p>x</p>");
            Assert.That(withModule.Descendants<AltChunk>(), Is.Empty);

            var (withoutModule, _) = RenderDocument(new Body(new Paragraph(new Run(new Text("{{ds}:html}")))), "<p>x</p>", registerHtml: false);
            Assert.That(withoutModule.Descendants<AltChunk>().Count(), Is.EqualTo(1));
        }

        [Test]
        public void MalformedHtmlDoesNotThrow()
        {
            var body = Render("{{ds}:html}", "<p>unclosed <b>bold <i>both</p><li>stray item<table><td>cell");

            Assert.That(body.InnerText, Does.Contain("unclosed bold both"));
            Assert.That(body.InnerText, Does.Contain("cell"));
        }

        [Test]
        public void ScriptsAndStylesAreIgnored()
        {
            var body = Render("{{ds}:html}", "<style>p{color:red}</style><script>alert(1)</script><p>visible</p><p style=\"display:none\">hidden</p>");

            Assert.That(body.InnerText, Is.EqualTo("visible"));
        }

        [Test]
        public void BlockquotePreAndHorizontalRule()
        {
            var body = Render("{{ds}:html}", "<blockquote>quote</blockquote><hr><pre>line1\n  line2</pre>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs[0].ParagraphProperties.ParagraphBorders.LeftBorder, Is.Not.Null);
            Assert.That(paragraphs[1].ParagraphProperties.ParagraphBorders.BottomBorder, Is.Not.Null);
            var pre = paragraphs[2];
            Assert.That(pre.Descendants<Break>().Count(), Is.EqualTo(1));
            Assert.That(pre.Descendants<Text>().Last().Text, Is.EqualTo("  line2"));
            Assert.That(pre.Descendants<RunFonts>().First().Ascii.Value, Is.EqualTo("Courier New"));
        }

        [Test]
        public void BreakInsideAvoidKeepsEntryOnOnePage()
        {
            var html = """
                       <div style="break-inside: avoid"><p><b>App one</b></p><p>line 1<br>line 2</p><p>line 3</p></div>
                       <div style="page-break-inside: avoid"><p><b>App two</b></p><p>description</p></div>
                       <p>unrelated</p>
                       """;
            var body = Render("{{ds}:html}", html);

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo((string[])["App one", "line 1line 2", "line 3", "App two", "description", "unrelated"]));
            bool KeepNext(int i)
            {
                return paragraphs[i].ParagraphProperties?.KeepNext != null;
            }

            bool KeepLines(int i)
            {
                return paragraphs[i].ParagraphProperties?.KeepLines != null;
            }

            // the first paragraph is merged into the template paragraph - its keep options must survive
            Assert.That(KeepNext(0) && KeepLines(0), Is.True);
            Assert.That(KeepNext(1) && KeepLines(1), Is.True);
            Assert.That(!KeepNext(2) && KeepLines(2), Is.True, "the last paragraph of an entry ends the chain");
            Assert.That(KeepNext(3) && KeepLines(3), Is.True);
            Assert.That(!KeepNext(4) && KeepLines(4), Is.True);
            Assert.That(!KeepNext(5) && !KeepLines(5), Is.True, "content outside the entries is unaffected");
        }

        [Test]
        public void BreakAfterAvoidKeepsParagraphWithNext()
        {
            var body = Render("{{ds}:html}", "<p style=\"break-after: avoid\">title</p><p>text</p>");

            var paragraphs = body.Elements<Paragraph>().ToList();
            Assert.That(paragraphs[0].ParagraphProperties.KeepNext, Is.Not.Null);
            Assert.That(paragraphs[0].ParagraphProperties.KeepLines, Is.Null);
            Assert.That(paragraphs[1].ParagraphProperties?.KeepNext, Is.Null);
        }

        [Test]
        public void BreakInsideAvoidOnTableAndListItem()
        {
            var html = """
                       <table style="break-inside: avoid"><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>
                       <ul><li style="break-inside: avoid"><p>item</p><p>more</p></li><li>other</li></ul>
                       """;
            var body = Render("{{ds}:html}", html);

            var rows = body.Elements<Table>().Single().Elements<TableRow>().ToList();
            Assert.That(rows.All(r => r.TableRowProperties?.GetFirstChild<CantSplit>() != null), "rows must not split");
            Assert.That(rows[0].Descendants<Paragraph>().All(p => p.ParagraphProperties.KeepNext != null), "all rows but the last keep with the next row");
            Assert.That(rows[1].Descendants<Paragraph>().All(p => p.ParagraphProperties.KeepNext == null), "the table is not chained to the following content");

            var listParagraphs = body.Elements<Paragraph>().ToList();
            Assert.That(listParagraphs.Select(p => p.InnerText), Is.EqualTo((string[])["item", "more", "other"]));
            Assert.That(listParagraphs[0].ParagraphProperties.KeepNext, Is.Not.Null);
            Assert.That(listParagraphs[1].ParagraphProperties.KeepNext, Is.Null);
            Assert.That(listParagraphs[1].ParagraphProperties.KeepLines, Is.Not.Null);
            Assert.That(listParagraphs[2].ParagraphProperties.KeepLines, Is.Null);
        }

        [Test]
        public void EmptyHtmlRemovesPlaceholder()
        {
            var body = Render("before {{ds}:html} after", "");

            Assert.That(body.InnerText, Is.EqualTo("before  after"));
        }

        [Test]
        public void HtmlInHeader()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                headerPart.Header = new Header(new Paragraph(new Run(new Text("{{ds}:html}"))));
                var headerId = mainPart.GetIdOfPart(headerPart);
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("body"))),
                    new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId })));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new HtmlFormatter());
            docTemplate.BindModel("ds", "<b>Header</b> <a href=\"https://example.com\">link</a>");
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var document = WordprocessingDocument.Open(result, false);
            var header = document.MainDocumentPart.HeaderParts.Single();
            Assert.That(header.Header.InnerText, Is.EqualTo("Header link"));
            var link = header.Header.Descendants<Hyperlink>().Single();
            Assert.That(header.HyperlinkRelationships.Single().Id, Is.EqualTo(link.Id.Value));
        }

        [Test]
        public void ExistingSimpleHtmlTemplate()
        {
            var docTemplate = DocxTemplate.Open("Resources/SimpleHtmlRendering.docx");
            docTemplate.RegisterFormatter(new HtmlFormatter());
            var html = "<h1>The Main Languages of the Web</h1><p>HTML is the standard markup language.</p><hr><p>CSS describes how HTML elements are displayed.</p><ul><li>one</li><li>two</li></ul>";
            docTemplate.BindModel("ds", new { CLAUSES = html });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var document = WordprocessingDocument.Open(result, false);
            var text = document.MainDocumentPart.Document.Body.InnerText;
            Assert.That(text, Does.Contain("The Main Languages of the Web"));
            Assert.That(text, Does.Contain("two"));
            Assert.That(document.MainDocumentPart.Document.Body.Descendants<AltChunk>(), Is.Empty);
        }

        private static int? Level(Paragraph paragraph)
        {
            return paragraph.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value;
        }

        private static int? NumId(Paragraph paragraph)
        {
            return paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        }

        private static Body Render(string templateText, object model, bool registerImages = false, Styles styles = null, HtmlFormatterConfiguration configuration = null)
        {
            return RenderDocument(templateText, model, registerImages, styles, configuration).Body;
        }

        private static Body Render(OpenXmlElement templateContent, object model)
        {
            return RenderDocument(new Body(templateContent), model).Body;
        }

        private static (Body Body, WordprocessingDocument Document) RenderDocument(string templateText, object model, bool registerImages = false, Styles styles = null, HtmlFormatterConfiguration configuration = null)
        {
            return RenderDocument(new Body(new Paragraph(new Run(new Text(templateText) { Space = SpaceProcessingModeValues.Preserve }))), model, registerImages, styles, configuration);
        }

        private static (Body Body, WordprocessingDocument Document) RenderDocument(Body templateBody, object model, bool registerImages = false, Styles styles = null, HtmlFormatterConfiguration configuration = null, bool registerHtml = true)
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(templateBody);
                if (styles != null)
                {
                    mainPart.AddNewPart<StyleDefinitionsPart>().Styles = styles;
                }
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            if (registerImages)
            {
                docTemplate.RegisterFormatter(new ImageFormatter());
            }
            if (registerHtml)
            {
                docTemplate.RegisterFormatter(new HtmlFormatter(configuration));
            }
            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            return (document.MainDocumentPart.Document.Body, document);
        }
    }
}

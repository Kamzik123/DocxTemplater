using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Pdf.Test
{
    internal class PdfFormatterSectionTest : PdfTestBase
    {
        private static readonly byte[] A3Landscape = FakePdfRenderer.Pdf((FakePdfRenderer.A3Height, FakePdfRenderer.A3Width));

        [Test]
        public void A3LandscapePlanInA4DocumentGetsASectionOfItsOwn()
        {
            var (body, document) = Render(new Body(P("Appendix B"), P("{{ds.Plan}:pdf(fit:section)}"), P("Appendix C"), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Appendix B", "A3L: [plan]", "A4: Appendix C"]));
            var planSection = GetSections(body)[1].Section;
            var pageSize = planSection.GetFirstChild<PageSize>();
            Assert.That(pageSize.Width.Value, Is.EqualTo(23811));
            Assert.That(pageSize.Height.Value, Is.EqualTo(16838));
            Assert.That(pageSize.Orient.Value, Is.EqualTo(PageOrientationValues.Landscape));
            var margin = planSection.GetFirstChild<PageMargin>();
            Assert.That(new long[] { margin.Top.Value, margin.Bottom.Value, margin.Left.Value, margin.Right.Value, margin.Header.Value, margin.Footer.Value }, Is.All.EqualTo(0));
            Assert.That(planSection.GetFirstChild<SectionType>().Val.Value, Is.EqualTo(SectionMarkValues.NextPage));

            // empty header and footer, so the document's header is not drawn over the plan
            var header = (HeaderPart)document.MainDocumentPart.GetPartById(planSection.Elements<HeaderReference>().Single().Id.Value);
            var footer = (FooterPart)document.MainDocumentPart.GetPartById(planSection.Elements<FooterReference>().Single().Id.Value);
            Assert.That(header.Header.InnerText, Is.Empty);
            Assert.That(footer.Footer.InnerText, Is.Empty);
        }

        [Test]
        public void PlanPictureHasTheExactPageSizeAndIsAnchoredAtThePageOrigin()
        {
            var (body, _) = Render(new Body(P("{{ds.Plan}:pdf(fit:section)}"), P("after"), A4Section()), new { Plan = A3Landscape });

            var anchor = body.Descendants<DW.Anchor>().Single();
            Assert.That(anchor.Extent.Cx.Value, Is.EqualTo((long)Math.Round(FakePdfRenderer.A3Height * 12700)));
            Assert.That(anchor.Extent.Cy.Value, Is.EqualTo((long)Math.Round(FakePdfRenderer.A3Width * 12700)));
            Assert.That(anchor.HorizontalPosition.RelativeFrom.Value, Is.EqualTo(DW.HorizontalRelativePositionValues.Page));
            Assert.That(anchor.HorizontalPosition.PositionOffset.Text, Is.EqualTo("0"));
            Assert.That(anchor.VerticalPosition.RelativeFrom.Value, Is.EqualTo(DW.VerticalRelativePositionValues.Page));
            Assert.That(anchor.VerticalPosition.PositionOffset.Text, Is.EqualTo("0"));
            Assert.That(anchor.GetFirstChild<DW.WrapNone>(), Is.Not.Null);
            Assert.That(anchor.Descendants<DW.DocProperties>().Single().Description.Value, Is.EqualTo("Document, page 1 of 1"));

            var spacing = anchor.GetFirstAncestor<Paragraph>().ParagraphProperties.SpacingBetweenLines;
            Assert.That(spacing.LineRule.Value, Is.EqualTo(LineSpacingRuleValues.Exact));
            Assert.That(spacing.Line.Value, Is.EqualTo("20"));
        }

        [Test]
        public void PlanAtTheEndOfTheDocumentReplacesTheLastSection()
        {
            var (body, _) = Render(new Body(P("Appendix"), P("{{ds.Plan}:pdf(fit:section)}"), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Appendix", "A3L: [plan]"]), "no empty A4 page after the plan");
        }

        [Test]
        public void PlanFollowedOnlyByEmptyParagraphsReplacesTheLastSection()
        {
            var (body, _) = Render(new Body(P("Appendix"), P("{{ds.Plan}:pdf(fit:section)}"), new Paragraph(), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Appendix", "A3L: [plan] | "]));
        }

        [Test]
        public void PlanAtTheStartOfTheDocumentNeedsNoSectionBefore()
        {
            var (body, _) = Render(new Body(P("{{ds.Plan}:pdf(fit:section)}"), P("text"), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A3L: [plan]", "A4: text"]));
        }

        [Test]
        public void TextBeforeAndAfterThePlaceholderInTheSameParagraph()
        {
            var (body, _) = Render(new Body(P("Before {{ds.Plan}:pdf(fit:section)} After"), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Before ", "A3L: [plan]", "A4:  After"]));
        }

        [Test]
        public void PlanAfterATableGetsAMinimalParagraphToEndTheSection()
        {
            var table = new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "2000" }), new TableRow(new TableCell(P("cell"))));
            var (body, _) = Render(new Body(table, P("{{ds.Plan}:pdf(fit:section)}"), P("text"), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: [table] | ", "A3L: [plan]", "A4: text"]));
            var closing = (Paragraph)body.ChildElements[1];
            Assert.That(closing.ParagraphProperties.SpacingBetweenLines.Line.Value, Is.EqualTo("20"), "1 pt paragraph");
        }

        [Test]
        public void PlaceholderParagraphEndingASection()
        {
            var firstSection = A4Section(new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdHeader" });
            var landscape = new SectionProperties(new PageSize { Width = (uint)A4HeightTwips, Height = (uint)A4WidthTwips, Orient = PageOrientationValues.Landscape });
            var (body, _) = Render(new Body(P("Intro"), PWithSection("{{ds.Plan}:pdf(fit:section)}", firstSection), P("landscape part"), landscape),
                new { Plan = A3Landscape },
                setup: d => d.MainDocumentPart.AddNewPart<HeaderPart>("rIdHeader").Header = new Header(P("Header")));

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro", "A3L: [plan]", "A4L: landscape part"]));
            var introSection = GetSections(body)[0].Section;
            Assert.That(introSection.GetFirstChild<HeaderReference>().Id.Value, Is.EqualTo("rIdHeader"), "the copy keeps the header");
        }

        [Test]
        public void MixedPageSizesAreGrouped()
        {
            var pdf = FakePdfRenderer.Pdf(
                (FakePdfRenderer.A4Width, FakePdfRenderer.A4Height),
                (FakePdfRenderer.A3Height, FakePdfRenderer.A3Width),
                (FakePdfRenderer.A3Height, FakePdfRenderer.A3Width),
                (FakePdfRenderer.A4Width, FakePdfRenderer.A4Height));
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section)}"), P("End"), A4Section()), new { Plan = pdf });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro | [pdf]", "A3L: [plan] | [plan]", "A4: [pdf] | End"]));
            var planParagraphs = GetSections(body)[1].Content.Count;
            var secondPlanPage = body.Descendants<DW.Anchor>().ElementAt(1).GetFirstAncestor<Paragraph>();
            Assert.That(secondPlanPage.ParagraphProperties.PageBreakBefore, Is.Not.Null, "each plan page starts a page");
            Assert.That(planParagraphs, Is.EqualTo(2));
        }

        [Test]
        public void SamePageSizeCanGetASectionToo()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section, same:section)}"), P("End"), A4Section()), new { Plan = FakePdfRenderer.A4Pages(1) });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro", "A4: [plan]", "A4: End"]));
            Assert.That(GetSections(body)[1].Section.GetFirstChild<PageMargin>().Left.Value, Is.Zero);
        }

        [Test]
        public void OnlyPagesOfTheDocumentSizeNeedNoSection()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section)}"), P("End"), A4Section()), new { Plan = FakePdfRenderer.A4Pages(2) });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Intro | [pdf] | [pdf] | End"]));
        }

        [Test]
        public void FirstPageHeaderAndPageNumberRestartStayWithTheFirstPart()
        {
            var section = A4Section(new TitlePage(), new PageNumberType { Start = 1 });
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section)}"), P("End"), section), new { Plan = A3Landscape });

            var sections = GetSections(body);
            Assert.That(sections[0].Section.GetFirstChild<TitlePage>(), Is.Not.Null);
            Assert.That(sections[0].Section.GetFirstChild<PageNumberType>().Start.Value, Is.EqualTo(1));
            Assert.That(sections[1].Section.GetFirstChild<PageNumberType>(), Is.Null, "numbering continues through the plan");
            Assert.That(sections[2].Section.GetFirstChild<TitlePage>(), Is.Null, "no first page header after the plan");
            Assert.That(sections[2].Section.GetFirstChild<PageNumberType>()?.Start, Is.Null, "no restart after the plan");
        }

        [Test]
        public void PageNumberRestartMovesToThePlanIfItStartsTheSection()
        {
            var section = A4Section(new PageNumberType { Start = 5 });
            var (body, _) = Render(new Body(P("{{ds.Plan}:pdf(fit:section)}"), P("End"), section), new { Plan = A3Landscape });

            var sections = GetSections(body);
            Assert.That(sections[0].Section.GetFirstChild<PageNumberType>().Start.Value, Is.EqualTo(5));
            Assert.That(sections[1].Section.GetFirstChild<PageNumberType>()?.Start, Is.Null);
        }

        [Test]
        public void EvenPageHeadersAreReplacedToo()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section)}"), A4Section()), new { Plan = A3Landscape },
                setup: d => d.MainDocumentPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings(new EvenAndOddHeaders()));

            var planSection = GetSections(body)[1].Section;
            Assert.That(planSection.Elements<HeaderReference>().Select(x => x.Type.Value), Is.EquivalentTo([HeaderFooterValues.Default, HeaderFooterValues.Even]));
            Assert.That(planSection.Elements<FooterReference>().Select(x => x.Type.Value), Is.EquivalentTo([HeaderFooterValues.Default, HeaderFooterValues.Even]));
        }

        [Test]
        public void EmptyHeaderAndFooterAreSharedByAllPlans()
        {
            var (body, document) = Render(new Body(P("{{ds.Plan}:pdf(fit:section)}"), P("x"), P("{{ds.Plan}:pdf(fit:section)}"), P("y"), A4Section()), new { Plan = A3Landscape });

            Assert.That(document.MainDocumentPart.HeaderParts.Count(), Is.EqualTo(1));
            Assert.That(document.MainDocumentPart.ImageParts.Count(), Is.EqualTo(1), "the picture is stored once");
            Assert.That(DescribeSections(body), Is.EqualTo(["A3L: [plan]", "A4: x", "A3L: [plan]", "A4: y"]));
        }

        [Test]
        public void InheritedHeadersAreKept()
        {
            var section = A4Section(new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdHeader" });
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section, hf:inherit)}"), section), new { Plan = A3Landscape },
                setup: d => d.MainDocumentPart.AddNewPart<HeaderPart>("rIdHeader").Header = new Header(P("Header")));

            var planSection = GetSections(body)[1].Section;
            Assert.That(planSection.GetFirstChild<HeaderReference>().Id.Value, Is.EqualTo("rIdHeader"));
            Assert.That(planSection.GetFirstChild<PageMargin>().Header.Value, Is.EqualTo(708));
        }

        [Test]
        public void MarginScalesThePlanIntoTheRemainingArea()
        {
            var (body, _) = Render(new Body(P("{{ds.Plan}:pdf(fit:section, margin:10mm)}"), P("x"), A4Section()), new { Plan = A3Landscape });

            var planSection = GetSections(body)[0].Section;
            Assert.That(planSection.GetFirstChild<PageMargin>().Left.Value, Is.EqualTo(566));
            var anchor = body.Descendants<DW.Anchor>().Single();
            Assert.That(anchor.Extent.Cx.Value, Is.LessThanOrEqualTo((23811 - (2 * 566)) * 635));
            Assert.That(long.Parse(anchor.HorizontalPosition.PositionOffset.Text, CultureInfo.InvariantCulture), Is.GreaterThanOrEqualTo(566 * 635));
        }

        [Test]
        public void OversizedPagesAreScaledWithAnOverlay()
        {
            var a1Landscape = FakePdfRenderer.Pdf((2383.94, 1683.78));
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section)}"), A4Section()), new { Plan = new PdfFile(a1Landscape, "Ground floor") });

            var planSection = GetSections(body)[1].Section;
            var pageSize = planSection.GetFirstChild<PageSize>();
            Assert.That(pageSize.Width.Value, Is.EqualTo(31680));
            Assert.That(pageSize.Height.Value, Is.LessThan(31680));
            Assert.That((double)pageSize.Width.Value / pageSize.Height.Value, Is.EqualTo(2383.94 / 1683.78).Within(0.001), "aspect ratio is kept");
            Assert.That(pageSize.Orient.Value, Is.EqualTo(PageOrientationValues.Landscape));

            var plan = body.Descendants<DW.Anchor>().First();
            Assert.That(plan.Extent.Cx.Value, Is.EqualTo(31680 * 635).Within(635), "the picture fills the scaled page");
            var overlay = body.Descendants<DW.Anchor>().Last();
            Assert.That(overlay.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>(), Is.Empty);
            Assert.That(overlay.HorizontalPosition.HorizontalAlignment.Text, Is.EqualTo("center"));
            Assert.That(overlay.RelativeHeight.Value, Is.GreaterThan(plan.RelativeHeight.Value), "drawn on top of the plan");
            var lines = overlay.Descendants<Paragraph>().Select(x => x.InnerText).ToList();
            Assert.That(lines[0], Is.EqualTo("PLACEHOLDER - replace with original"));
            Assert.That(lines[1], Does.StartWith("Target size: A1 landscape, 841 × 594 mm (shown at "));
            Assert.That(lines[2], Is.EqualTo("Plan: Ground floor, page 1 of 1"));
            Assert.That(plan.Descendants<DW.DocProperties>().Single().Description.Value, Does.Contain("PLACEHOLDER"));
        }

        [Test]
        public void OversizedPagesCanBeRejected()
        {
            var a1Landscape = FakePdfRenderer.Pdf((2383.94, 1683.78));
            var exception = Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(P("{{ds.Plan}:pdf(fit:section, oversize:error)}"), A4Section()), new { Plan = a1Landscape }));
            Assert.That(exception.InnerException.Message, Does.Contain("841 × 594 mm"));
        }

        [Test]
        public void OverlayOnEveryPlanPageWithTranslatedTexts()
        {
            var configuration = new PdfFormatterConfiguration
            {
                OverlayTitle = "PLATZHALTER - durch Original ersetzen",
                OverlaySizeLine = "Zielformat: {sheet} {orientation}, {width} × {height} mm (dargestellt mit {scale})",
                OverlaySourceLine = "",
                OrientationLandscape = "quer",
                OrientationPortrait = "hoch",
            };
            var (body, _) = Render(new Body(P("{{ds.Plan}:pdf(fit:section, overlay:always)}"), A4Section()), new { Plan = A3Landscape }, configuration: configuration);

            var lines = body.Descendants<DW.Anchor>().Last().Descendants<Paragraph>().Select(x => x.InnerText).ToList();
            Assert.That(lines, Has.Count.EqualTo(2), "the empty source line is dropped");
            Assert.That(lines[1], Does.StartWith("Zielformat: A3 quer, 420 × 297 mm (dargestellt mit 100"));
        }

        [Test]
        public void NoOverlayIfDisabled()
        {
            var a1 = FakePdfRenderer.Pdf((1683.78, 2383.94));
            var (body, _) = Render(new Body(P("{{ds.Plan}:pdf(fit:section, overlay:never)}"), A4Section()), new { Plan = a1 });
            Assert.That(body.Descendants<DW.Anchor>().Count(), Is.EqualTo(1));
        }

        [Test]
        public void NotPossibleInTablesAndHeaders()
        {
            var table = new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "2000" }), new TableRow(new TableCell(P("{{ds.Plan}:pdf(fit:section)}"))));
            var renderer = new FakePdfRenderer();
            var exception = Assert.Throws<OpenXmlTemplateException>(() => Render(new Body(table, P(""), A4Section()), new { Plan = A3Landscape }, renderer));
            Assert.That(exception.InnerException.Message, Does.Contain("fit:section"));
            Assert.That(renderer.GetPageSizesCalls, Is.Zero, "fails before rendering");

            Assert.Throws<OpenXmlTemplateException>(() => Render(
                new Body(P("x"), A4Section(new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdHeader" })),
                new { Plan = A3Landscape },
                setup: d => d.MainDocumentPart.AddNewPart<HeaderPart>("rIdHeader").Header = new Header(P("{{ds.Plan}:pdf(fit:section)}"))));
        }

        [Test]
        public void PlansInALoopAtTheEndOfTheDocument()
        {
            var (body, _) = Render(new Body(P("Plans"), P("{{#ds.Plans}}"), P("{{.}:pdf(fit:section)}"), P("{{/ds.Plans}}"), A4Section()),
                new { Plans = new[] { A3Landscape, A3Landscape, A3Landscape } });

            // the empty paragraphs are the 1 pt paragraphs that end the sections (loop items are rendered last to first)
            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Plans | ", "A3L: [plan] | ", "A3L: [plan] | ", "A3L: [plan]"]));
        }

        [Test]
        public void PlansInALoopWithTitles()
        {
            var (body, _) = Render(
                new Body(P("Plans"), P("{{#ds.Plans}}"), P("Plan {{.Name}}"), P("{{.Pdf}:pdf(fit:section)}"), P("{{/ds.Plans}}"), P("End"), A4Section()),
                new { Plans = new[] { new { Name = "A", Pdf = A3Landscape }, new { Name = "B", Pdf = A3Landscape } } });

            var sections = DescribeSections(body);
            Assert.That(sections, Is.EqualTo(["A4: Plans | Plan A", "A3L: [plan]", "A4: Plan B", "A3L: [plan]", "A4: End"]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void PlansInALoopFollowedByText()
        {
            var (body, _) = Render(new Body(P("Plans"), P("{{#ds.Plans}}"), P("{{.}:pdf(fit:section)}"), P("{{/ds.Plans}}"), P("End"), A4Section()),
                new { Plans = new[] { A3Landscape, A3Landscape } });

            var sections = DescribeSections(body);
            Assert.That(sections.Where(x => x.StartsWith("A3L", StringComparison.Ordinal)), Has.All.StartsWith("A3L: [plan]"), string.Join(Environment.NewLine, sections));
            Assert.That(sections.Count(x => x.StartsWith("A3L", StringComparison.Ordinal)), Is.EqualTo(2));
            Assert.That(sections[^1], Does.EndWith("End"));
            Assert.That(sections.Where(x => x.StartsWith("A4", StringComparison.Ordinal)).SelectMany(x => x.Split(" | ")).Any(x => x.Contains("[plan]", StringComparison.Ordinal)), Is.False);
        }

        [Test]
        public void PlansInAnInlineLoopFollowedByText()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{#ds.Plans}}{{.}:pdf(fit:section)}{{/ds.Plans}}"), P("End"), A4Section()),
                new { Plans = new[] { A3Landscape, A3Landscape, A3Landscape } });

            // the empty entries: 1 pt paragraphs closing the plan sections, the rest of the loop paragraph before "End"
            var sections = DescribeSections(body);
            Assert.That(sections, Is.EqualTo(["A4: Intro | ", "A3L: [plan] | ", "A3L: [plan] | ", "A3L: [plan]", "A4:  | End"]), string.Join(Environment.NewLine, sections));
        }

        [Test]
        public void DocumentWithPlansInsertedAsSubTemplate()
        {
            var plans = FakePdfRenderer.Pdf((FakePdfRenderer.A3Height, FakePdfRenderer.A3Width), (FakePdfRenderer.A3Height, FakePdfRenderer.A3Width));
            var (_, appendix) = Render(new Body(P("Intro"), P("{{ds.Plans}:pdf(fit:section)}"), P("After plans"), A4Section()), new { Plans = plans });
            byte[] appendixBytes;
            using (var stream = new MemoryStream())
            {
                using (appendix.Clone(stream))
                {
                }
                appendixBytes = stream.ToArray();
            }
            appendix.Dispose();

            var hostSection = A4Section(new PageNumberType { Start = 1 }, new TitlePage());
            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), PWithSection("", hostSection), P("Back cover"), A4Section()), new { Appendix = appendixBytes });

            var sections = GetSections(body);
            var described = DescribeSections(body);
            Assert.That(described, Is.EqualTo(["A4: Cover | Intro", "A3L: [plan] | [plan]", "A4: After plans | ", "A4: Back cover"]), string.Join(Environment.NewLine, described));
            Assert.That(sections.Select(x => x.Section.GetFirstChild<TitlePage>() != null), Is.EqualTo([true, false, false, false]), "only the start of the host section has a different first page");
            Assert.That(sections.Select(x => x.Section.GetFirstChild<PageNumberType>()?.Start?.Value), Is.EqualTo(new int?[] { 1, null, null, null }));
        }

        [Test]
        public void FinalDocumentKeepsThePageSetupOfTheDocumentAsLastSection()
        {
            var (body, _) = Render(new Body(P("Appendix"), P("{{ds.Plan}:pdf(fit:section, final:document)}"), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Appendix", "A3L: [plan]", "A4: "]));
            Assert.That(body.LastChild, Is.InstanceOf<SectionProperties>());
        }

        [Test]
        public void FinalDocumentDoesNotChangePagesFollowedByContent()
        {
            var (body, _) = Render(new Body(P("Appendix"), P("{{ds.Plan}:pdf(fit:section, final:document)}"), P("After"), A4Section()), new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Appendix", "A3L: [plan]", "A4: After"]));
        }

        [Test]
        public void FinalDocumentDoesNotAddAnEmptySectionInsideTheDocument()
        {
            var (body, _) = Render(new Body(P("Appendix"), P("{{ds.Plan}:pdf(fit:section, final:document)}"), PWithSection("", A4Section()), P("Next"), A4Section()),
                new { Plan = A3Landscape });

            Assert.That(DescribeSections(body), Is.EqualTo(["A4: Appendix", "A3L: [plan] | ", "A4: Next"]));
        }

        [Test]
        public void FinalDocumentWithPlansInALoopAtTheEndOfTheDocument()
        {
            var (body, _) = Render(new Body(P("Plans"), P("{{#ds.Plans}}"), P("{{.}:pdf(fit:section, final:document)}"), P("{{/ds.Plans}}"), A4Section()),
                new { Plans = new[] { A3Landscape, A3Landscape } });

            var sections = DescribeSections(body);
            Assert.That(sections.Count(x => x.StartsWith("A3L: [plan]", StringComparison.Ordinal)), Is.EqualTo(2), string.Join(Environment.NewLine, sections));
            Assert.That(sections[^1], Does.StartWith("A4"), string.Join(Environment.NewLine, sections));
            Assert.That(sections[^1], Does.Not.Contain("[plan]"));
        }

        [TestCase("fit:section, same:section, overlay:always")]
        [TestCase("content:blank")]
        public void DocumentEndingWithPlansInsertedAsSubTemplateKeepsThePlanSection(string arguments)
        {
            var (_, appendix) = Render(new Body(P("Intro"), P($"{{{{ds.Plan}}:pdf({arguments}, final:document, marker:'[[PDF|%page%]]')}}"), A4Section()),
                new { Plan = A3Landscape });
            byte[] appendixBytes;
            using (var stream = new MemoryStream())
            {
                using (appendix.Clone(stream))
                {
                }
                appendixBytes = stream.ToArray();
            }
            appendix.Dispose();

            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), P("Back cover"), A4Section()), new { Appendix = appendixBytes });

            var sections = GetSections(body);
            var described = DescribeSections(body);
            Assert.That(described.Select(x => x.Split(':')[0]), Is.EqualTo(["A4", "A3L", "A4"]), string.Join(Environment.NewLine, described));
            Assert.That(described[0], Does.EndWith("Cover | Intro"));
            Assert.That(described[2], Does.EndWith("Back cover"));
            Assert.That(sections[1].Content.Single(), Does.Contain("[[PDF|1]]"));
            var margin = sections[1].Section.GetFirstChild<PageMargin>();
            Assert.That(new long[] { margin.Top.Value, margin.Bottom.Value, margin.Left.Value, margin.Right.Value }, Is.All.EqualTo(0));
        }

        [Test]
        public void ConditionAfterThePlanCanStillAddContent()
        {
            var (body, _) = Render(new Body(P("Intro"), P("{{ds.Plan}:pdf(fit:section)}"), P("{?{ds.More}}"), P("More text"), P("{{/}}"), A4Section()),
                new { Plan = A3Landscape, More = true });

            var sections = DescribeSections(body);
            Assert.That(sections[1], Does.StartWith("A3L: [plan]"));
            Assert.That(sections[^1], Does.StartWith("A4").And.Contain("More text"), string.Join(Environment.NewLine, sections));
        }

        [TestCase("fit:section, same:section")]
        [TestCase("content:blank")]
        public void FinalDocumentInsertedBeforeEmptyHostSectionAddsNoPage(string arguments)
        {
            var appendix = RenderAppendix(arguments);
            var hostSection = A4Section(new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdFooter" });
            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), P(""), new Paragraph(new Run(new Break { Type = BreakValues.Page })),
                    PWithSection("", hostSection), P("Back cover"), A4Section()),
                new { Appendix = appendix },
                setup: AddFooters);

            var described = DescribeSections(body);
            Assert.That(described.Select(x => x.Split(':')[0]), Is.EqualTo(["A4", "A3L", "A4"]), string.Join(Environment.NewLine, described));
            Assert.That(described[0], Does.EndWith("Cover | Intro"));
            Assert.That(described[2], Is.EqualTo("A4: Back cover"));
            var backCover = GetSections(body)[2].Section;
            Assert.That(backCover.Elements<FooterReference>().Select(x => x.Id.Value), Is.EqualTo(["rIdFooter"]), "the back cover inherited the host footer from the removed section");
        }

        [Test]
        public void FinalDocumentInsertedBeforeHostContentKeepsTheHostSection()
        {
            var appendix = RenderAppendix("fit:section");
            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), P(""), P("After"), PWithSection("", A4Section()), P("Back cover"), A4Section()),
                new { Appendix = appendix });

            var described = DescribeSections(body);
            Assert.That(described.Select(x => x.Split(':')[0]), Is.EqualTo(["A4", "A3L", "A4", "A4"]), string.Join(Environment.NewLine, described));
            Assert.That(described[2], Does.Contain("After"));
            Assert.That(described[3], Is.EqualTo("A4: Back cover"));
        }

        [Test]
        public void FinalDocumentInsertedAtTheEndOfTheHostBecomesTheLastSection()
        {
            var appendix = RenderAppendix("fit:section");
            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), P(""), A4Section()), new { Appendix = appendix });

            var described = DescribeSections(body);
            Assert.That(described.Select(x => x.Split(':')[0]), Is.EqualTo(["A4", "A3L"]), string.Join(Environment.NewLine, described));
            Assert.That(body.LastChild, Is.InstanceOf<SectionProperties>());
            Assert.That(PageSizeName((SectionProperties)body.LastChild), Is.EqualTo("A3L"));
            Assert.That(body.Elements<Paragraph>().Last().ParagraphProperties?.SectionProperties, Is.Null);
        }

        [Test]
        public void BookmarksOfTheRemovedSectionKeepTheirPair()
        {
            var appendix = RenderAppendix("fit:section");
            var cover = P("Cover");
            cover.AppendChild(new BookmarkStart { Id = "7", Name = "Span" });
            var removed = new Paragraph(new BookmarkEnd { Id = "7" }, new BookmarkStart { Id = "8", Name = "Host" });
            var (body, _) = Render(new Body(cover, P("{{ds}:T('ds.Appendix')}"), removed, PWithSection("", A4Section()), P("Back cover"), new BookmarkEnd { Id = "8" }, A4Section()),
                new { Appendix = appendix });

            Assert.That(DescribeSections(body).Select(x => x.Split(':')[0]), Is.EqualTo(["A4", "A3L", "A4"]));
            var planParagraph = GetSections(body)[1].Section.Parent.Parent;
            Assert.That(planParagraph.Elements<BookmarkStart>().Select(x => x.Name.Value), Is.EqualTo(["Host"]));
            Assert.That(planParagraph.Elements<BookmarkEnd>().Count(), Is.EqualTo(1));
            var starts = body.Descendants<BookmarkStart>().ToList();
            Assert.That(starts.Select(x => x.Name.Value), Is.EquivalentTo(["Span", "Host"]));
            Assert.That(body.Descendants<BookmarkEnd>().Select(x => x.Id.Value), Is.EquivalentTo(starts.Select(x => x.Id.Value)));
        }

        [Test]
        public void HeaderAndFooterTypesOnlyTheRemovedSectionHadAreInherited()
        {
            var appendix = RenderAppendix("fit:section");
            var hostSection = A4Section(new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdFooter" }, new FooterReference { Type = HeaderFooterValues.Even, Id = "rIdEven" });
            var backSection = A4Section(new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdBack" });
            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), PWithSection("", hostSection), P("Back cover"), backSection),
                new { Appendix = appendix },
                setup: AddFooters);

            var backCover = GetSections(body)[^1].Section;
            Assert.That(backCover.Elements<FooterReference>().Select(x => (x.Type.Value, x.Id.Value)),
                Is.EqualTo([(HeaderFooterValues.Default, "rIdBack"), (HeaderFooterValues.Even, "rIdEven")]));
        }

        [Test]
        public void FinalPdfInsertedBeforeEmptyHostSectionIsUnchanged()
        {
            var appendix = RenderAppendix("fit:section", finalArgument: "final:pdf");
            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), P(""), PWithSection("", A4Section()), P("Back cover"), A4Section()),
                new { Appendix = appendix });

            var described = DescribeSections(body);
            // the plan section is the body sectPr of the inserted document and replaced by the host's: nothing is marked
            Assert.That(described.Select(x => x.Split(':')[0]), Is.EqualTo(["A4", "A4", "A4", "A4"]), string.Join(Environment.NewLine, described));
            Assert.That(described[^1], Is.EqualTo("A4: Back cover"));
        }

        [Test]
        public void ConditionRenderedEmptyAfterTheInsertedDocumentRemovesTheHostSection()
        {
            var appendix = RenderAppendix("fit:section");
            var (body, _) = Render(new Body(P("Cover"), P("{{ds}:T('ds.Appendix')}"), P("{?{ds.Show}}"), P("After"), P("{{/}}"), PWithSection("", A4Section()), P("Back cover"), A4Section()),
                new { Appendix = appendix, Show = false });

            var described = DescribeSections(body);
            Assert.That(described.Select(x => x.Split(':')[0]), Is.EqualTo(["A4", "A3L", "A4"]), string.Join(Environment.NewLine, described));
            Assert.That(described[2], Is.EqualTo("A4: Back cover"));
        }

        private static byte[] RenderAppendix(string arguments, string finalArgument = "final:document")
        {
            var (_, appendix) = Render(new Body(P("Intro"), P($"{{{{ds.Plan}}:pdf({arguments}, {finalArgument})}}"), A4Section()), new { Plan = A3Landscape });
            using var stream = new MemoryStream();
            using (appendix.Clone(stream))
            {
            }
            appendix.Dispose();
            return stream.ToArray();
        }

        private static void AddFooters(WordprocessingDocument document)
        {
            document.MainDocumentPart.AddNewPart<FooterPart>("rIdFooter").Footer = new Footer(P("Host footer"));
            document.MainDocumentPart.AddNewPart<FooterPart>("rIdEven").Footer = new Footer(P("Even footer"));
            document.MainDocumentPart.AddNewPart<FooterPart>("rIdBack").Footer = new Footer(P("Back footer"));
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// <c>fit:section</c>: every PDF page gets a page of exactly its own size. Consecutive pages of the same size
    /// share a section without margins (and by default without header and footer); the picture is anchored to the
    /// page at (0,0) and has the size of the page. The document around the pages keeps its own page setup.
    /// </summary>
    /// <remarks>
    /// A section's page setup is stored in the <c>w:sectPr</c> of its last paragraph (the last section in
    /// <c>w:body/w:sectPr</c>). So the content before the pages ends with a copy of the section properties that
    /// govern the placeholder, each group of pages ends with its own section properties, and the original section
    /// properties keep governing the content after the pages. If nothing follows the pages in their section, the
    /// last group takes the place of the original section properties instead - otherwise an empty page would follow.
    /// Loops render their items last to first, so a placeholder may be followed by pages of a later item that are
    /// already laid out: section properties created here are annotated with the page setup of the document.
    /// </remarks>
    internal sealed class SectionLayout
    {
        /// <summary>Page sizes within this tolerance (1 mm) count as equal.</summary>
        private const long SizeToleranceTwips = 57;

        private readonly PdfInsertContext m_context;
        private readonly PdfFormatterConfiguration m_configuration;

        private SectionLayout(PdfInsertContext context)
        {
            m_context = context;
            m_configuration = context.Configuration;
        }

        /// <summary>
        /// Throws if <c>fit:section</c> is not possible at the placeholder: section breaks only exist in the body.
        /// </summary>
        public static void EnsureSupported(Text target)
        {
            var paragraph = target.GetFirstAncestor<Paragraph>();
            if (paragraph?.Parent is not Body body || body.Parent is not Document)
            {
                throw new OpenXmlTemplateException(
                    "pdf fit:section is only possible in a paragraph of the document body - not in tables, text boxes, " +
                    "headers, footers or content controls, because these cannot hold section breaks");
            }
        }

        public static void Insert(Text target, IReadOnlyList<PlacedPage> pages, PdfInsertContext context)
        {
            EnsureSupported(target);
            new SectionLayout(context).InsertPages(target, pages);
        }

        private void InsertPages(Text target, IReadOnlyList<PlacedPage> pages)
        {
            var body = (Body)target.GetFirstAncestor<Paragraph>().Parent;
            var (before, paragraph, after) = PdfPlacement.SplitAround(target);

            // the section properties governing the placeholder and where they are stored
            var governing = (after ?? paragraph).GetSectionProperties();
            if (governing == null)
            {
                governing = new SectionProperties(
                    new PageSize { Width = 12240U, Height = 15840U },
                    new PageMargin { Top = 1440, Right = 1440U, Bottom = 1440, Left = 1440U, Header = 720U, Footer = 720U, Gutter = 0U });
                body.AppendChild(governing);
            }
            var holder = governing.Parent?.Parent as Paragraph; // null: body/sectPr
            var plan = governing.Annotation<PlanSectionAnnotation>();
            var followedByPlan = plan != null && plan.FirstParagraph?.Parent == body;
            // the page setup of the document (for pages laid out like fit:page and for the sections closing content)
            var documentSection = plan?.Original ?? governing;
            var originalCopy = (SectionProperties)documentSection.CloneNode(true);
            var continuedCopy = (SectionProperties)documentSection.CloneNode(true);
            OpenXmlHelper.RemoveSectionStartSettings(continuedCopy);

            var groups = CreateGroups(pages, documentSection);
            if (groups.All(x => !x.IsSection))
            {
                // every page has the size of the document pages - nothing to split
                InsertWithoutSections(paragraph, groups.SelectMany(x => x.Pages).ToList(), documentSection, before != null);
                return;
            }

            // (1) end the content before the pages with a copy of the governing section properties - pages of the
            // document size at the start stay in that section and end it themselves
            var sectionEmitted = groups[0].IsSection && CloseContentBefore(before, paragraph, originalCopy);

            // (2) the pages
            var output = new List<Paragraph>();
            SectionProperties lastPlanSection = null;
            Paragraph lastGroupFirstParagraph = null;
            for (var groupIndex = 0; groupIndex < groups.Count; groupIndex++)
            {
                var group = groups[groupIndex];
                var isLastGroup = groupIndex == groups.Count - 1;
                if (group.IsSection)
                {
                    var groupParagraphs = group.Pages.Select((page, index) => CreatePlanPage(page, group, index > 0)).ToList();
                    var sectionProperties = CreatePlanSectionProperties(group, documentSection, sectionEmitted ? null : originalCopy);
                    sectionProperties.AddAnnotation(new PlanSectionAnnotation((SectionProperties)continuedCopy.CloneNode(true), groupParagraphs[0]));
                    groupParagraphs[^1].ParagraphProperties.SectionProperties = sectionProperties;
                    output.AddRange(groupParagraphs);
                    lastPlanSection = sectionProperties;
                    lastGroupFirstParagraph = groupParagraphs[0];
                }
                else
                {
                    var (width, height) = PdfPlacement.GetAvailableArea(paragraph, documentSection);
                    var breakBeforeFirst = groupIndex == 0 && before != null && m_configuration.PageBreakBefore;
                    var groupParagraphs = FlowLayout.CreateParagraphs(group.Pages, m_context, paragraph.ParagraphProperties, width, height, true, breakBeforeFirst);
                    if (!isLastGroup)
                    {
                        // the next group is a section of its own - end these pages with the page setup of the document
                        groupParagraphs[^1].ParagraphProperties.SectionProperties = (SectionProperties)(sectionEmitted ? continuedCopy : originalCopy).CloneNode(true);
                    }
                    output.AddRange(groupParagraphs);
                    lastPlanSection = null;
                }
                sectionEmitted = true;
            }

            // (3) the end of the pages
            var lastParagraph = output[^1];
            var lastIsSection = lastPlanSection != null;
            if (followedByPlan)
            {
                // the next section holds plan pages laid out before (loop items are rendered last to first)
                if (IsNothingBetween(paragraph, plan.FirstParagraph))
                {
                    if (!lastIsSection)
                    {
                        lastParagraph.ParagraphProperties.SectionProperties = (SectionProperties)continuedCopy.CloneNode(true);
                    }
                }
                else
                {
                    CloseContentBeforePlan(plan.FirstParagraph, continuedCopy);
                }
            }
            else if (holder == paragraph)
            {
                // the placeholder paragraph ended the section and is replaced: the pages end the section now
                governing.Remove();
                if (!lastIsSection)
                {
                    OpenXmlHelper.RemoveSectionStartSettings(governing);
                    lastParagraph.ParagraphProperties.SectionProperties = governing;
                }
            }
            else if (after == null && lastIsSection && IsNothingAfter(paragraph, holder))
            {
                // the pages are the last thing in the section - their section properties take the place of the
                // original ones, otherwise an empty page with the original setup would follow
                lastParagraph.ParagraphProperties.SectionProperties = null;
                lastPlanSection.RemoveAnnotations<PlanSectionAnnotation>();
                lastPlanSection.AddAnnotation(new PlanSectionAnnotation((SectionProperties)continuedCopy.CloneNode(true), lastGroupFirstParagraph));
                if (holder != null)
                {
                    holder.ParagraphProperties.SectionProperties = lastPlanSection;
                }
                else
                {
                    governing.InsertBeforeSelf(lastPlanSection);
                    governing.Remove();
                }
            }
            else
            {
                // the original section properties govern the content after the pages - it continues the section
                OpenXmlHelper.RemoveSectionStartSettings(governing);
            }

            foreach (var element in output)
            {
                paragraph.InsertBeforeSelf(element);
            }
            paragraph.Remove();
        }

        /// <summary>
        /// All pages have the size of the document pages (<c>same:inline</c>): laid out like <c>fit:page</c>.
        /// </summary>
        private void InsertWithoutSections(Paragraph paragraph, List<PlacedPage> pages, SectionProperties documentSection, bool hasContentBefore)
        {
            var (width, height) = PdfPlacement.GetAvailableArea(paragraph, documentSection);
            var generated = FlowLayout.CreateParagraphs(pages, m_context, paragraph.ParagraphProperties, width, height, true,
                hasContentBefore && m_configuration.PageBreakBefore);
            var sectionProperties = paragraph.ParagraphProperties?.SectionProperties;
            if (sectionProperties != null)
            {
                sectionProperties.Remove();
                generated[^1].ParagraphProperties.SectionProperties = sectionProperties;
            }
            foreach (var element in generated)
            {
                paragraph.InsertBeforeSelf(element);
            }
            paragraph.Remove();
        }

        /// <summary>
        /// Ends the section before the pages: the split-off text before the placeholder, the previous paragraph or a
        /// new minimal paragraph gets a copy of the governing section properties. Nothing is needed if the
        /// placeholder starts the document or a section. Returns whether a section was ended.
        /// </summary>
        private static bool CloseContentBefore(Paragraph before, Paragraph paragraph, SectionProperties originalCopy)
        {
            if (before != null && PdfPlacement.HasMarkers(before))
            {
                // e.g. the start of an inline loop: the engine removes the marker later, the section break with it
                var closingAfterMarkers = PdfDrawings.CreateMinimalParagraph();
                closingAfterMarkers.ParagraphProperties.SectionProperties = (SectionProperties)originalCopy.CloneNode(true);
                paragraph.InsertBeforeSelf(closingAfterMarkers);
                return true;
            }
            if (before != null)
            {
                before.ParagraphProperties ??= new ParagraphProperties();
                before.ParagraphProperties.SectionProperties = (SectionProperties)originalCopy.CloneNode(true);
                return true;
            }

            var previous = paragraph.PreviousSibling();
            while (previous != null && !EndsSection(previous) && PdfPlacement.IsIgnorable(previous))
            {
                previous = previous.PreviousSibling();
            }
            if (previous == null || EndsSection(previous))
            {
                return false;
            }
            if (previous == paragraph.PreviousSibling() && previous is Paragraph previousParagraph && !PdfPlacement.HasMarkers(previousParagraph))
            {
                previousParagraph.ParagraphProperties ??= new ParagraphProperties();
                previousParagraph.ParagraphProperties.SectionProperties = (SectionProperties)originalCopy.CloneNode(true);
                return true;
            }
            var closing = PdfDrawings.CreateMinimalParagraph();
            closing.ParagraphProperties.SectionProperties = (SectionProperties)originalCopy.CloneNode(true);
            paragraph.InsertBeforeSelf(closing);
            return true;
        }

        /// <summary>
        /// Content between the placeholder and plan pages laid out before must not end up in their section.
        /// </summary>
        private static void CloseContentBeforePlan(Paragraph firstPlanParagraph, SectionProperties continuedCopy)
        {
            if (firstPlanParagraph.PreviousSibling() is Paragraph previous && previous.ParagraphProperties?.SectionProperties == null && !PdfPlacement.HasMarkers(previous))
            {
                previous.ParagraphProperties ??= new ParagraphProperties();
                previous.ParagraphProperties.SectionProperties = (SectionProperties)continuedCopy.CloneNode(true);
                return;
            }
            var closing = PdfDrawings.CreateMinimalParagraph();
            closing.ParagraphProperties.SectionProperties = (SectionProperties)continuedCopy.CloneNode(true);
            firstPlanParagraph.InsertBeforeSelf(closing);
        }

        private static bool EndsSection(OpenXmlElement element)
        {
            return element is Paragraph { ParagraphProperties.SectionProperties: not null };
        }

        /// <summary>
        /// True if nothing with content follows the placeholder paragraph in its section. <paramref name="holder"/>
        /// is the paragraph that ends the section (null: the last section of the document).
        /// </summary>
        private static bool IsNothingAfter(Paragraph paragraph, Paragraph holder)
        {
            for (var element = paragraph.NextSibling(); element != null; element = element.NextSibling())
            {
                if (element is SectionProperties)
                {
                    return true;
                }
                if (!PdfPlacement.IsIgnorable(element))
                {
                    return false;
                }
                if (element == holder)
                {
                    return true;
                }
            }
            return true;
        }

        private static bool IsNothingBetween(Paragraph paragraph, Paragraph end)
        {
            for (var element = paragraph.NextSibling(); element != null && element != end; element = element.NextSibling())
            {
                if (!PdfPlacement.IsIgnorable(element))
                {
                    return false;
                }
            }
            return true;
        }

        private List<PageGroup> CreateGroups(IReadOnlyList<PlacedPage> pages, SectionProperties documentSection)
        {
            var documentPageSize = documentSection.GetFirstChild<PageSize>();
            long documentWidth = documentPageSize?.Width?.Value ?? 12240;
            long documentHeight = documentPageSize?.Height?.Value ?? 15840;
            var maxSize = m_configuration.MaxPageSizeTwips > 0 ? m_configuration.MaxPageSizeTwips : PdfFormatterConfiguration.WordMaxPageSizeTwips;

            var groups = new List<PageGroup>();
            foreach (var page in pages)
            {
                var exactWidth = page.WidthPoints * 20;
                var exactHeight = page.HeightPoints * 20;
                var width = (long)Math.Round(exactWidth);
                var height = (long)Math.Round(exactHeight);
                var isSection = true;
                var scaled = false;
                if (m_configuration.SamePageSize == PdfSamePageSize.Inline
                    && Math.Abs(width - documentWidth) <= SizeToleranceTwips
                    && Math.Abs(height - documentHeight) <= SizeToleranceTwips)
                {
                    isSection = false;
                }
                else if (width > maxSize || height > maxSize)
                {
                    if (m_configuration.Oversize == PdfOversizePolicy.Error)
                    {
                        throw new OpenXmlTemplateException(
                            $"PDF page {page.PageIndex + 1} of '{m_context.Label}' is {Math.Round(PageTexts.PointsToMillimeters(page.WidthPoints))} × " +
                            $"{Math.Round(PageTexts.PointsToMillimeters(page.HeightPoints))} mm - Word supports pages of at most " +
                            $"{Math.Round(maxSize / 20.0 * 25.4 / 72, 1).ToString(System.Globalization.CultureInfo.InvariantCulture)} mm");
                    }
                    // uniform scale, so the placeholder page keeps the aspect ratio and orientation of the original
                    var factor = Math.Min(maxSize / exactWidth, maxSize / exactHeight);
                    width = Math.Min(maxSize, (long)Math.Floor(exactWidth * factor));
                    height = Math.Min(maxSize, (long)Math.Floor(exactHeight * factor));
                    scaled = true;
                }

                var last = groups.Count > 0 ? groups[^1] : null;
                if (last != null && last.IsSection == isSection && (!isSection || (last.WidthTwips == width && last.HeightTwips == height)))
                {
                    last.Pages.Add(page);
                    last.Scaled.Add(scaled);
                }
                else
                {
                    groups.Add(new PageGroup(isSection, width, height, [page], [scaled]));
                }
            }
            return groups;
        }

        private Paragraph CreatePlanPage(PlacedPage page, PageGroup group, bool pageBreakBefore)
        {
            var margin = Math.Max(0, m_configuration.SectionMarginTwips);
            var scaled = group.Scaled[group.Pages.IndexOf(page)];
            long cx;
            long cy;
            double scale;
            if (margin == 0 && !scaled)
            {
                // true size: the extent comes from the PDF points, the pixel size only determines the resolution
                cx = OpenXmlHelper.PointsToEmu(page.WidthPoints);
                cy = OpenXmlHelper.PointsToEmu(page.HeightPoints);
                scale = 1;
            }
            else
            {
                var availableWidth = OpenXmlHelper.TwipsToEmu(Math.Max(1, group.WidthTwips - (2 * margin)));
                var availableHeight = OpenXmlHelper.TwipsToEmu(Math.Max(1, group.HeightTwips - (2 * margin)));
                double width = OpenXmlHelper.PointsToEmu(page.WidthPoints);
                double height = OpenXmlHelper.PointsToEmu(page.HeightPoints);
                scale = PdfPlacement.FitFactor(width, height, availableWidth, availableHeight);
                cx = (long)Math.Round(width * scale);
                cy = (long)Math.Round(height * scale);
            }
            var offset = OpenXmlHelper.TwipsToEmu(margin);
            // centered in the area inside the margins (a scaled page may not fill it in both directions)
            var offsetX = offset + Math.Max(0, (OpenXmlHelper.TwipsToEmu(group.WidthTwips - (2 * margin)) - cx) / 2);
            var offsetY = offset + Math.Max(0, (OpenXmlHelper.TwipsToEmu(group.HeightTwips - (2 * margin)) - cy) / 2);
            if (margin == 0 && !scaled)
            {
                offsetX = 0;
                offsetY = 0;
            }

            var description = m_context.Describe(page, scale);
            var showOverlay = m_configuration.Overlay switch
            {
                PdfOverlayMode.Always => true,
                PdfOverlayMode.Never => false,
                _ => scaled,
            };
            var altText = PageTexts.AltText(m_configuration, description, m_context.Culture);
            var drawings = new List<Drawing>();
            List<string> overlayLines = null;
            if (showOverlay)
            {
                overlayLines = PageTexts.OverlayLines(m_configuration, description, m_context.Culture);
                if (overlayLines.Count > 0)
                {
                    altText = string.Join(Environment.NewLine, new[] { altText }.Concat(overlayLines).Where(x => !string.IsNullOrEmpty(x)));
                }
            }
            drawings.Add(PdfDrawings.CreatePageAnchor(page.RelationshipId, m_context.NextId(), offsetX, offsetY, cx, cy, altText, m_configuration.Border));
            if (overlayLines is { Count: > 0 })
            {
                drawings.Add(PdfDrawings.CreateOverlay(m_context.NextId(), overlayLines, OpenXmlHelper.TwipsToEmu(group.WidthTwips), string.Join(Environment.NewLine, overlayLines)));
            }
            return PdfDrawings.CreateAnchorParagraph(pageBreakBefore, drawings);
        }

        /// <summary>
        /// Section properties of a group of pages: exact page size, margins (default 0), empty or inherited header
        /// and footer. <paramref name="startSection"/> is set for the first section that replaces the start of the
        /// original section: a restart of the page numbering moves to it.
        /// </summary>
        private SectionProperties CreatePlanSectionProperties(PageGroup group, SectionProperties documentSection, SectionProperties startSection)
        {
            var margin = Math.Max(0, m_configuration.SectionMarginTwips);
            var sectionProperties = new SectionProperties();
            var documentMargin = documentSection.GetFirstChild<PageMargin>();
            uint headerDistance = 0;
            uint footerDistance = 0;
            if (m_configuration.HeaderFooter == PdfHeaderFooter.None)
            {
                var (headerId, footerId) = m_context.EmptyHeaderFooter();
                var evenAndOdd = m_context.MainDocumentPart.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>() is { } setting
                                 && (setting.Val == null || setting.Val.Value);
                sectionProperties.AppendChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId });
                if (evenAndOdd)
                {
                    sectionProperties.AppendChild(new HeaderReference { Type = HeaderFooterValues.Even, Id = headerId });
                }
                sectionProperties.AppendChild(new FooterReference { Type = HeaderFooterValues.Default, Id = footerId });
                if (evenAndOdd)
                {
                    sectionProperties.AppendChild(new FooterReference { Type = HeaderFooterValues.Even, Id = footerId });
                }
            }
            else
            {
                foreach (var reference in documentSection.ChildElements.Where(x => x is HeaderReference or FooterReference))
                {
                    sectionProperties.AppendChild(reference.CloneNode(true));
                }
                headerDistance = documentMargin?.Header?.Value ?? 720;
                footerDistance = documentMargin?.Footer?.Value ?? 720;
            }

            sectionProperties.AppendChild(new SectionType { Val = SectionMarkValues.NextPage });
            var pageSize = new PageSize { Width = (uint)group.WidthTwips, Height = (uint)group.HeightTwips };
            if (group.WidthTwips > group.HeightTwips)
            {
                pageSize.Orient = PageOrientationValues.Landscape;
            }
            sectionProperties.AppendChild(pageSize);
            sectionProperties.AppendChild(new PageMargin
            {
                Top = margin,
                Right = (uint)margin,
                Bottom = margin,
                Left = (uint)margin,
                Header = headerDistance,
                Footer = footerDistance,
                Gutter = 0U,
            });
            var start = startSection?.GetFirstChild<PageNumberType>()?.Start;
            if (start != null)
            {
                sectionProperties.AppendChild(new PageNumberType { Start = start.Value });
            }
            return sectionProperties;
        }

        private sealed record PageGroup(bool IsSection, long WidthTwips, long HeightTwips, List<PlacedPage> Pages, List<bool> Scaled);
    }
}

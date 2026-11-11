using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// Inserts PDF pages as inline pictures into the text flow (<c>fit:page</c>, <c>fit:width</c>, <c>fit:none</c>,
    /// <c>w</c>/<c>h</c>): one paragraph per page, a page break before every page after the first.
    /// </summary>
    internal static class FlowLayout
    {
        public static void Insert(Text target, IReadOnlyList<PlacedPage> pages, PdfInsertContext context)
        {
            var section = PdfPlacement.GetGoverningSection(target, context.MainDocumentPart);
            var (availableWidth, availableHeight) = PdfPlacement.GetAvailableArea(target, section);
            var part = TargetPartAnnotation.GetTargetPart(target);
            // no page breaks in table cells, headers and footers - the pages just follow each other
            var withPageBreaks = target.GetFirstAncestor<TableCell>() == null && part is not HeaderPart and not FooterPart;

            var (_, paragraph, _) = PdfPlacement.SplitAround(target);
            var generated = CreateParagraphs(pages, context, paragraph.ParagraphProperties, availableWidth, availableHeight,
                withPageBreaks, withPageBreaks && context.Configuration.PageBreakBefore);

            // the placeholder paragraph ends a section - the last page paragraph takes over the section break
            var sectionProperties = paragraph.ParagraphProperties?.SectionProperties;
            if (sectionProperties != null && generated.Count > 0)
            {
                sectionProperties.Remove();
                generated[^1].ParagraphProperties.SectionProperties = sectionProperties;
            }
            foreach (var page in generated)
            {
                paragraph.InsertBeforeSelf(page);
            }
            paragraph.Remove();
        }

        /// <summary>
        /// One paragraph per page, each page scaled into the available area.
        /// </summary>
        public static List<Paragraph> CreateParagraphs(
            IReadOnlyList<PlacedPage> pages,
            PdfInsertContext context,
            ParagraphProperties templateProperties,
            long availableWidthEmu,
            long availableHeightEmu,
            bool withPageBreaks,
            bool breakBeforeFirst)
        {
            var configuration = context.Configuration;
            var result = new List<Paragraph>();
            for (var i = 0; i < pages.Count; i++)
            {
                var page = pages[i];
                var (cx, cy, scale) = PdfPlacement.GetFlowExtent(page, configuration, availableWidthEmu, availableHeightEmu);
                var id = context.NextId();
                var altText = PageTexts.AltText(configuration, context.Describe(page, scale), context.Culture);
                var drawing = PdfDrawings.CreateInline(page.RelationshipId, id, cx, cy, altText, configuration.Border);
                var pageBreak = withPageBreaks && (i > 0 || breakBeforeFirst);
                result.Add(PdfDrawings.CreateFlowParagraph(templateProperties, pageBreak, drawing));
            }
            return result;
        }
    }
}

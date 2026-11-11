using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// A rendered page stored in an image part of the target part.
    /// </summary>
    internal sealed record PlacedPage(int PageIndex, double WidthPoints, double HeightPoints, string RelationshipId);

    /// <summary>
    /// Everything the layouts need to insert the pages of one placeholder.
    /// </summary>
    internal sealed class PdfInsertContext
    {
        private uint m_nextId;

        public PdfInsertContext(PdfFormatterConfiguration configuration, MainDocumentPart mainDocumentPart, string label, int pageCount, CultureInfo culture, uint firstId, Func<(string HeaderId, string FooterId)> emptyHeaderFooter)
        {
            Configuration = configuration;
            MainDocumentPart = mainDocumentPart;
            Label = label;
            PageCount = pageCount;
            Culture = culture ?? CultureInfo.CurrentCulture;
            m_nextId = firstId;
            EmptyHeaderFooter = emptyHeaderFooter;
        }

        public PdfFormatterConfiguration Configuration { get; }
        public MainDocumentPart MainDocumentPart { get; }
        public string Label { get; }
        public int PageCount { get; }
        public CultureInfo Culture { get; }
        public Func<(string HeaderId, string FooterId)> EmptyHeaderFooter { get; }
        public uint LastId => m_nextId - 1;

        public uint NextId()
        {
            return m_nextId++;
        }

        public PageDescription Describe(PlacedPage page, double scale)
        {
            return new PageDescription(Label, page.PageIndex + 1, PageCount, page.WidthPoints, page.HeightPoints, scale);
        }
    }

    internal static class PdfPlacement
    {
        private const long DefaultCellMarginTwips = 108;

        /// <summary>
        /// Splits the placeholder paragraph into the content before the placeholder, the placeholder itself and the
        /// content after it. Before/after are null if there is nothing there. A section break of the paragraph is
        /// kept only on the last of the parts.
        /// </summary>
        public static (Paragraph Before, Paragraph Target, Paragraph After) SplitAround(Text target)
        {
            var paragraph = target.GetFirstAncestor<Paragraph>();
            var hasBefore = paragraph.Descendants().TakeWhile(x => x != target).Any(IsSignificant);

            var afterParts = paragraph.SplitAfterElement(target);
            var after = afterParts.Count > 1 ? afterParts.Last() as Paragraph : null;
            if (after != null && !after.Descendants().Any(IsSignificant))
            {
                // only empty runs, bookmark ends, ... - the section break (if any) is on the target part as well
                after.Remove();
                after = null;
            }

            Paragraph before = null;
            if (hasBefore)
            {
                var parts = paragraph.SplitBeforeElement(target);
                before = parts.Count > 1 ? parts.First() as Paragraph : null;
                paragraph = target.GetFirstAncestor<Paragraph>();
            }

            var ordered = new[] { before, paragraph, after }.Where(x => x != null).ToList();
            foreach (var part in ordered.Take(ordered.Count - 1))
            {
                part.ParagraphProperties?.SectionProperties = null;
            }
            return (before, paragraph, after);
        }

        /// <summary>
        /// Content that makes a paragraph part non-empty: visible text, pictures, breaks, fields, and the markers the
        /// template engine still needs (loop and condition markers, insertion points).
        /// </summary>
        public static bool IsSignificant(OpenXmlElement element)
        {
            return element switch
            {
                Text text => text.IsMarked() || text.Text.Trim().Length > 0,
                Drawing or Picture or EmbeddedObject or Break or TabChar or SymbolChar or FieldChar or FieldCode or SimpleField or PositionalTab => true,
                _ => HasInsertionPoint(element),
            };
        }

        /// <summary>
        /// True if loops or conditions may insert content at the element later. The end markers of blocks
        /// (<c>End_...</c>) only mark where a block ends - nothing is inserted there and the template engine removes
        /// them (with their paragraph, if it holds nothing else).
        /// </summary>
        public static bool HasInsertionPoint(OpenXmlElement element)
        {
            return element.ExtendedAttributes.Any(a => a.LocalName == InsertionPoint.InsertionPointAttributeName
                                                       && !(a.Value ?? string.Empty).StartsWith("End_", StringComparison.Ordinal));
        }

        public static bool ContainsInsertionPoint(OpenXmlElement element)
        {
            return HasInsertionPoint(element) || element.Descendants().Any(HasInsertionPoint);
        }

        /// <summary>
        /// True for body elements that do not produce content of their own in the final document: empty paragraphs
        /// and paragraphs holding only markers that the template engine removes, bookmarks, ... Insertion points of
        /// loops and conditions are not ignorable - content may be inserted there later.
        /// </summary>
        public static bool IsIgnorable(OpenXmlElement element)
        {
            switch (element)
            {
                case BookmarkStart or BookmarkEnd or ProofError or PermStart or PermEnd or CommentRangeStart or CommentRangeEnd:
                    return !HasInsertionPoint(element);
                case Paragraph paragraph:
                    if (ContainsInsertionPoint(paragraph))
                    {
                        return false;
                    }
                    foreach (var descendant in paragraph.Descendants())
                    {
                        switch (descendant)
                        {
                            case Text text:
                                if (text.IsMarked())
                                {
                                    // variables are replaced by content, other markers are removed
                                    var marker = text.GetMarker();
                                    if (marker is PatternType.Variable or PatternType.Expression)
                                    {
                                        return false;
                                    }
                                }
                                else if (text.Text.Length > 0)
                                {
                                    return false;
                                }
                                break;
                            case Drawing or Picture or EmbeddedObject or Break or TabChar or SymbolChar or FieldChar or FieldCode or SimpleField or PositionalTab:
                                return false;
                        }
                    }
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// True if the element holds insertion points or block markers (loops, conditions) - the template engine
        /// removes or expands those later, so the element must not carry a section break. Replaced variables are fine.
        /// </summary>
        public static bool HasMarkers(OpenXmlElement element)
        {
            return ContainsInsertionPoint(element)
                   || element.Descendants<Text>().Any(x => x.IsMarked() && x.GetMarker() is not PatternType.Variable and not PatternType.Expression);
        }

        /// <summary>
        /// The section properties governing the placeholder: of its section in the body, otherwise (headers, footers,
        /// detached content) the last section of the document. Null if the document has none.
        /// </summary>
        public static SectionProperties GetGoverningSection(OpenXmlElement element, MainDocumentPart mainDocumentPart)
        {
            var section = element.GetSectionProperties()
                          ?? mainDocumentPart?.Document?.Body?.Elements<SectionProperties>().LastOrDefault();
            var plan = section?.Annotation<PlanSectionAnnotation>();
            return plan?.Original ?? section;
        }

        /// <summary>
        /// The area available for a page picture in the text flow in EMU: the text area of the section, in a table
        /// cell the width of the cell without its margins.
        /// </summary>
        public static (long WidthEmu, long HeightEmu) GetAvailableArea(OpenXmlElement target, SectionProperties section)
        {
            var (width, height) = section.GetContentArea();
            var cell = target.GetFirstAncestor<TableCell>();
            if (cell != null)
            {
                var cellWidth = GetCellTextWidth(cell);
                if (cellWidth > 0)
                {
                    width = Math.Min(width, cellWidth);
                }
            }
            return (OpenXmlHelper.TwipsToEmu(width), OpenXmlHelper.TwipsToEmu(height));
        }

        /// <summary>
        /// Width for the content of a cell in twips: the cell width (dxa) or the sum of its grid columns, minus the
        /// left and right cell margins. 0 if unknown.
        /// </summary>
        public static long GetCellTextWidth(TableCell cell)
        {
            var cellProperties = cell.TableCellProperties;
            long width = 0;
            var cellWidth = cellProperties?.TableCellWidth;
            if (cellWidth?.Type?.Value == TableWidthUnitValues.Dxa && long.TryParse(cellWidth.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dxa))
            {
                width = dxa;
            }
            var table = cell.GetFirstAncestor<Table>();
            if (width <= 0 && table != null && cell.Parent is TableRow row)
            {
                var gridColumns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
                    .Select(x => long.TryParse(x.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var w) ? w : 0)
                    .ToList() ?? [];
                var index = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
                foreach (var previous in row.Elements<TableCell>().TakeWhile(x => x != cell))
                {
                    index += previous.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                }
                var span = cellProperties?.GridSpan?.Val?.Value ?? 1;
                width = gridColumns.Skip(index).Take(span).Sum();
            }
            if (width <= 0)
            {
                return 0;
            }

            var defaults = table?.GetFirstChild<TableProperties>()?.TableCellMarginDefault;
            var left = ParseMargin(cellProperties?.TableCellMargin?.LeftMargin?.Width?.Value)
                       ?? ParseMargin(cellProperties?.TableCellMargin?.StartMargin?.Width?.Value)
                       ?? (long?)defaults?.TableCellLeftMargin?.Width?.Value
                       ?? ParseMargin(defaults?.StartMargin?.Width?.Value)
                       ?? DefaultCellMarginTwips;
            var right = ParseMargin(cellProperties?.TableCellMargin?.RightMargin?.Width?.Value)
                        ?? ParseMargin(cellProperties?.TableCellMargin?.EndMargin?.Width?.Value)
                        ?? (long?)defaults?.TableCellRightMargin?.Width?.Value
                        ?? ParseMargin(defaults?.EndMargin?.Width?.Value)
                        ?? DefaultCellMarginTwips;
            return Math.Max(0, width - left - right);
        }

        private static long? ParseMargin(string value)
        {
            return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var margin) ? margin : null;
        }

        /// <summary>
        /// Scales a page into the given box: uniform factor, both the width and the height fit.
        /// </summary>
        public static double FitFactor(double widthEmu, double heightEmu, double maxWidthEmu, double maxHeightEmu)
        {
            return Math.Min(maxWidthEmu / widthEmu, maxHeightEmu / heightEmu);
        }

        /// <summary>
        /// Size of a page picture in the text flow (fit page/width/none or explicit w/h). Returns the extent and the scale.
        /// </summary>
        public static (long Cx, long Cy, double Scale) GetFlowExtent(PlacedPage page, PdfFormatterConfiguration configuration, long availableWidthEmu, long availableHeightEmu)
        {
            double width = OpenXmlHelper.PointsToEmu(page.WidthPoints);
            double height = OpenXmlHelper.PointsToEmu(page.HeightPoints);
            var usableHeight = availableHeightEmu * (1 - Math.Clamp(configuration.HeightSafetyMargin, 0, 0.5));
            double scale;
            if (configuration.MaxWidthEmu.HasValue || configuration.MaxHeightEmu.HasValue)
            {
                scale = FitFactor(width, height, configuration.MaxWidthEmu ?? double.MaxValue, configuration.MaxHeightEmu ?? double.MaxValue);
            }
            else
            {
                scale = configuration.Fit switch
                {
                    PdfFit.Width => availableWidthEmu / width,
                    PdfFit.None => Math.Min(1, FitFactor(width, height, availableWidthEmu, usableHeight)),
                    _ => FitFactor(width, height, availableWidthEmu, usableHeight),
                };
            }
            return ((long)Math.Round(width * scale), (long)Math.Round(height * scale), scale);
        }

        public static bool IsInTextBox(OpenXmlElement element)
        {
            return element.GetFirstAncestor<TextBoxContent>() != null
                   || element.GetFirstAncestor<DocumentFormat.OpenXml.Vml.TextBox>() != null
                   || element.GetFirstAncestor<DW.Inline>() != null
                   || element.GetFirstAncestor<DW.Anchor>() != null;
        }
    }

    /// <summary>
    /// Marks section properties created for <c>fit:section</c> pages. <see cref="Original"/> holds the page setup of
    /// the document around them, <see cref="FirstParagraph"/> the first paragraph of the section.
    /// </summary>
    internal sealed class PlanSectionAnnotation
    {
        public PlanSectionAnnotation(SectionProperties original, Paragraph firstParagraph)
        {
            Original = original;
            FirstParagraph = firstParagraph;
        }

        public SectionProperties Original { get; }

        public Paragraph FirstParagraph { get; }
    }
}

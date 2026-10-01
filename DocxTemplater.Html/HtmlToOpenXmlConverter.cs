using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.ImageBase;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater.Html
{
    /// <summary>
    /// Converts an HTML fragment into WordprocessingML paragraphs and tables.
    /// </summary>
    internal sealed partial class HtmlToOpenXmlConverter
    {
        private const int DefaultAvailableWidthTwips = 9360;
        private const int EmuPerTwip = 635;
        private const int EmuPerPixel = 9525;
        private const double DefaultFontSizePt = 11;

        private static readonly double[] HeadingFallbackSizes = [20, 16, 14, 12, 11, 10];

        private static readonly HashSet<string> SkippedElements = new(StringComparer.OrdinalIgnoreCase)
        {
            "script", "style", "head", "title", "meta", "link", "template", "noscript", "iframe", "object", "embed",
            "svg", "math", "input", "button", "select", "textarea", "canvas", "audio", "video", "source", "track", "map"
        };

        private static readonly HashSet<string> BlockElements = new(StringComparer.OrdinalIgnoreCase)
        {
            "p", "div", "section", "article", "header", "footer", "main", "nav", "aside", "figure", "figcaption",
            "address", "center", "dl", "dt", "dd", "details", "summary", "form", "fieldset", "legend", "li", "caption"
        };

        private readonly HtmlFormatterConfiguration m_configuration;
        private readonly MainDocumentPart m_mainDocumentPart;
        private readonly OpenXmlPart m_part;
        private readonly OpenXmlElement m_partRoot;
        private readonly IImageService m_imageService;
        private readonly RunProperties m_templateRunProperties;
        private readonly ParagraphProperties m_templateParagraphProperties;
        private readonly ListNumbering m_listNumbering;
        private readonly int m_availableWidthTwips;
        private readonly Stack<RunFormat> m_runFormats = new();
        private readonly Stack<BlockFormat> m_blockFormats = new();

        private OpenXmlCompositeElement m_container;
        private Paragraph m_paragraph;
        private bool m_paragraphIsPreformatted;
        private bool m_lastCharWasSpace;
        private bool m_pageBreakPending;
        private LinkContext m_link;
        private int m_listLevel = -1;
        private uint m_nextDrawingId;
        private Dictionary<string, string> m_paragraphStyleIds;
        private readonly string m_hyperlinkStyleId;

        public HtmlToOpenXmlConverter(
            HtmlFormatterConfiguration configuration,
            MainDocumentPart mainDocumentPart,
            OpenXmlPart part,
            OpenXmlElement partRoot,
            IImageService imageService,
            RunProperties templateRunProperties,
            ParagraphProperties templateParagraphProperties)
        {
            m_configuration = configuration;
            m_mainDocumentPart = mainDocumentPart;
            m_part = part;
            m_partRoot = partRoot;
            m_imageService = imageService;
            m_templateRunProperties = templateRunProperties;
            m_templateParagraphProperties = templateParagraphProperties;
            m_listNumbering = new ListNumbering(mainDocumentPart, configuration);
            m_availableWidthTwips = GetAvailableWidth(mainDocumentPart);
            m_hyperlinkStyleId = FindHyperlinkStyle(mainDocumentPart);
        }

        public Body Convert(string html)
        {
            var parser = new HtmlParser();
            var document = parser.ParseDocument(html);
            var body = new Body();
            m_container = body;
            m_runFormats.Push(new RunFormat());
            m_blockFormats.Push(new BlockFormat());
            if (document.Body != null)
            {
                WalkChildren(document.Body);
            }
            FinishParagraph();
            return body;
        }

        private void WalkChildren(INode node)
        {
            foreach (var child in node.ChildNodes)
            {
                Walk(child);
            }
        }

        private void Walk(INode node)
        {
            switch (node)
            {
                case IText text:
                    WriteText(text.Data);
                    break;
                case IElement element:
                    HandleElement(element);
                    break;
            }
        }

        private void HandleElement(IElement element)
        {
            var tag = element.LocalName.ToLowerInvariant();
            if (SkippedElements.Contains(tag) || element.HasAttribute("hidden"))
            {
                return;
            }

            var css = CssParser.ParseStyleAttribute(element.GetAttribute("style"));
            if (css.TryGetValue("display", out var display) && display.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var pageBreakAfter = IsPageBreak(css, "page-break-after", "break-after");
            if (IsPageBreak(css, "page-break-before", "break-before"))
            {
                FinishParagraph();
                m_pageBreakPending = true;
            }

            var keepTogether = IsBreakAvoided(css, "page-break-inside", "break-inside");
            var keepWithNext = IsBreakAvoided(css, "page-break-after", "break-after");
            var applyKeepOptions = (keepTogether || keepWithNext) && IsBlockLevel(tag);
            var firstElementIndex = 0;
            if (applyKeepOptions)
            {
                FinishParagraph();
                firstElementIndex = m_container.ChildElements.Count;
            }

            switch (tag)
            {
                case "br":
                    WriteBreak();
                    break;
                case "hr":
                    WriteHorizontalRule();
                    break;
                case "img":
                    WriteImage(element, css);
                    break;
                case "table":
                    WriteTable(element, css);
                    break;
                case "ul":
                case "ol":
                    WriteList(element, css, tag == "ol");
                    break;
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    WriteHeading(element, css, tag[1] - '0');
                    break;
                case "blockquote":
                    WriteQuote(element, css);
                    break;
                case "pre":
                    WriteBlock(element, css, m_blockFormats.Peek() with { Preformatted = true, IsPlain = false },
                        CurrentRunFormat with { Font = m_configuration.MonospaceFont });
                    break;
                case "a":
                    WriteLink(element, css);
                    break;
                default:
                    if (BlockElements.Contains(tag))
                    {
                        var block = m_blockFormats.Peek();
                        if (tag == "center")
                        {
                            block = block with { Alignment = JustificationValues.Center };
                        }
                        else if (tag == "dd")
                        {
                            block = block with { IndentLeft = block.IndentLeft + m_configuration.ListIndentPerLevel };
                        }
                        WriteBlock(element, css, block, ApplyTagFormat(tag, element, CurrentRunFormat));
                    }
                    else
                    {
                        WriteInline(element, css, ApplyTagFormat(tag, element, CurrentRunFormat));
                    }
                    break;
            }

            if (applyKeepOptions)
            {
                FinishParagraph();
                ApplyKeepOptions(firstElementIndex, keepTogether, keepWithNext);
            }

            if (pageBreakAfter)
            {
                FinishParagraph();
                m_pageBreakPending = true;
            }
        }

        private static bool IsBlockLevel(string tag)
        {
            return BlockElements.Contains(tag) || tag is "table" or "ul" or "ol" or "blockquote" or "pre"
                or "h1" or "h2" or "h3" or "h4" or "h5" or "h6";
        }

        /// <summary>
        /// Maps css <c>break-inside: avoid</c> and <c>break-after: avoid</c> to Word's "keep lines together" and
        /// "keep with next" for everything an element produced (the container children from <paramref name="firstElementIndex"/>).
        /// Keeping a block together means: no page break inside a paragraph (keepLines), every paragraph but the last
        /// stays with the following one (keepNext) and table rows do not split (cantSplit).
        /// </summary>
        private void ApplyKeepOptions(int firstElementIndex, bool keepTogether, bool keepWithNext)
        {
            var elements = m_container.ChildElements.Skip(firstElementIndex).ToList();
            if (elements.Count == 0)
            {
                return;
            }

            // the paragraphs that end the block - for a closing table all paragraphs of its last row,
            // chaining them to the next paragraph would keep the table together with what follows
            var lastElement = elements[^1];
            var endParagraphs = lastElement is Table lastTable
                ? lastTable.Elements<TableRow>().LastOrDefault()?.Descendants<Paragraph>().ToHashSet() ?? []
                : new HashSet<Paragraph> { lastElement as Paragraph ?? lastElement.Descendants<Paragraph>().Last() };

            if (keepTogether)
            {
                foreach (var row in elements.SelectMany(x => x.Descendants<TableRow>()))
                {
                    row.TableRowProperties ??= new TableRowProperties();
                    if (row.TableRowProperties.GetFirstChild<CantSplit>() == null)
                    {
                        row.TableRowProperties.AppendChild(new CantSplit());
                    }
                }

                var paragraphs = elements.SelectMany(x => x is Paragraph paragraph ? [paragraph] : x.Descendants<Paragraph>());
                foreach (var paragraph in paragraphs)
                {
                    var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
                    properties.KeepLines = new KeepLines();
                    if (!endParagraphs.Contains(paragraph))
                    {
                        properties.KeepNext = new KeepNext();
                    }
                }
            }

            if (keepWithNext)
            {
                foreach (var paragraph in endParagraphs)
                {
                    (paragraph.ParagraphProperties ??= new ParagraphProperties()).KeepNext = new KeepNext();
                }
            }
        }

        private RunFormat CurrentRunFormat => m_runFormats.Peek();

        private void WriteInline(IElement element, Dictionary<string, string> css, RunFormat format)
        {
            m_runFormats.Push(ApplyInlineCss(css, format, isBlock: false));
            try
            {
                WriteChildrenOf(element);
            }
            finally
            {
                m_runFormats.Pop();
            }
        }

        private void WriteBlock(IElement element, Dictionary<string, string> css, BlockFormat block, RunFormat format)
        {
            FinishParagraph();
            m_blockFormats.Push(ApplyBlockCss(element, css, block));
            m_runFormats.Push(ApplyInlineCss(css, format, isBlock: true));
            try
            {
                WriteChildrenOf(element);
                FinishParagraph();
            }
            finally
            {
                m_runFormats.Pop();
                m_blockFormats.Pop();
            }
        }

        private void WriteChildrenOf(IElement element)
        {
            WalkChildren(element);
        }

        private void WriteHeading(IElement element, Dictionary<string, string> css, int level)
        {
            var styleId = ResolveParagraphStyle($"heading {level}", $"Heading{level}");
            var block = m_blockFormats.Peek() with { StyleId = styleId, IsPlain = false, InheritTemplateRunProperties = false, KeepNext = true };
            var format = CurrentRunFormat;
            if (styleId == null)
            {
                // the template has no heading style - emulate it with direct formatting;
                // the outline level keeps it in the navigation pane and table of contents
                format = format with { Bold = true, FontSizePt = HeadingFallbackSizes[level - 1] };
                block = block with { OutlineLevel = level - 1 };
            }
            WriteBlock(element, css, block, format);
        }

        private void WriteQuote(IElement element, Dictionary<string, string> css)
        {
            var styleId = ResolveParagraphStyle(m_configuration.QuoteStyle, m_configuration.QuoteStyle?.Replace(" ", string.Empty));
            var block = m_blockFormats.Peek() with { IsPlain = false };
            block = styleId != null
                ? block with { StyleId = styleId }
                : block with { IndentLeft = block.IndentLeft + m_configuration.ListIndentPerLevel, QuoteBorder = true };
            WriteBlock(element, css, block, CurrentRunFormat);
        }

        #region Text

        private void WriteText(string data)
        {
            if (string.IsNullOrEmpty(data))
            {
                return;
            }
            if (m_blockFormats.Peek().Preformatted)
            {
                WritePreformattedText(data);
                return;
            }

            if (m_paragraph == null)
            {
                m_lastCharWasSpace = true;
            }

            var sb = new StringBuilder(data.Length);
            foreach (var c in data)
            {
                if (c is ' ' or '\t' or '\n' or '\r' or '\f')
                {
                    if (!m_lastCharWasSpace)
                    {
                        sb.Append(' ');
                        m_lastCharWasSpace = true;
                    }
                }
                else if (IsValidXmlChar(c))
                {
                    sb.Append(c);
                    m_lastCharWasSpace = false;
                }
            }

            if (sb.Length > 0)
            {
                AppendRun(CreateText(sb.ToString()));
            }
        }

        private void WritePreformattedText(string data)
        {
            var lines = data.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    AppendRun(new Break());
                }
                var segments = lines[i].Split('\t');
                for (var s = 0; s < segments.Length; s++)
                {
                    if (s > 0)
                    {
                        AppendRun(new TabChar());
                    }
                    var text = new string(segments[s].Where(IsValidXmlChar).ToArray());
                    if (text.Length > 0)
                    {
                        AppendRun(CreateText(text));
                    }
                }
            }
            m_lastCharWasSpace = false;
        }

        private void WriteBreak()
        {
            AppendRun(new Break());
            m_lastCharWasSpace = true;
        }

        private static Text CreateText(string value)
        {
            var text = new Text(value);
            if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
            {
                text.Space = SpaceProcessingModeValues.Preserve;
            }
            return text;
        }

        private static bool IsValidXmlChar(char c)
        {
            return c is >= (char)0x20 or '\t' or '\n' or '\r';
        }

        #endregion

        #region Paragraph and run handling

        private Paragraph EnsureParagraph()
        {
            if (m_paragraph != null)
            {
                return m_paragraph;
            }

            var block = m_blockFormats.Peek();
            var paragraph = new Paragraph();
            var properties = block.IsPlain && m_templateParagraphProperties != null
                ? (ParagraphProperties)m_templateParagraphProperties.CloneNode(true)
                : new ParagraphProperties();

            if (block.StyleId != null)
            {
                properties.ParagraphStyleId = new ParagraphStyleId { Val = block.StyleId };
            }
            if (block.KeepNext)
            {
                properties.KeepNext = new KeepNext();
            }
            if (block.OutlineLevel.HasValue)
            {
                properties.OutlineLevel = new OutlineLevel { Val = block.OutlineLevel.Value };
            }

            var indentLeft = block.IndentLeft;
            if (block.ListItem is { } listItem)
            {
                if (!listItem.NumberUsed)
                {
                    listItem.NumberUsed = true;
                    properties.NumberingProperties = new NumberingProperties(
                        new NumberingLevelReference { Val = listItem.Level },
                        new NumberingId { Val = listItem.NumberingId });
                    if (properties.ParagraphStyleId == null && ResolveParagraphStyle("List Paragraph", "ListParagraph") is { } listStyle)
                    {
                        properties.ParagraphStyleId = new ParagraphStyleId { Val = listStyle };
                    }
                    indentLeft = 0;
                }
                else
                {
                    // a further paragraph of the same list item - indent it to the text of the item
                    indentLeft += m_configuration.ListIndentPerLevel * (listItem.Level + 1);
                }
            }

            if (indentLeft > 0 || block.FirstLineIndent.HasValue)
            {
                var indentation = new Indentation();
                if (indentLeft > 0)
                {
                    indentation.Left = indentLeft.ToString();
                }
                if (block.FirstLineIndent is > 0)
                {
                    indentation.FirstLine = block.FirstLineIndent.Value.ToString();
                }
                else if (block.FirstLineIndent is < 0)
                {
                    indentation.Hanging = (-block.FirstLineIndent.Value).ToString();
                }
                properties.Indentation = indentation;
            }

            if (block.Alignment.HasValue)
            {
                properties.Justification = new Justification { Val = block.Alignment.Value };
            }

            if (block.ShadingFill != null)
            {
                properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = block.ShadingFill };
            }

            if (block.QuoteBorder)
            {
                properties.ParagraphBorders = new ParagraphBorders(new LeftBorder
                {
                    Val = BorderValues.Single,
                    Size = 18,
                    Space = 8,
                    Color = "A6A6A6"
                });
            }

            if (m_pageBreakPending)
            {
                properties.PageBreakBefore = new PageBreakBefore();
                m_pageBreakPending = false;
            }

            if (properties.HasChildren)
            {
                paragraph.ParagraphProperties = properties;
            }

            m_container.AppendChild(paragraph);
            m_paragraph = paragraph;
            m_paragraphIsPreformatted = block.Preformatted;
            return paragraph;
        }

        /// <summary>
        /// Completes the current paragraph - the next inline content starts a new one.
        /// </summary>
        private void FinishParagraph()
        {
            if (m_paragraph == null)
            {
                return;
            }

            if (!m_paragraphIsPreformatted)
            {
                TrimTrailingWhitespace(m_paragraph);
            }

            // a trailing <br> does not create an empty line in HTML
            var lastRun = m_paragraph.Descendants<Run>().LastOrDefault();
            if (lastRun != null && lastRun.ChildElements.Where(x => x is not RunProperties).ToList() is [Break { Type: null }])
            {
                lastRun.Remove();
            }

            m_paragraph = null;
        }

        private static void TrimTrailingWhitespace(Paragraph paragraph)
        {
            var lastText = paragraph.Descendants<Text>().LastOrDefault();
            if (lastText == null || lastText.Parent?.NextSibling() != null)
            {
                return;
            }
            var trimmed = lastText.Text.TrimEnd(' ');
            if (trimmed.Length == lastText.Text.Length)
            {
                return;
            }
            if (trimmed.Length == 0)
            {
                var run = lastText.Parent;
                lastText.Remove();
                if (run != null && run.ChildElements.All(x => x is RunProperties))
                {
                    run.Remove();
                }
            }
            else
            {
                lastText.Text = trimmed;
            }
        }

        private void AppendRun(OpenXmlElement content)
        {
            var paragraph = EnsureParagraph();
            var run = new Run();
            var runProperties = CreateRunProperties();
            if (runProperties != null)
            {
                run.RunProperties = runProperties;
            }
            run.AppendChild(content);

            if (m_link != null)
            {
                if (m_link.Element == null || m_link.Element.Parent != paragraph || paragraph.LastChild != m_link.Element)
                {
                    m_link.Element = new Hyperlink { History = true };
                    if (m_link.RelationshipId != null)
                    {
                        m_link.Element.Id = m_link.RelationshipId;
                    }
                    if (m_link.Anchor != null)
                    {
                        m_link.Element.Anchor = m_link.Anchor;
                    }
                    paragraph.AppendChild(m_link.Element);
                }
                m_link.Element.AppendChild(run);
            }
            else
            {
                paragraph.AppendChild(run);
            }
        }

        private RunProperties CreateRunProperties()
        {
            var format = CurrentRunFormat;
            var properties = m_blockFormats.Peek().InheritTemplateRunProperties && m_templateRunProperties != null
                ? (RunProperties)m_templateRunProperties.CloneNode(true)
                : new RunProperties();

            if (format.RunStyleId != null)
            {
                properties.RunStyle = new RunStyle { Val = format.RunStyleId };
            }
            if (format.Font != null)
            {
                properties.RunFonts = new RunFonts { Ascii = format.Font, HighAnsi = format.Font, ComplexScript = format.Font, EastAsia = format.Font };
            }
            if (format.Bold.HasValue)
            {
                properties.Bold = format.Bold.Value ? new Bold() : new Bold { Val = false };
                properties.BoldComplexScript = format.Bold.Value ? new BoldComplexScript() : new BoldComplexScript { Val = false };
            }
            if (format.Italic.HasValue)
            {
                properties.Italic = format.Italic.Value ? new Italic() : new Italic { Val = false };
                properties.ItalicComplexScript = format.Italic.Value ? new ItalicComplexScript() : new ItalicComplexScript { Val = false };
            }
            if (format.Caps == true)
            {
                properties.Caps = new Caps();
            }
            if (format.SmallCaps == true)
            {
                properties.SmallCaps = new SmallCaps();
            }
            if (format.Strike.HasValue)
            {
                properties.Strike = format.Strike.Value ? new Strike() : new Strike { Val = false };
            }
            if (format.Color != null)
            {
                properties.Color = new Color { Val = format.Color };
            }
            if (format.FontSizePt.HasValue)
            {
                var halfPoints = Math.Max(2, (int)Math.Round(format.FontSizePt.Value * 2)).ToString();
                properties.FontSize = new FontSize { Val = halfPoints };
                properties.FontSizeComplexScript = new FontSizeComplexScript { Val = halfPoints };
            }
            if (format.Highlight.HasValue)
            {
                properties.Highlight = new Highlight { Val = format.Highlight.Value };
            }
            if (format.Underline.HasValue)
            {
                properties.Underline = new Underline { Val = format.Underline.Value };
            }
            if (format.BackgroundColor != null)
            {
                properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = format.BackgroundColor };
            }
            if (format.VerticalPosition.HasValue)
            {
                properties.VerticalTextAlignment = new VerticalTextAlignment { Val = format.VerticalPosition.Value };
            }

            return properties.HasChildren ? properties : null;
        }

        /// <summary>
        /// Renders content into another container (e.g. a table cell) and restores the current state afterwards.
        /// </summary>
        private void RenderInto(OpenXmlCompositeElement container, BlockFormat block, RunFormat format, Action render)
        {
            var previousContainer = m_container;
            var previousParagraph = m_paragraph;
            var previousPreformatted = m_paragraphIsPreformatted;
            var previousLastCharWasSpace = m_lastCharWasSpace;
            var previousLink = m_link;
            var previousListLevel = m_listLevel;
            m_container = container;
            m_paragraph = null;
            m_link = null;
            m_listLevel = -1;
            m_blockFormats.Push(block);
            m_runFormats.Push(format);
            try
            {
                render();
                FinishParagraph();
            }
            finally
            {
                m_runFormats.Pop();
                m_blockFormats.Pop();
                m_container = previousContainer;
                m_paragraph = previousParagraph;
                m_paragraphIsPreformatted = previousPreformatted;
                m_lastCharWasSpace = previousLastCharWasSpace;
                m_link = previousLink;
                m_listLevel = previousListLevel;
            }
        }

        #endregion

        #region Formatting

        private RunFormat ApplyTagFormat(string tag, IElement element, RunFormat format)
        {
            switch (tag)
            {
                case "b":
                case "strong":
                case "th":
                    return format with { Bold = true };
                case "i":
                case "em":
                case "cite":
                case "dfn":
                case "var":
                case "address":
                    return format with { Italic = true };
                case "u":
                case "ins":
                    return format with { Underline = UnderlineValues.Single };
                case "s":
                case "strike":
                case "del":
                    return format with { Strike = true };
                case "sub":
                    return format with { VerticalPosition = VerticalPositionValues.Subscript };
                case "sup":
                    return format with { VerticalPosition = VerticalPositionValues.Superscript };
                case "code":
                case "kbd":
                case "samp":
                case "tt":
                    return format with { Font = m_configuration.MonospaceFont };
                case "mark":
                    return format with { Highlight = HighlightColorValues.Yellow };
                case "small":
                    return format with { FontSizePt = (format.FontSizePt ?? TemplateFontSize) / 1.2 };
                case "big":
                    return format with { FontSizePt = (format.FontSizePt ?? TemplateFontSize) * 1.2 };
                case "dt":
                    return format with { Bold = true };
                case "font":
                    var color = CssParser.ParseColor(element.GetAttribute("color"));
                    var face = ParseFontFamily(element.GetAttribute("face"));
                    var size = CssParser.ParseLegacyFontSize(element.GetAttribute("size"));
                    return format with
                    {
                        Color = color ?? format.Color,
                        Font = face ?? format.Font,
                        FontSizePt = size ?? format.FontSizePt
                    };
                default:
                    return format;
            }
        }

        private RunFormat ApplyInlineCss(Dictionary<string, string> css, RunFormat format, bool isBlock)
        {
            if (css.Count == 0)
            {
                return format;
            }

            if (css.TryGetValue("color", out var colorValue) && CssParser.ParseColor(colorValue) is { } color)
            {
                format = format with { Color = color };
            }
            if (!isBlock && (css.TryGetValue("background-color", out var bgValue) || css.TryGetValue("background", out bgValue))
                && CssParser.ParseColor(bgValue) is { } background)
            {
                format = format with { BackgroundColor = background };
            }
            if (css.TryGetValue("font-weight", out var weight))
            {
                weight = weight.Trim().ToLowerInvariant();
                if (weight is "bold" or "bolder" || (int.TryParse(weight, out var numericWeight) && numericWeight >= 600))
                {
                    format = format with { Bold = true };
                }
                else if (weight is "normal" or "lighter" || int.TryParse(weight, out _))
                {
                    format = format with { Bold = false };
                }
            }
            if (css.TryGetValue("font-style", out var fontStyle))
            {
                fontStyle = fontStyle.Trim().ToLowerInvariant();
                if (fontStyle is "italic" or "oblique")
                {
                    format = format with { Italic = true };
                }
                else if (fontStyle == "normal")
                {
                    format = format with { Italic = false };
                }
            }
            if (css.TryGetValue("text-decoration", out var decoration) || css.TryGetValue("text-decoration-line", out decoration))
            {
                decoration = decoration.ToLowerInvariant();
                if (decoration.Contains("underline"))
                {
                    format = format with { Underline = UnderlineValues.Single };
                }
                if (decoration.Contains("line-through"))
                {
                    format = format with { Strike = true };
                }
                if (decoration.Trim() == "none")
                {
                    format = format with { Underline = UnderlineValues.None, Strike = false };
                }
            }
            if (css.TryGetValue("font-size", out var fontSizeValue)
                && CssParser.ParseFontSize(fontSizeValue, format.FontSizePt ?? TemplateFontSize) is { } fontSize)
            {
                format = format with { FontSizePt = fontSize };
            }
            if (css.TryGetValue("font-family", out var family) && ParseFontFamily(family) is { } font)
            {
                format = format with { Font = font };
            }
            if (css.TryGetValue("vertical-align", out var verticalAlign))
            {
                verticalAlign = verticalAlign.Trim().ToLowerInvariant();
                if (verticalAlign == "sub")
                {
                    format = format with { VerticalPosition = VerticalPositionValues.Subscript };
                }
                else if (verticalAlign == "super")
                {
                    format = format with { VerticalPosition = VerticalPositionValues.Superscript };
                }
            }
            if (css.TryGetValue("text-transform", out var transform) && transform.Trim().Equals("uppercase", StringComparison.OrdinalIgnoreCase))
            {
                format = format with { Caps = true };
            }
            if (css.TryGetValue("font-variant", out var variant) && variant.Contains("small-caps", StringComparison.OrdinalIgnoreCase))
            {
                format = format with { SmallCaps = true };
            }
            return format;
        }

        private static BlockFormat ApplyBlockCss(IElement element, Dictionary<string, string> css, BlockFormat block)
        {
            var ownFormatting = false;
            var align = css.TryGetValue("text-align", out var textAlign) ? textAlign : element.GetAttribute("align");
            if (ParseAlignment(align) is { } alignment)
            {
                block = block with { Alignment = alignment };
                ownFormatting = true;
            }

            var indent = 0;
            if (css.TryGetValue("margin-left", out var marginLeft) && CssParser.ParseLengthToTwips(marginLeft) is { } margin)
            {
                indent += margin;
            }
            if (css.TryGetValue("padding-left", out var paddingLeft) && CssParser.ParseLengthToTwips(paddingLeft) is { } padding)
            {
                indent += padding;
            }
            if (indent > 0)
            {
                block = block with { IndentLeft = block.IndentLeft + indent };
                ownFormatting = true;
            }

            if (css.TryGetValue("text-indent", out var textIndent) && CssParser.ParseLengthToTwips(textIndent) is { } firstLine)
            {
                block = block with { FirstLineIndent = firstLine };
                ownFormatting = true;
            }

            if ((css.TryGetValue("background-color", out var bgValue) || css.TryGetValue("background", out bgValue))
                && CssParser.ParseColor(bgValue) is { } background)
            {
                block = block with { ShadingFill = background };
                ownFormatting = true;
            }

            return ownFormatting ? block with { IsPlain = false } : block;
        }

        private static JustificationValues? ParseAlignment(string value)
        {
            return value?.Trim().ToLowerInvariant() switch
            {
                "left" or "start" => JustificationValues.Left,
                "center" or "middle" => JustificationValues.Center,
                "right" or "end" => JustificationValues.Right,
                "justify" => JustificationValues.Both,
                _ => null
            };
        }

        private static string ParseFontFamily(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            var first = value.Split(',')[0].Trim().Trim('"', '\'').Trim();
            return first.ToLowerInvariant() switch
            {
                "" or "inherit" or "initial" => null,
                "serif" => "Times New Roman",
                "sans-serif" => "Arial",
                "monospace" => "Courier New",
                _ => first
            };
        }

        private static bool IsBreakAvoided(Dictionary<string, string> css, string legacyName, string name)
        {
            return (css.TryGetValue(legacyName, out var legacy) && legacy.Trim().Equals("avoid", StringComparison.OrdinalIgnoreCase))
                   || (css.TryGetValue(name, out var value) && value.Trim().ToLowerInvariant() is "avoid" or "avoid-page");
        }

        private static bool IsPageBreak(Dictionary<string, string> css, string legacyName, string name)
        {
            return (css.TryGetValue(legacyName, out var legacy) && legacy.Trim().Equals("always", StringComparison.OrdinalIgnoreCase))
                   || (css.TryGetValue(name, out var value) && value.Trim().Equals("page", StringComparison.OrdinalIgnoreCase));
        }

        private double TemplateFontSize
        {
            get
            {
                var size = m_templateRunProperties?.FontSize?.Val?.Value;
                return size != null && double.TryParse(size, out var halfPoints) ? halfPoints / 2 : DefaultFontSizePt;
            }
        }

        #endregion

        #region Styles

        /// <summary>
        /// Finds a paragraph style by its (english, built in) name or by its id. Style ids are localized
        /// by Word (e.g. "berschrift1" in German templates), the name of built in styles is not.
        /// </summary>
        private string ResolveParagraphStyle(string name, string id)
        {
            m_paragraphStyleIds ??= new Dictionary<string, string>();
            var key = name + "|" + id;
            if (m_paragraphStyleIds.TryGetValue(key, out var cached))
            {
                return cached;
            }

            var styles = m_mainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .Where(x => x.Type?.Value == StyleValues.Paragraph).ToList();
            var style = styles?.FirstOrDefault(x => name != null && string.Equals(x.StyleName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase))
                        ?? styles?.FirstOrDefault(x => id != null && string.Equals(x.StyleId?.Value, id, StringComparison.OrdinalIgnoreCase));
            var result = style?.StyleId?.Value;
            m_paragraphStyleIds[key] = result;
            return result;
        }

        private static string FindHyperlinkStyle(MainDocumentPart mainDocumentPart)
        {
            return mainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .FirstOrDefault(x => x.Type?.Value == StyleValues.Character &&
                                     (string.Equals(x.StyleName?.Val?.Value, "Hyperlink", StringComparison.OrdinalIgnoreCase) || x.StyleId?.Value == "Hyperlink"))
                ?.StyleId?.Value;
        }

        #endregion

        #region Links

        private void WriteLink(IElement element, Dictionary<string, string> css)
        {
            var href = element.GetAttribute("href")?.Trim();
            var link = CreateLink(href);
            if (link == null)
            {
                WriteInline(element, css, CurrentRunFormat);
                return;
            }

            var format = m_hyperlinkStyleId != null
                ? CurrentRunFormat with { RunStyleId = m_hyperlinkStyleId }
                : CurrentRunFormat with { Color = "0563C1", Underline = UnderlineValues.Single };
            var previousLink = m_link;
            m_link = link;
            try
            {
                WriteInline(element, css, format);
            }
            finally
            {
                m_link = previousLink;
            }
        }

        private LinkContext CreateLink(string href)
        {
            if (string.IsNullOrEmpty(href))
            {
                return null;
            }
            if (href.StartsWith('#'))
            {
                return href.Length > 1 ? new LinkContext(null, href[1..]) : null;
            }
            if (m_part == null || !Uri.TryCreate(href, UriKind.RelativeOrAbsolute, out var uri))
            {
                return null;
            }
            if (uri.IsAbsoluteUri && uri.Scheme is not ("http" or "https" or "mailto" or "ftp" or "tel"))
            {
                // e.g. javascript: or file: links are rendered as plain text
                return null;
            }
            var relationship = m_part.AddHyperlinkRelationship(uri, true);
            return new LinkContext(relationship.Id, null);
        }

        #endregion

        #region Lists

        private void WriteList(IElement element, Dictionary<string, string> css, bool ordered)
        {
            FinishParagraph();
            m_listLevel++;
            try
            {
                var level = Math.Min(m_listLevel, 8);
                int? start = null;
                if (ordered && int.TryParse(element.GetAttribute("start"), out var startValue))
                {
                    start = startValue;
                }
                var listType = css.TryGetValue("list-style-type", out var styleType) ? styleType
                    : css.TryGetValue("list-style", out var listStyle) ? listStyle.Split(' ')[0]
                    : element.GetAttribute("type");
                var numberingId = m_listNumbering.GetNumberingId(ordered, level, start, listType);
                var listFormat = ApplyInlineCss(css, CurrentRunFormat, isBlock: true);

                foreach (var child in element.Children)
                {
                    if (child.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase))
                    {
                        WriteListItem(child, numberingId, level, listFormat);
                    }
                    else
                    {
                        // e.g. a nested list directly inside the list (invalid html, but common)
                        m_runFormats.Push(listFormat);
                        try
                        {
                            HandleElement(child);
                        }
                        finally
                        {
                            m_runFormats.Pop();
                        }
                    }
                }
            }
            finally
            {
                m_listLevel--;
            }
        }

        private void WriteListItem(IElement item, int numberingId, int level, RunFormat listFormat)
        {
            FinishParagraph();
            var css = CssParser.ParseStyleAttribute(item.GetAttribute("style"));
            if (css.TryGetValue("display", out var display) && display.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var keepTogether = IsBreakAvoided(css, "page-break-inside", "break-inside");
            var keepWithNext = IsBreakAvoided(css, "page-break-after", "break-after");
            var firstElementIndex = m_container.ChildElements.Count;

            var listItem = new ListItemState(numberingId, level);
            var block = m_blockFormats.Peek() with { ListItem = listItem, IsPlain = false, StyleId = null };
            m_blockFormats.Push(ApplyBlockCss(item, css, block));
            m_runFormats.Push(ApplyInlineCss(css, listFormat, isBlock: true));
            try
            {
                WalkChildren(item);
                if (!listItem.NumberUsed)
                {
                    // an empty list item still shows its bullet
                    EnsureParagraph();
                }
                FinishParagraph();
                if (keepTogether || keepWithNext)
                {
                    ApplyKeepOptions(firstElementIndex, keepTogether, keepWithNext);
                }
            }
            finally
            {
                m_runFormats.Pop();
                m_blockFormats.Pop();
            }
        }

        #endregion

        #region Horizontal rule

        private void WriteHorizontalRule()
        {
            FinishParagraph();
            var paragraph = new Paragraph(new ParagraphProperties(
                new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Size = 6, Space = 1, Color = "auto" })));
            m_container.AppendChild(paragraph);
        }

        #endregion

        #region Images

        private void WriteImage(IElement element, Dictionary<string, string> css)
        {
            var alt = element.GetAttribute("alt");
            var imageBytes = TryReadDataUri(element.GetAttribute("src"));
            if (imageBytes == null || m_imageService == null || m_partRoot == null)
            {
                // external images are never downloaded
                if (!string.IsNullOrWhiteSpace(alt))
                {
                    WriteText(alt);
                }
                return;
            }

            uint maxPropertyId;
            ImageInformation imageInfo;
            try
            {
                maxPropertyId = m_imageService.GetImage(m_partRoot, imageBytes, out imageInfo);
            }
            catch (Exception e) when (e is not OpenXmlTemplateException)
            {
                throw new OpenXmlTemplateException("Invalid image data in HTML", e);
            }

            var naturalWidth = Math.Max(1, imageInfo.PixelWidth);
            var naturalHeight = Math.Max(1, imageInfo.PixelHeight);
            var width = CssParser.ParseLengthToPixels(css.TryGetValue("width", out var cssWidth) ? cssWidth : element.GetAttribute("width"));
            var height = CssParser.ParseLengthToPixels(css.TryGetValue("height", out var cssHeight) ? cssHeight : element.GetAttribute("height"));
            if (width.HasValue && !height.HasValue)
            {
                height = width.Value * naturalHeight / naturalWidth;
            }
            else if (height.HasValue && !width.HasValue)
            {
                width = height.Value * naturalWidth / naturalHeight;
            }
            long cx = (long)Math.Round((width ?? naturalWidth) * EmuPerPixel);
            long cy = (long)Math.Round((height ?? naturalHeight) * EmuPerPixel);

            var maxWidth = (long)m_availableWidthTwips * EmuPerTwip;
            if (cx > maxWidth)
            {
                cy = cy * maxWidth / cx;
                cx = maxWidth;
            }

            // images of this html are not yet part of the document - track the ids ourselves
            var propertyId = Math.Max(maxPropertyId + 1, m_nextDrawingId);
            m_nextDrawingId = propertyId + 1;

            AppendRun(CreateDrawing(imageInfo, propertyId, cx, cy, alt));
            m_lastCharWasSpace = false;
        }

        private static byte[] TryReadDataUri(string src)
        {
            if (string.IsNullOrWhiteSpace(src) || !src.TrimStart().StartsWith("data:image", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            var comma = src.IndexOf(',');
            if (comma < 0)
            {
                return null;
            }
            var header = src[..comma];
            var data = src[(comma + 1)..];
            try
            {
                return header.EndsWith(";base64", StringComparison.OrdinalIgnoreCase)
                    ? System.Convert.FromBase64String(Uri.UnescapeDataString(data).Trim())
                    : Encoding.UTF8.GetBytes(Uri.UnescapeDataString(data));
            }
            catch (FormatException e)
            {
                throw new OpenXmlTemplateException("Invalid base64 image data in HTML", e);
            }
        }

        private Drawing CreateDrawing(ImageInformation imageInfo, uint propertyId, long cx, long cy, string alt)
        {
            var docProperties = new DW.DocProperties { Id = propertyId, Name = $"Picture {propertyId}" };
            if (!string.IsNullOrEmpty(alt))
            {
                docProperties.Description = alt;
            }
            return new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    docProperties,
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(m_imageService.CreatePicture(imageInfo.ImagePartRelationId, propertyId, cx, cy, imageInfo.ExifRotation))
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });
        }

        #endregion

        private static int GetAvailableWidth(MainDocumentPart mainDocumentPart)
        {
            var sectionProperties = mainDocumentPart?.Document?.Body?.Elements<SectionProperties>().LastOrDefault();
            var pageWidth = sectionProperties?.GetFirstChild<PageSize>()?.Width?.Value;
            var margin = sectionProperties?.GetFirstChild<PageMargin>();
            if (pageWidth == null)
            {
                return DefaultAvailableWidthTwips;
            }
            var available = (int)pageWidth.Value - (int)(margin?.Left?.Value ?? 1440) - (int)(margin?.Right?.Value ?? 1440);
            return available > 0 ? available : DefaultAvailableWidthTwips;
        }

        private sealed record RunFormat
        {
            public bool? Bold { get; init; }
            public bool? Italic { get; init; }
            public UnderlineValues? Underline { get; init; }
            public bool? Strike { get; init; }
            public bool? Caps { get; init; }
            public bool? SmallCaps { get; init; }
            public VerticalPositionValues? VerticalPosition { get; init; }
            public string Color { get; init; }
            public string BackgroundColor { get; init; }
            public HighlightColorValues? Highlight { get; init; }
            public string Font { get; init; }
            public double? FontSizePt { get; init; }
            public string RunStyleId { get; init; }
        }

        private sealed record BlockFormat
        {
            public JustificationValues? Alignment { get; init; }
            public int IndentLeft { get; init; }
            public int? FirstLineIndent { get; init; }
            public string StyleId { get; init; }
            public string ShadingFill { get; init; }
            public bool QuoteBorder { get; init; }
            public bool Preformatted { get; init; }
            public bool KeepNext { get; init; }
            public int? OutlineLevel { get; init; }
            public ListItemState ListItem { get; init; }

            /// <summary>
            /// Paragraphs without own paragraph formatting - they take over the paragraph properties of the template.
            /// </summary>
            public bool IsPlain { get; init; } = true;

            public bool InheritTemplateRunProperties { get; init; } = true;
        }

        private sealed class ListItemState
        {
            public ListItemState(int numberingId, int level)
            {
                NumberingId = numberingId;
                Level = level;
            }

            public int NumberingId { get; }
            public int Level { get; }

            /// <summary>
            /// Only the first paragraph of a list item shows the bullet / number.
            /// </summary>
            public bool NumberUsed { get; set; }
        }

        private sealed class LinkContext
        {
            public LinkContext(string relationshipId, string anchor)
            {
                RelationshipId = relationshipId;
                Anchor = anchor;
            }

            public string RelationshipId { get; }
            public string Anchor { get; }
            public Hyperlink Element { get; set; }
        }
    }
}

using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;

namespace DocxTemplater.Html
{
    /// <summary>
    /// Converts HTML to native Word content (paragraphs, runs, lists, tables, images and hyperlinks).
    /// Usage in the template: <c>{{ds.Text}:html}</c>. Optional arguments: <c>ts:'TableStyle'</c>,
    /// <c>ls:'UnorderedListStyle'</c>, <c>ols:'OrderedListStyle'</c>.
    /// </summary>
    /// <remarks>
    /// Replaces the built-in <c>html</c> formatter, which embeds the HTML as an altChunk that only Word itself can render.
    /// </remarks>
    public class HtmlFormatter : IFormatter
    {
        private const int MaxNestingDepth = 3;

        private readonly HtmlFormatterConfiguration m_configuration;
        private int m_nestingDepth;

        public HtmlFormatter(HtmlFormatterConfiguration configuration = null)
        {
            m_configuration = configuration ?? HtmlFormatterConfiguration.Default;
        }

        public bool CanHandle(Type type, string prefix)
        {
            return type == typeof(string) && prefix.Equals("HTML", StringComparison.OrdinalIgnoreCase);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext, Text target)
        {
            if (formatterContext.Value is not string html)
            {
                return;
            }

            if (m_nestingDepth > MaxNestingDepth)
            {
                throw new OpenXmlTemplateException("HTML nesting depth exceeded");
            }

            var targetParagraph = target.GetFirstAncestor<Paragraph>()
                                  ?? throw new OpenXmlTemplateException("HTML placeholder is not in a paragraph");
            var configuration = CreateContextConfiguration(formatterContext.Args);

            m_nestingDepth++;
            try
            {
                var root = target.GetRoot();
                var part = (root as OpenXmlPartRootElement)?.OpenXmlPart;
                var templateParagraphProperties = (ParagraphProperties)targetParagraph.ParagraphProperties?.CloneNode(true);
                // a section break must not be copied to the paragraphs created from html
                templateParagraphProperties?.SectionProperties?.Remove();
                if (templateParagraphProperties is { HasChildren: false })
                {
                    templateParagraphProperties = null;
                }

                var converter = new HtmlToOpenXmlConverter(
                    configuration,
                    templateContext.MainDocumentPart,
                    part,
                    part != null ? root : null,
                    templateContext.ImageService,
                    target.GetFirstAncestor<Run>()?.RunProperties,
                    templateParagraphProperties);
                var container = converter.Convert(html);

                if (configuration.ReplacePlaceholdersInHtml)
                {
                    ReplacePlaceholders(container, templateContext);
                }

                InsertIntoTemplate(target, targetParagraph, container, templateParagraphProperties?.OuterXml ?? string.Empty);
            }
            finally
            {
                m_nestingDepth--;
            }
        }

        private HtmlFormatterConfiguration CreateContextConfiguration(string[] args)
        {
            if (args.Length == 0)
            {
                return m_configuration;
            }
            var configuration = m_configuration.Clone();
            var arguments = HelperFunctions.ParseArguments(args);
            if (arguments.TryGetValue("ts", out var tableStyle))
            {
                configuration.TableStyle = tableStyle;
            }
            if (arguments.TryGetValue("ls", out var listStyle))
            {
                configuration.UnorderedListStyle = listStyle;
            }
            if (arguments.TryGetValue("ols", out var orderedListStyle))
            {
                configuration.OrderedListStyle = orderedListStyle;
            }
            return configuration;
        }

        private static void ReplacePlaceholders(Body container, ITemplateProcessingContext templateContext)
        {
            if (!container.InnerText.Contains('{'))
            {
                return;
            }
            try
            {
                var processor = new XmlNodeTemplate(container, (ITemplateProcessingContextAccess)templateContext);
                processor.Process();
            }
            catch (Exception e)
            {
                throw new OpenXmlTemplateException("Variable replacement in HTML failed", e);
            }
        }

        /// <summary>
        /// Replaces the placeholder with the converted content. The first converted paragraph continues the
        /// template paragraph (text before the placeholder), the last one is continued by the template text after the placeholder.
        /// Paragraphs with own formatting (headings, list items, ...) are only merged with the template paragraph
        /// if the placeholder is the only content there - template text is never turned into a heading or list item.
        /// </summary>
        private static void InsertIntoTemplate(Text target, Paragraph targetParagraph, Body container, string templateParagraphPropertiesXml)
        {
            var elements = container.ChildElements.ToList();
            container.RemoveAllChildren();

            var targetRun = target.GetFirstAncestor<Run>();
            OpenXmlElement anchor = (OpenXmlElement)targetRun ?? target;
            // the placeholder text may share its run with other text - look at the whole paragraph
            var hasContentBefore = targetParagraph.Descendants()
                .TakeWhile(x => x != target)
                .Any(x => x is Text { Text.Length: > 0 } or Drawing or Break or TabChar);
            var cell = targetParagraph.GetFirstAncestor<TableCell>();

            var parts = targetParagraph.SplitAfterElement(target);
            var paragraphAfter = parts.Count > 1 ? parts.Last() as Paragraph : null;

            bool IsPlain(Paragraph paragraph)
            {
                return (paragraph.ParagraphProperties?.OuterXml ?? string.Empty) == templateParagraphPropertiesXml;
            }

            OpenXmlElement last = targetParagraph;
            var lastIsPlain = true;
            var index = 0;
            if (elements.Count > 0 && elements[0] is Paragraph first && (!hasContentBefore || IsPlain(first)))
            {
                if (!IsPlain(first))
                {
                    MergeParagraphProperties(targetParagraph, first);
                    lastIsPlain = false;
                }
                var content = first.ChildElements.Where(x => x is not ParagraphProperties).ToList();
                first.RemoveAllChildren();
                foreach (var element in content)
                {
                    anchor = anchor.InsertAfterSelf(element);
                }
                index = 1;
            }

            for (; index < elements.Count; index++)
            {
                last = last.InsertAfterSelf(elements[index]);
                lastIsPlain = last is Paragraph paragraph && IsPlain(paragraph);
            }

            // continue the last converted paragraph with the template content after the placeholder
            if (paragraphAfter != null && last is Paragraph lastParagraph && lastIsPlain)
            {
                var content = paragraphAfter.ChildElements.Where(x => x is not ParagraphProperties).ToList();
                foreach (var element in content)
                {
                    element.Remove();
                    lastParagraph.AppendChild(element);
                }
                paragraphAfter.Remove();
            }

            target.RemoveWithEmptyParent();

            if (cell != null && cell.LastChild is not Paragraph)
            {
                cell.AppendChild(new Paragraph());
            }
        }

        private static void MergeParagraphProperties(Paragraph target, Paragraph source)
        {
            var sourceProperties = source.ParagraphProperties;
            if (sourceProperties == null)
            {
                return;
            }
            var targetProperties = target.ParagraphProperties ??= new ParagraphProperties();
            if (sourceProperties.ParagraphStyleId != null)
            {
                targetProperties.ParagraphStyleId = (ParagraphStyleId)sourceProperties.ParagraphStyleId.CloneNode(true);
            }
            if (sourceProperties.KeepNext != null)
            {
                targetProperties.KeepNext = (KeepNext)sourceProperties.KeepNext.CloneNode(true);
            }
            if (sourceProperties.KeepLines != null)
            {
                targetProperties.KeepLines = (KeepLines)sourceProperties.KeepLines.CloneNode(true);
            }
            if (sourceProperties.OutlineLevel != null)
            {
                targetProperties.OutlineLevel = (OutlineLevel)sourceProperties.OutlineLevel.CloneNode(true);
            }
            if (sourceProperties.PageBreakBefore != null)
            {
                targetProperties.PageBreakBefore = (PageBreakBefore)sourceProperties.PageBreakBefore.CloneNode(true);
            }
            if (sourceProperties.NumberingProperties != null)
            {
                targetProperties.NumberingProperties = (NumberingProperties)sourceProperties.NumberingProperties.CloneNode(true);
            }
            if (sourceProperties.ParagraphBorders != null)
            {
                targetProperties.ParagraphBorders = (ParagraphBorders)sourceProperties.ParagraphBorders.CloneNode(true);
            }
            if (sourceProperties.Shading != null)
            {
                targetProperties.Shading = (Shading)sourceProperties.Shading.CloneNode(true);
            }
            if (sourceProperties.Indentation != null)
            {
                targetProperties.Indentation = (Indentation)sourceProperties.Indentation.CloneNode(true);
            }
            if (sourceProperties.Justification != null)
            {
                targetProperties.Justification = (Justification)sourceProperties.Justification.CloneNode(true);
            }
        }
    }
}

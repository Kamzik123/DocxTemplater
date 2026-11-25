using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater.Formatter
{
    /// <summary>
    /// Marks the section break of a section of its own that ends a document inserted with the sub-template formatter.
    /// Only kept in memory during <see cref="DocxTemplate.Process"/>.
    /// </summary>
    internal sealed class EndsInsertedDocumentAnnotation
    {
    }

    /// <summary>
    /// Merges the section breaks of a document inserted with the sub-template formatter into the section of the host
    /// document that holds the placeholder.
    /// The page setup the inserted document has of its own is replaced by the host's - as for content without section
    /// breaks, which simply takes the host section. Only sections with a different setup survive the insertion (pages
    /// of their own like <c>fit:section</c> plans, a landscape section made on purpose); the host section is split
    /// around them.
    /// </summary>
    internal sealed class SubTemplateSections
    {
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private readonly SectionProperties m_hostSection;
        private readonly bool m_placeholderEndsSection;
        private readonly List<bool> m_ownSetup;
        // per section break of the inserted content in document order: where a copy of the host section goes,
        // null for a section of its own that is kept
        private readonly List<ParagraphProperties> m_replaced = new();

        private SubTemplateSections(SectionProperties hostSection, bool placeholderEndsSection, List<bool> ownSetup)
        {
            m_hostSection = hostSection;
            m_placeholderEndsSection = placeholderEndsSection;
            m_ownSetup = ownSetup;
        }

        /// <summary>
        /// To be called before the placeholder paragraph is split and the elements are inserted, while the header and
        /// footer references of the inserted document still refer to its own parts. Returns null if there is nothing
        /// to do: no section breaks in the inserted document, and the placeholder paragraph does not end a section.
        /// </summary>
        public static SubTemplateSections Prepare(Body insertedBody, Paragraph placeholder)
        {
            var innerSections = GetInnerSections(insertedBody).ToList();
            // section breaks are only possible in paragraphs directly in the body - not in a table cell, a header or
            // detached content
            var inBody = placeholder.Parent is Body;
            var placeholderSection = inBody ? placeholder.ParagraphProperties?.SectionProperties : null;
            if (innerSections.Count == 0 && placeholderSection == null)
            {
                return null;
            }
            var hostSection = inBody ? placeholder.GetSectionProperties() : null;
            if (hostSection == null)
            {
                return new SubTemplateSections(null, false, null);
            }
            var sourceSection = insertedBody.Elements<SectionProperties>().LastOrDefault();
            // the split of the placeholder paragraph keeps the section on its first part, which holds the content
            // before the placeholder - the section has to end after the inserted content
            placeholderSection?.Remove();
            return new SubTemplateSections(hostSection, placeholderSection != null, innerSections.Select(x => SameSetup(x, sourceSection)).ToList());
        }

        /// <summary>
        /// To be called after inserting, before the relationships of the inserted elements are imported: the sections
        /// that are replaced by the host's do not need the headers and footers of the inserted document.
        /// </summary>
        public void BeforeImport(IReadOnlyList<OpenXmlElement> insertedElements)
        {
            var innerSections = insertedElements.SelectMany(GetInnerSections).ToList();
            for (var i = 0; i < innerSections.Count; i++)
            {
                if (m_hostSection == null)
                {
                    innerSections[i].Remove();
                }
                else if (m_ownSetup[i])
                {
                    m_replaced.Add((ParagraphProperties)innerSections[i].Parent);
                    innerSections[i].Remove();
                }
                else
                {
                    // a section of its own: kept, with the headers and footers of the inserted document
                    m_replaced.Add(null);
                }
            }
        }

        /// <summary>
        /// To be called after the relationships are imported: copies of the host section carry relationship ids of the
        /// host document, they must not be mapped as ids of the inserted document.
        /// </summary>
        public void AfterImport(Paragraph firstPart, Paragraph lastPart, IReadOnlyList<OpenXmlElement> insertedElements, Text placeholder)
        {
            if (m_hostSection == null)
            {
                return;
            }

            // the first section break after the start of the host section keeps its start settings (titlePg, restart
            // of the page numbering), all later parts continue it
            var startUsed = false;
            if (m_replaced.Count > 0 && m_replaced[0] == null)
            {
                // the inserted document starts with a section of its own - the host content before it must not
                // take its page setup
                startUsed = true;
                var firstPartHasContent = HasContent(firstPart, placeholder);
                if (firstPartHasContent && !HasMarkers(firstPart))
                {
                    firstPart.ParagraphProperties ??= new ParagraphProperties();
                    firstPart.ParagraphProperties.SectionProperties = (SectionProperties)m_hostSection.CloneNode(true);
                }
                else if (firstPartHasContent || HasHostContentBefore(firstPart))
                {
                    var closing = CreateMinimalParagraph();
                    closing.ParagraphProperties.SectionProperties = (SectionProperties)m_hostSection.CloneNode(true);
                    firstPart.InsertAfterSelf(closing);
                }
            }
            foreach (var paragraphProperties in m_replaced)
            {
                if (paragraphProperties != null)
                {
                    var copy = (SectionProperties)m_hostSection.CloneNode(true);
                    if (startUsed)
                    {
                        OpenXmlHelper.RemoveSectionStartSettings(copy);
                    }
                    paragraphProperties.SectionProperties = copy;
                }
                startUsed = true;
            }

            if (m_placeholderEndsSection)
            {
                // the content after the placeholder, else the last inserted paragraph, else a new one
                var holder = lastPart
                             ?? (insertedElements.Count > 0 && insertedElements[^1] is Paragraph lastInserted && lastInserted.ParagraphProperties?.SectionProperties == null ? lastInserted : null)
                             ?? (insertedElements.Count > 0 ? insertedElements[^1] : firstPart).InsertAfterSelf(CreateMinimalParagraph());
                holder.ParagraphProperties ??= new ParagraphProperties();
                holder.ParagraphProperties.SectionProperties = m_hostSection;
            }
            if (m_replaced.Count > 0)
            {
                // the content after the last inserted section break continues the host section
                OpenXmlHelper.RemoveSectionStartSettings(m_hostSection);
            }

            if (m_replaced.Count > 0 && m_replaced[^1] == null)
            {
                // the inserted document ends with a section of its own (e.g. PDF pages with final:document): if the
                // host has nothing more before its next section break, that host section would only add an empty
                // page. Whether it does is only known once the host is rendered - see
                // OpenXmlHelper.RemoveEmptySectionsAfterInsertedDocuments
                var lastOwn = insertedElements.SelectMany(GetInnerSections).Last();
                if (lastOwn.Parent?.Parent is Paragraph lastOwnParagraph && !HasContentAfter(lastOwnParagraph, insertedElements))
                {
                    lastOwn.AddAnnotation(new EndsInsertedDocumentAnnotation());
                }
            }
        }

        private static bool HasContentAfter(Paragraph paragraph, IReadOnlyList<OpenXmlElement> insertedElements)
        {
            for (var i = 0; i < insertedElements.Count; i++)
            {
                if (insertedElements[i] == paragraph)
                {
                    return insertedElements.Skip(i + 1).Any(OpenXmlHelper.HasVisibleContent);
                }
                if (paragraph.IsChildOf(insertedElements[i]))
                {
                    // nested in a block level content control: not the end of the document as a section of its own
                    return true;
                }
            }
            return true;
        }

        private static IEnumerable<SectionProperties> GetInnerSections(OpenXmlElement root)
        {
            return root.Descendants<ParagraphProperties>().Select(x => x.SectionProperties).Where(x => x != null);
        }

        private static bool SameSetup(SectionProperties a, SectionProperties b)
        {
            if (a == null || b == null)
            {
                return false;
            }
            return Normalize(a) == Normalize(b);
        }

        private static string Normalize(SectionProperties section)
        {
            var clone = (SectionProperties)section.CloneNode(true);
            // the copies continuing a section have no titlePg / page numbering restart
            OpenXmlHelper.RemoveSectionStartSettings(clone);
            clone.RemoveAttribute("rsidR", WordNamespace);
            clone.RemoveAttribute("rsidRPr", WordNamespace);
            clone.RemoveAttribute("rsidDel", WordNamespace);
            clone.RemoveAttribute("rsidSect", WordNamespace);
            var sectionType = clone.GetFirstChild<SectionType>();
            if (sectionType != null && (sectionType.Val == null || sectionType.Val.Value == SectionMarkValues.NextPage))
            {
                // nextPage is the default
                sectionType.Remove();
            }
            return clone.OuterXml;
        }

        private static bool HasContent(Paragraph paragraph, Text placeholder)
        {
            return paragraph.Descendants<Run>().Any(run => run.ChildElements.Any(x => x is not RunProperties && x != placeholder && x is not Text { Text.Length: 0 }));
        }

        /// <summary>
        /// Markers of loops and conditions are removed by the template engine later - with their paragraph if it holds
        /// nothing else, a section break on it would be lost.
        /// </summary>
        private static bool HasMarkers(OpenXmlElement element)
        {
            return element.Descendants().Prepend(element).Any(x =>
                x.ExtendedAttributes.Any(a => a.LocalName == InsertionPoint.InsertionPointAttributeName)
                || (x is Text && x.IsMarked() && x.GetMarker() is not PatternType.Variable and not PatternType.Expression));
        }

        private static bool HasHostContentBefore(Paragraph firstPart)
        {
            var previous = firstPart.PreviousSibling();
            return previous != null && previous is not SectionProperties && (previous as Paragraph)?.ParagraphProperties?.SectionProperties == null;
        }

        /// <summary>
        /// An empty paragraph of 1 pt height that ends a section where no other paragraph is available.
        /// </summary>
        private static Paragraph CreateMinimalParagraph()
        {
            return new Paragraph(new ParagraphProperties
            {
                KeepLines = new KeepLines(),
                SnapToGrid = new SnapToGrid { Val = false },
                SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0", Line = "20", LineRule = LineSpacingRuleValues.Exact },
                Indentation = new Indentation { Left = "0", Right = "0", FirstLine = "0" },
                ParagraphMarkRunProperties = new ParagraphMarkRunProperties(new FontSize { Val = "2" }, new FontSizeComplexScript { Val = "2" }),
            });
        }
    }
}

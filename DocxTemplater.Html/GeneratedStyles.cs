using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html
{
    /// <summary>
    /// Paragraph styles created for run formatting that repeats on every run of a paragraph (font, size, language, ...).
    /// The runs then only carry what differs, instead of the same run properties on every run of a large html table.
    /// </summary>
    /// <remarks>
    /// The style id is derived from the definition (based-on style + run properties). The same formatting gets the
    /// same style in every document, so documents built from html can be merged into each other without id clashes.
    /// </remarks>
    internal sealed class GeneratedStyles
    {
        private const string IdPrefix = "html";

        private readonly MainDocumentPart m_mainDocumentPart;
        private readonly Dictionary<string, string> m_styleIds = new(StringComparer.Ordinal);
        private Styles m_styles;
        private Dictionary<string, Style> m_existingStyles;
        private string m_defaultParagraphStyleId;

        public GeneratedStyles(MainDocumentPart mainDocumentPart)
        {
            m_mainDocumentPart = mainDocumentPart;
        }

        public bool IsAvailable => m_mainDocumentPart != null;

        /// <summary>
        /// The style a paragraph without own style is shown in - generated styles are based on it,
        /// so they keep its paragraph formatting.
        /// </summary>
        public string DefaultParagraphStyleId
        {
            get
            {
                EnsureStyles();
                return m_defaultParagraphStyleId;
            }
        }

        /// <summary>
        /// False for headings (and other styles with an outline level): a table of contents collects paragraphs by these
        /// styles, a paragraph in a style derived from them would be missing there.
        /// </summary>
        public bool CanDeriveFrom(string styleId)
        {
            if (styleId == null)
            {
                return true;
            }
            EnsureStyles();
            if (!m_existingStyles.TryGetValue(styleId, out var style))
            {
                return true;
            }
            var name = style.StyleName?.Val?.Value ?? string.Empty;
            return !name.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
                   && !name.Equals("Title", StringComparison.OrdinalIgnoreCase)
                   && style.StyleParagraphProperties?.OutlineLevel == null;
        }

        /// <summary>
        /// Returns the id of a paragraph style based on <paramref name="basedOn"/> with the given run properties,
        /// creating the style if the document does not have it yet.
        /// </summary>
        public string GetParagraphStyle(string basedOn, IReadOnlyList<OpenXmlElement> runProperties, IReadOnlyList<string> runPropertiesXml)
        {
            var key = (basedOn ?? string.Empty) + "\n" + string.Concat(runPropertiesXml);
            if (m_styleIds.TryGetValue(key, out var cached))
            {
                return cached;
            }

            EnsureStyles();
            var hash = Hash(key).ToString("x8", CultureInfo.InvariantCulture);
            var styleId = IdPrefix + hash;
            for (var suffix = 2; ; suffix++)
            {
                if (!m_existingStyles.TryGetValue(styleId, out var existing))
                {
                    var style = CreateStyle(styleId, basedOn, runProperties);
                    m_styles.AppendChild(style);
                    m_existingStyles[styleId] = style;
                    break;
                }
                if (IsSameDefinition(existing, basedOn, runPropertiesXml))
                {
                    break;
                }
                // the document has another style with this id (edited by hand or a hash collision)
                styleId = IdPrefix + hash + "_" + suffix.ToString(CultureInfo.InvariantCulture);
            }
            m_styleIds[key] = styleId;
            return styleId;
        }

        private static Style CreateStyle(string styleId, string basedOn, IReadOnlyList<OpenXmlElement> runProperties)
        {
            var style = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            };
            style.AppendChild(new StyleName { Val = "html " + styleId[IdPrefix.Length..] });
            if (basedOn != null)
            {
                style.AppendChild(new BasedOn { Val = basedOn });
            }
            style.AppendChild(new UIPriority { Val = 99 });
            style.AppendChild(new SemiHidden());
            var properties = new StyleRunProperties();
            foreach (var element in runProperties)
            {
                properties.AppendChild(element.CloneNode(true));
            }
            style.AppendChild(properties);
            return style;
        }

        private static bool IsSameDefinition(Style style, string basedOn, IReadOnlyList<string> runPropertiesXml)
        {
            if (style.Type?.Value != StyleValues.Paragraph || style.BasedOn?.Val?.Value != basedOn)
            {
                return false;
            }
            var children = style.StyleRunProperties?.ChildElements.Select(x => x.OuterXml).ToList() ?? [];
            return children.SequenceEqual(runPropertiesXml);
        }

        private void EnsureStyles()
        {
            if (m_styles != null)
            {
                return;
            }
            var stylesPart = m_mainDocumentPart.StyleDefinitionsPart ?? m_mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles ??= new Styles();
            m_styles = stylesPart.Styles;
            m_existingStyles = new Dictionary<string, Style>(StringComparer.Ordinal);
            foreach (var style in m_styles.Elements<Style>())
            {
                var id = style.StyleId?.Value;
                if (id != null)
                {
                    m_existingStyles.TryAdd(id, style);
                }
                if (m_defaultParagraphStyleId == null && style.Type?.Value == StyleValues.Paragraph && style.Default?.Value == true)
                {
                    m_defaultParagraphStyleId = id;
                }
            }
        }

        /// <summary>
        /// FNV-1a - stable across processes and runtimes, unlike <see cref="string.GetHashCode()"/>.
        /// </summary>
        private static uint Hash(string value)
        {
            var hash = 2166136261u;
            foreach (var c in value)
            {
                hash = (hash ^ c) * 16777619u;
            }
            return hash;
        }
    }
}

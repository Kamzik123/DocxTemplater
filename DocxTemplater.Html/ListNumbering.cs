using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html
{
    /// <summary>
    /// Creates the Word numbering definitions for HTML lists.
    /// Every ordered list gets its own numbering instance so its numbering restarts (or starts at the value of the <c>start</c> attribute).
    /// Unordered lists share one instance unless they override the bullet type.
    /// </summary>
    internal sealed class ListNumbering
    {
        private const string OrderedAbstractName = "DocxTemplater.Html.Ordered";
        private const string UnorderedAbstractName = "DocxTemplater.Html.Unordered";

        private readonly MainDocumentPart m_mainDocumentPart;
        private readonly HtmlFormatterConfiguration m_configuration;
        private Numbering m_numbering;
        private int? m_orderedAbstractId;
        private int? m_unorderedAbstractId;
        private int? m_sharedUnorderedNumId;

        public ListNumbering(MainDocumentPart mainDocumentPart, HtmlFormatterConfiguration configuration)
        {
            m_mainDocumentPart = mainDocumentPart;
            m_configuration = configuration;
        }

        /// <summary>
        /// Returns the numbering id (w:numId) to use for the paragraphs of a list.
        /// </summary>
        /// <param name="ordered">true for &lt;ol&gt;</param>
        /// <param name="level">zero based nesting level</param>
        /// <param name="start">start value of an ordered list</param>
        /// <param name="listStyleType">value of the css list-style-type or the html type attribute, may be null</param>
        public int GetNumberingId(bool ordered, int level, int? start, string listStyleType)
        {
            EnsureNumberingPart();
            var abstractId = ordered
                ? m_orderedAbstractId ??= GetOrCreateAbstractNumbering(true)
                : m_unorderedAbstractId ??= GetOrCreateAbstractNumbering(false);

            var levelOverride = TryCreateLevel(level, listStyleType);
            if (!ordered && levelOverride == null)
            {
                m_sharedUnorderedNumId ??= FindPlainInstance(abstractId) ?? CreateInstance(abstractId).NumberID!.Value;
                return m_sharedUnorderedNumId.Value;
            }

            var instance = CreateInstance(abstractId);
            var lvlOverride = new LevelOverride { LevelIndex = level };
            if (ordered)
            {
                lvlOverride.Append(new StartOverrideNumberingValue { Val = Math.Max(0, start ?? 1) });
            }
            if (levelOverride != null)
            {
                levelOverride.StartNumberingValue = new StartNumberingValue { Val = Math.Max(0, start ?? 1) };
                lvlOverride.Append(levelOverride);
            }
            instance.Append(lvlOverride);
            m_numbering.Save();
            return instance.NumberID!.Value;
        }

        private void EnsureNumberingPart()
        {
            if (m_numbering != null)
            {
                return;
            }
            var part = m_mainDocumentPart.NumberingDefinitionsPart ?? m_mainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering ??= new Numbering();
            m_numbering = part.Numbering;
        }

        private int GetOrCreateAbstractNumbering(bool ordered)
        {
            // 1. a list style defined in the template
            var styleName = ordered ? m_configuration.OrderedListStyle : m_configuration.UnorderedListStyle;
            var style = m_mainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .FirstOrDefault(x => x.Type?.Value == StyleValues.Numbering && string.Equals(x.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase));
            if (style != null)
            {
                var styled = m_numbering.Elements<AbstractNum>()
                    .Where(x => x.StyleLink?.Val?.Value == style.StyleId?.Value)
                    .MaxBy(x => x.ChildElements.Count);
                if (styled?.AbstractNumberId != null)
                {
                    return styled.AbstractNumberId.Value;
                }
            }

            // 2. a numbering created by a previous html placeholder in this document
            var name = ordered ? OrderedAbstractName : UnorderedAbstractName;
            var existing = m_numbering.Elements<AbstractNum>().FirstOrDefault(x => x.AbstractNumDefinitionName?.Val?.Value == name);
            if (existing?.AbstractNumberId != null)
            {
                return existing.AbstractNumberId.Value;
            }

            // 3. create a new one
            var id = m_numbering.Elements<AbstractNum>().Select(x => x.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
            var abstractNum = new AbstractNum(
                new MultiLevelType { Val = MultiLevelValues.HybridMultilevel },
                new AbstractNumDefinitionName { Val = name })
            {
                AbstractNumberId = id
            };
            var levels = ordered ? m_configuration.OrderedListLevels : m_configuration.UnorderedListLevels;
            for (var i = 0; i < 9; i++)
            {
                var config = levels[i % levels.Count];
                abstractNum.Append(CreateLevel(i, config.NumberingFormat, config.LevelText));
            }

            // schema order: numPicBullet*, abstractNum*, num*, numIdMacAtCleanup?
            var lastAbstract = m_numbering.Elements<AbstractNum>().LastOrDefault();
            OpenXmlElement insertAfter = lastAbstract ?? (OpenXmlElement)m_numbering.Elements<NumberingPictureBullet>().LastOrDefault();
            if (insertAfter != null)
            {
                insertAfter.InsertAfterSelf(abstractNum);
            }
            else
            {
                m_numbering.PrependChild(abstractNum);
            }
            m_numbering.Save();
            return id;
        }

        private int? FindPlainInstance(int abstractId)
        {
            return m_numbering.Elements<NumberingInstance>()
                .FirstOrDefault(x => x.AbstractNumId?.Val?.Value == abstractId && !x.Elements<LevelOverride>().Any())
                ?.NumberID?.Value;
        }

        private NumberingInstance CreateInstance(int abstractId)
        {
            var numberId = m_numbering.Elements<NumberingInstance>().Select(x => x.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
            var instance = new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = numberId };
            var lastInstance = m_numbering.Elements<NumberingInstance>().LastOrDefault();
            if (lastInstance != null)
            {
                lastInstance.InsertAfterSelf(instance);
            }
            else if (m_numbering.GetFirstChild<NumberingIdMacAtCleanup>() is { } cleanup)
            {
                cleanup.InsertBeforeSelf(instance);
            }
            else
            {
                m_numbering.AppendChild(instance);
            }
            m_numbering.Save();
            return instance;
        }

        private Level CreateLevel(int level, NumberFormatValues format, string levelText)
        {
            return new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = format },
                new LevelText { Val = levelText },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(new Indentation
                {
                    Left = (m_configuration.ListIndentPerLevel * (level + 1)).ToString(),
                    Hanging = "360"
                }))
            {
                LevelIndex = level
            };
        }

        /// <summary>
        /// Creates a level definition for an explicit html list type, or null if the list uses the default type.
        /// </summary>
        private Level TryCreateLevel(int level, string listStyleType)
        {
            if (string.IsNullOrWhiteSpace(listStyleType))
            {
                return null;
            }
            var numberText = $"%{level + 1}.";
            (NumberFormatValues Format, string Text)? definition = listStyleType.Trim() switch
            {
                "1" => (NumberFormatValues.Decimal, numberText),
                "a" => (NumberFormatValues.LowerLetter, numberText),
                "A" => (NumberFormatValues.UpperLetter, numberText),
                "i" => (NumberFormatValues.LowerRoman, numberText),
                "I" => (NumberFormatValues.UpperRoman, numberText),
                _ => listStyleType.Trim().ToLowerInvariant() switch
                {
                    "decimal" => (NumberFormatValues.Decimal, numberText),
                    "decimal-leading-zero" => (NumberFormatValues.DecimalZero, numberText),
                    "lower-alpha" or "lower-latin" => (NumberFormatValues.LowerLetter, numberText),
                    "upper-alpha" or "upper-latin" => (NumberFormatValues.UpperLetter, numberText),
                    "lower-roman" => (NumberFormatValues.LowerRoman, numberText),
                    "upper-roman" => (NumberFormatValues.UpperRoman, numberText),
                    "disc" => (NumberFormatValues.Bullet, "•"),
                    "circle" => (NumberFormatValues.Bullet, "◦"),
                    "square" => (NumberFormatValues.Bullet, "▪"),
                    "none" => (NumberFormatValues.None, string.Empty),
                    _ => null
                }
            };
            return definition.HasValue ? CreateLevel(level, definition.Value.Format, definition.Value.Text) : null;
        }
    }
}

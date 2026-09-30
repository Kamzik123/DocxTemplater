using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html
{
    /// <summary>
    /// Definition of a single list level used when the template has no list style for HTML lists.
    /// <paramref name="LevelText"/> uses the Word syntax, e.g. "%1." for the number of the first level.
    /// </summary>
    public sealed record HtmlListLevel(string LevelText, NumberFormatValues NumberingFormat);

    public class HtmlFormatterConfiguration
    {
        public static readonly HtmlFormatterConfiguration Default = new();

        public HtmlFormatterConfiguration()
        {
            OrderedListLevels = new List<HtmlListLevel>
            {
                new("%1.", NumberFormatValues.Decimal),
                new("%2.", NumberFormatValues.LowerLetter),
                new("%3.", NumberFormatValues.LowerRoman),
                new("%4.", NumberFormatValues.Decimal),
                new("%5.", NumberFormatValues.LowerLetter),
                new("%6.", NumberFormatValues.LowerRoman),
                new("%7.", NumberFormatValues.Decimal),
                new("%8.", NumberFormatValues.LowerLetter),
                new("%9.", NumberFormatValues.LowerRoman),
            };
            UnorderedListLevels = new List<HtmlListLevel>
            {
                new("•", NumberFormatValues.Bullet),
                new("◦", NumberFormatValues.Bullet),
                new("▪", NumberFormatValues.Bullet),
                new("•", NumberFormatValues.Bullet),
                new("◦", NumberFormatValues.Bullet),
                new("▪", NumberFormatValues.Bullet),
                new("•", NumberFormatValues.Bullet),
                new("◦", NumberFormatValues.Bullet),
                new("▪", NumberFormatValues.Bullet),
            };
        }

        /// <summary>
        /// Level definitions for ordered lists (<c>&lt;ol&gt;</c>), used if <see cref="OrderedListStyle"/> is not found in the template.
        /// </summary>
        public List<HtmlListLevel> OrderedListLevels { get; private set; }

        /// <summary>
        /// Level definitions for unordered lists (<c>&lt;ul&gt;</c>), used if <see cref="UnorderedListStyle"/> is not found in the template.
        /// </summary>
        public List<HtmlListLevel> UnorderedListLevels { get; private set; }

        /// <summary>
        /// Indentation per list level in twips (1/20 pt).
        /// </summary>
        public int ListIndentPerLevel { get; set; } = 720;

        /// <summary>
        /// Name of a list style in the template applied to <c>&lt;ol&gt;</c> lists.
        /// If not found, a numbering based on <see cref="OrderedListLevels"/> is created.
        /// </summary>
        public string OrderedListStyle { get; set; } = "html_OrderedListStyle";

        /// <summary>
        /// Name of a list style in the template applied to <c>&lt;ul&gt;</c> lists.
        /// If not found, a numbering based on <see cref="UnorderedListLevels"/> is created.
        /// </summary>
        public string UnorderedListStyle { get; set; } = "html_ListStyle";

        /// <summary>
        /// Name of a table style in the template applied to tables.
        /// If not found, "Table Grid" is used; if that does not exist either, the table gets simple single-line borders.
        /// </summary>
        public string TableStyle { get; set; } = "html_TableStyle";

        /// <summary>
        /// Name of the paragraph style applied to <c>&lt;blockquote&gt;</c>. If not found, the quote is indented and gets a left border.
        /// </summary>
        public string QuoteStyle { get; set; } = "Quote";

        /// <summary>
        /// Font used for <c>&lt;pre&gt;</c>, <c>&lt;code&gt;</c>, <c>&lt;kbd&gt;</c> and <c>&lt;samp&gt;</c>.
        /// </summary>
        public string MonospaceFont { get; set; } = "Courier New";

        /// <summary>
        /// If true, template placeholders (e.g. <c>{{ds.Name}}</c>) contained in the HTML are replaced as well.
        /// Disable this when rendering untrusted HTML (e.g. user input), as placeholders can read any bound model value
        /// and evaluate expressions.
        /// </summary>
        public bool ReplacePlaceholdersInHtml { get; set; } = true;

        public HtmlFormatterConfiguration Clone()
        {
            return new HtmlFormatterConfiguration
            {
                OrderedListLevels = [.. OrderedListLevels],
                UnorderedListLevels = [.. UnorderedListLevels],
                ListIndentPerLevel = ListIndentPerLevel,
                OrderedListStyle = OrderedListStyle,
                UnorderedListStyle = UnorderedListStyle,
                TableStyle = TableStyle,
                QuoteStyle = QuoteStyle,
                MonospaceFont = MonospaceFont,
                ReplacePlaceholdersInHtml = ReplacePlaceholdersInHtml,
            };
        }
    }
}

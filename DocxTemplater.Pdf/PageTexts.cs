using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// A PDF page as it is described in alt texts and overlays.
    /// </summary>
    /// <param name="Name">Name of the PDF (label).</param>
    /// <param name="Page">1-based page number in the PDF.</param>
    /// <param name="Pages">Page count of the PDF.</param>
    /// <param name="WidthPoints">Original page width.</param>
    /// <param name="HeightPoints">Original page height.</param>
    /// <param name="Scale">Displayed size / original size.</param>
    internal sealed record PageDescription(string Name, int Page, int Pages, double WidthPoints, double HeightPoints, double Scale);

    /// <summary>
    /// Fills the configurable text templates (alt text, overlay lines) with the tokens of a page.
    /// </summary>
    internal static class PageTexts
    {
        private static readonly Regex TokenRegex = new(@"\{(?<token>[A-Za-z]+)\}", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(500));

        // portrait sizes in mm
        private static readonly (string Name, double Width, double Height)[] Sheets =
        [
            ("A0", 841, 1189), ("A1", 594, 841), ("A2", 420, 594), ("A3", 297, 420), ("A4", 210, 297), ("A5", 148, 210), ("A6", 105, 148),
            ("B0", 1000, 1414), ("B1", 707, 1000), ("B2", 500, 707), ("B3", 353, 500), ("B4", 250, 353), ("B5", 176, 250),
            ("Letter", 215.9, 279.4), ("Legal", 215.9, 355.6), ("Tabloid", 279.4, 431.8),
            ("ANSI C", 431.8, 558.8), ("ANSI D", 558.8, 863.6), ("ANSI E", 863.6, 1117.6),
        ];

        public static double PointsToMillimeters(double points)
        {
            return points * 25.4 / 72.0;
        }

        /// <summary>
        /// Name of a common sheet size (ISO A/B, US sizes) with ±1 mm tolerance, in either orientation; null if unknown.
        /// </summary>
        public static string FindSheetName(double widthPoints, double heightPoints)
        {
            var shortSide = PointsToMillimeters(Math.Min(widthPoints, heightPoints));
            var longSide = PointsToMillimeters(Math.Max(widthPoints, heightPoints));
            foreach (var (name, width, height) in Sheets)
            {
                if (Math.Abs(shortSide - width) <= 1 && Math.Abs(longSide - height) <= 1)
                {
                    return name;
                }
            }
            return null;
        }

        /// <summary>
        /// Replaces the tokens of the template. Unknown tokens are kept as they are, so a typo in a translation
        /// does not break the document generation.
        /// </summary>
        public static string Format(string template, IReadOnlyDictionary<string, string> tokens)
        {
            if (string.IsNullOrEmpty(template))
            {
                return string.Empty;
            }
            return TokenRegex.Replace(template, m => tokens.TryGetValue(m.Groups["token"].Value, out var value) ? value : m.Value);
        }

        public static Dictionary<string, string> CreateTokens(PdfFormatterConfiguration configuration, PageDescription page, CultureInfo culture)
        {
            culture ??= CultureInfo.CurrentCulture;
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["name"] = page.Name,
                ["page"] = page.Page.ToString(culture),
                ["pages"] = page.Pages.ToString(culture),
                ["sheet"] = FindSheetName(page.WidthPoints, page.HeightPoints) ?? string.Empty,
                ["orientation"] = page.WidthPoints > page.HeightPoints ? configuration.OrientationLandscape : configuration.OrientationPortrait,
                ["width"] = Math.Round(PointsToMillimeters(page.WidthPoints)).ToString("0", culture),
                ["height"] = Math.Round(PointsToMillimeters(page.HeightPoints)).ToString("0", culture),
                ["scale"] = Math.Round(page.Scale, 2).ToString("P0", culture),
            };
        }

        public static string AltText(PdfFormatterConfiguration configuration, PageDescription page, CultureInfo culture)
        {
            return Format(configuration.AltText, CreateTokens(configuration, page, culture));
        }

        /// <summary>
        /// The lines of the placeholder overlay; empty lines are dropped.
        /// </summary>
        public static List<string> OverlayLines(PdfFormatterConfiguration configuration, PageDescription page, CultureInfo culture)
        {
            var tokens = CreateTokens(configuration, page, culture);
            var sizeLine = tokens["sheet"].Length > 0 ? configuration.OverlaySizeLine : configuration.OverlaySizeLineUnknownSheet;
            var lines = new List<string>();
            foreach (var template in new[] { configuration.OverlayTitle, sizeLine, configuration.OverlaySourceLine })
            {
                var line = Format(template, tokens);
                if (!string.IsNullOrWhiteSpace(line))
                {
                    lines.Add(line);
                }
            }
            return lines;
        }
    }
}

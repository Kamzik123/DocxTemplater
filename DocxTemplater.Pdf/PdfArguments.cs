using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// Parses the template arguments of <c>{{x}:pdf(...)}</c> into a configuration.
    /// </summary>
    internal static class PdfArguments
    {
        private static readonly Regex KeyRegex = new(@"^\s*(?<key>[A-Za-z]+)\s*:(?<value>.*)$", RegexOptions.Singleline | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(500));
        private static readonly Regex PercentTokenRegex = new(@"%(?<token>[A-Za-z]+)%", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(500));
        private static readonly Regex LengthRegex = new(@"^\s*(?<value>\d+(?:[.,]\d+)?)\s*(?<unit>cm|mm|in|pt|px)?\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(500));

        /// <summary>
        /// Splits <c>key:value</c> arguments. The core splits the arguments at every comma, also inside quotes
        /// (<c>pages:'1-3,7'</c> arrives as <c>pages:'1-3</c> and <c>7'</c>), so pieces without a key are joined
        /// to the previous argument again. Only the first colon separates key and value (<c>label:'Plan: A'</c>).
        /// </summary>
        public static Dictionary<string, string> Parse(string[] args)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string lastKey = null;
            foreach (var arg in args ?? [])
            {
                if (string.IsNullOrWhiteSpace(arg))
                {
                    continue;
                }
                var match = KeyRegex.Match(arg);
                if (match.Success)
                {
                    lastKey = match.Groups["key"].Value;
                    result[lastKey] = match.Groups["value"].Value;
                }
                else if (lastKey != null)
                {
                    result[lastKey] += "," + arg;
                }
                else
                {
                    throw new OpenXmlTemplateException($"Invalid pdf argument '{arg}' - arguments must be in the form key:value");
                }
            }
            foreach (var key in new List<string>(result.Keys))
            {
                result[key] = Unquote(result[key]);
            }
            return result;
        }

        public static PdfFormatterConfiguration Apply(PdfFormatterConfiguration configuration, string[] args)
        {
            var arguments = Parse(args);
            if (arguments.Count == 0)
            {
                return configuration;
            }
            var result = configuration.Clone();
            foreach (var (key, value) in arguments)
            {
                switch (key.ToUpperInvariant())
                {
                    case "MODE":
                        if (!value.Equals("pages", StringComparison.OrdinalIgnoreCase))
                        {
                            throw new OpenXmlTemplateException($"pdf mode '{value}' is not supported - only mode:pages is implemented");
                        }
                        break;
                    case "PAGES":
                        result.Pages = value;
                        break;
                    case "DPI":
                        result.Dpi = ParseInt(key, value, 72, 600);
                        break;
                    case "FMT":
                        result.ImageFormat = value.ToUpperInvariant() switch
                        {
                            "PNG" => PdfImageFormat.Png,
                            "JPG" or "JPEG" => PdfImageFormat.Jpeg,
                            _ => throw InvalidValue(key, value, "png, jpg"),
                        };
                        break;
                    case "Q":
                        result.JpegQuality = ParseInt(key, value, 1, 100);
                        break;
                    case "FIT":
                        result.Fit = ParseEnum<PdfFit>(key, value);
                        break;
                    case "CONTENT":
                        result.Content = ParseEnum<PdfContent>(key, value);
                        break;
                    case "MARKER":
                        // braces are not allowed in template arguments: tokens are written %page% there
                        result.MarkerText = PercentTokenRegex.Replace(value, "{${token}}");
                        break;
                    case "KEY":
                        result.Key = value;
                        break;
                    case "MARGIN":
                        result.SectionMarginTwips = (int)(ParseLengthEmu(key, value) / 635);
                        break;
                    case "HF":
                        result.HeaderFooter = ParseEnum<PdfHeaderFooter>(key, value);
                        break;
                    case "OVERSIZE":
                        result.Oversize = value.ToUpperInvariant() switch
                        {
                            "SCALE" => PdfOversizePolicy.ScaleWithOverlay,
                            "ERROR" => PdfOversizePolicy.Error,
                            _ => throw InvalidValue(key, value, "scale, error"),
                        };
                        break;
                    case "OVERLAY":
                        result.Overlay = ParseEnum<PdfOverlayMode>(key, value);
                        break;
                    case "SAME":
                        result.SamePageSize = ParseEnum<PdfSamePageSize>(key, value);
                        break;
                    case "W":
                        result.MaxWidthEmu = ParseLengthEmu(key, value);
                        break;
                    case "H":
                        result.MaxHeightEmu = ParseLengthEmu(key, value);
                        break;
                    case "PB":
                        result.PageBreakBefore = ParseBool(key, value);
                        break;
                    case "BORDER":
                        result.Border = ParseBool(key, value);
                        break;
                    case "LABEL":
                        result.Label = value;
                        break;
                    case "PWD":
                        result.Password = value;
                        break;
                    default:
                        throw new OpenXmlTemplateException($"Unknown pdf argument '{key}'");
                }
            }
            return result;
        }

        /// <summary>
        /// Parses a 1-based page selection like <c>1-3,5,8-</c> (also separated by <c>;</c> or spaces) into 0-based
        /// indices in the given order. Null or empty selects all pages.
        /// </summary>
        public static List<int> ParsePageSelection(string selection, int pageCount)
        {
            var result = new List<int>();
            if (string.IsNullOrWhiteSpace(selection))
            {
                for (var i = 0; i < pageCount; i++)
                {
                    result.Add(i);
                }
                return result;
            }
            foreach (var rawPart in selection.Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries))
            {
                var part = rawPart.Trim();
                int from;
                int to;
                var dash = part.IndexOf('-', StringComparison.Ordinal);
                if (dash < 0)
                {
                    from = to = ParsePageNumber(part, selection);
                }
                else
                {
                    var start = part[..dash].Trim();
                    var end = part[(dash + 1)..].Trim();
                    from = start.Length == 0 ? 1 : ParsePageNumber(start, selection);
                    to = end.Length == 0 ? pageCount : ParsePageNumber(end, selection);
                }
                if (from > to)
                {
                    throw new OpenXmlTemplateException($"Invalid page range '{part}' in pdf page selection '{selection}'");
                }
                if (from > pageCount || to > pageCount)
                {
                    throw new OpenXmlTemplateException($"Page {Math.Max(from, to)} of pdf page selection '{selection}' does not exist - the PDF has {pageCount} page(s)");
                }
                for (var page = from; page <= to; page++)
                {
                    result.Add(page - 1);
                }
            }
            return result;
        }

        private static int ParsePageNumber(string value, string selection)
        {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var page) || page < 1)
            {
                throw new OpenXmlTemplateException($"Invalid page '{value}' in pdf page selection '{selection}'");
            }
            return page;
        }

        /// <summary>
        /// Parses a length like <c>16cm</c>, <c>40.5mm</c>, <c>3in</c>, <c>72pt</c> or <c>400px</c> (96 dpi) to EMU.
        /// A number without unit is taken as pixels, like the image formatter does.
        /// </summary>
        public static long ParseLengthEmu(string key, string value)
        {
            var match = LengthRegex.Match(value);
            if (!match.Success)
            {
                throw InvalidValue(key, value, "a length like 16cm, 40mm, 3in, 72pt or 400px");
            }
            var number = double.Parse(match.Groups["value"].Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture);
            var factor = match.Groups["unit"].Value.ToUpperInvariant() switch
            {
                "CM" => 360000.0,
                "MM" => 36000.0,
                "IN" => 914400.0,
                "PT" => 12700.0,
                _ => 9525.0,
            };
            return (long)Math.Round(number * factor);
        }

        private static string Unquote(string value)
        {
            var trimmed = value.Trim();
            if (trimmed.Length >= 2 && IsQuote(trimmed[0]) && IsQuote(trimmed[^1]))
            {
                return trimmed[1..^1];
            }
            return trimmed;
        }

        private static bool IsQuote(char c)
        {
            return c is '\'' or '"' or '‘' or '’' or '“' or '”' or '«' or '»';
        }

        private static int ParseInt(string key, string value, int min, int max)
        {
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) || number < min || number > max)
            {
                throw InvalidValue(key, value, $"a number from {min} to {max}");
            }
            return number;
        }

        private static bool ParseBool(string key, string value)
        {
            return value.ToUpperInvariant() switch
            {
                "TRUE" or "YES" or "1" => true,
                "FALSE" or "NO" or "0" => false,
                _ => throw InvalidValue(key, value, "true, false"),
            };
        }

        private static TEnum ParseEnum<TEnum>(string key, string value)
            where TEnum : struct, Enum
        {
            if (Enum.TryParse<TEnum>(value, true, out var result) && Enum.IsDefined(result) && !int.TryParse(value, out _))
            {
                return result;
            }
            throw InvalidValue(key, value, string.Join(", ", Enum.GetNames<TEnum>()).ToLowerInvariant());
        }

        private static OpenXmlTemplateException InvalidValue(string key, string value, string expected)
        {
            return new OpenXmlTemplateException($"Invalid value '{value}' for pdf argument '{key}' - expected {expected}");
        }
    }
}

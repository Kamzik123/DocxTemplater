using System;
using System.Collections.Generic;
using System.Globalization;

namespace DocxTemplater.Html
{
    /// <summary>
    /// Minimal parser for inline <c>style</c> attributes, CSS colors and CSS lengths.
    /// </summary>
    internal static class CssParser
    {
        private const double DefaultFontSizePt = 12;

        private static readonly char[] RgbSeparators = [',', ' ', '/'];

        private static readonly Dictionary<string, string> NamedColors = new(StringComparer.OrdinalIgnoreCase)
        {
            ["black"] = "000000",
            ["white"] = "FFFFFF",
            ["red"] = "FF0000",
            ["green"] = "008000",
            ["blue"] = "0000FF",
            ["yellow"] = "FFFF00",
            ["orange"] = "FFA500",
            ["purple"] = "800080",
            ["gray"] = "808080",
            ["grey"] = "808080",
            ["silver"] = "C0C0C0",
            ["maroon"] = "800000",
            ["olive"] = "808000",
            ["lime"] = "00FF00",
            ["aqua"] = "00FFFF",
            ["cyan"] = "00FFFF",
            ["teal"] = "008080",
            ["navy"] = "000080",
            ["fuchsia"] = "FF00FF",
            ["magenta"] = "FF00FF",
            ["pink"] = "FFC0CB",
            ["brown"] = "A52A2A",
            ["gold"] = "FFD700",
            ["indigo"] = "4B0082",
            ["violet"] = "EE82EE",
            ["darkred"] = "8B0000",
            ["darkgreen"] = "006400",
            ["darkblue"] = "00008B",
            ["darkgray"] = "A9A9A9",
            ["darkgrey"] = "A9A9A9",
            ["lightgray"] = "D3D3D3",
            ["lightgrey"] = "D3D3D3",
            ["lightblue"] = "ADD8E6",
            ["lightgreen"] = "90EE90",
            ["lightyellow"] = "FFFFE0",
            ["crimson"] = "DC143C",
            ["coral"] = "FF7F50",
            ["salmon"] = "FA8072",
            ["tomato"] = "FF6347",
            ["orangered"] = "FF4500",
            ["darkorange"] = "FF8C00",
            ["khaki"] = "F0E68C",
            ["beige"] = "F5F5DC",
            ["tan"] = "D2B48C",
            ["chocolate"] = "D2691E",
            ["skyblue"] = "87CEEB",
            ["steelblue"] = "4682B4",
            ["royalblue"] = "4169E1",
            ["dodgerblue"] = "1E90FF",
            ["turquoise"] = "40E0D0",
            ["seagreen"] = "2E8B57",
            ["forestgreen"] = "228B22",
            ["limegreen"] = "32CD32",
            ["slategray"] = "708090",
            ["dimgray"] = "696969",
            ["whitesmoke"] = "F5F5F5",
        };

        public static Dictionary<string, string> ParseStyleAttribute(string style)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(style))
            {
                return result;
            }

            foreach (var declaration in style.Split(';'))
            {
                var colon = declaration.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }
                var name = declaration[..colon].Trim();
                var value = declaration[(colon + 1)..].Trim();
                var important = value.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
                if (important >= 0)
                {
                    value = value[..important].Trim();
                }
                if (name.Length > 0 && value.Length > 0)
                {
                    result[name] = value;
                }
            }
            return result;
        }

        /// <summary>
        /// Converts a CSS color to a 6 digit hex string (without '#'), or null if not a supported color.
        /// Transparent colors return null.
        /// </summary>
        public static string ParseColor(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            value = value.Trim();
            if (value.StartsWith('#'))
            {
                var hex = value[1..];
                if (hex.Length is 3 or 4 && IsHex(hex))
                {
                    return string.Concat(hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]).ToUpperInvariant();
                }
                if (hex.Length is 6 or 8 && IsHex(hex))
                {
                    return hex[..6].ToUpperInvariant();
                }
                return null;
            }

            if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
            {
                var open = value.IndexOf('(');
                var close = value.LastIndexOf(')');
                if (open < 0 || close <= open)
                {
                    return null;
                }
                var parts = value[(open + 1)..close].Split(RgbSeparators, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 3)
                {
                    return null;
                }
                if (parts.Length >= 4 && TryParseNumber(parts[3].TrimEnd('%'), out var alpha) && alpha == 0)
                {
                    return null;
                }
                var result = string.Empty;
                for (var i = 0; i < 3; i++)
                {
                    var part = parts[i];
                    double component;
                    if (part.EndsWith('%'))
                    {
                        if (!TryParseNumber(part[..^1], out component))
                        {
                            return null;
                        }
                        component = component * 255 / 100;
                    }
                    else if (!TryParseNumber(part, out component))
                    {
                        return null;
                    }
                    result += ((int)Math.Round(Math.Clamp(component, 0, 255))).ToString("X2", CultureInfo.InvariantCulture);
                }
                return result;
            }

            return NamedColors.TryGetValue(value, out var named) ? named : null;
        }

        /// <summary>
        /// Converts a CSS length to twips (1/20 pt). Percentages are not supported and return null.
        /// Unitless values are interpreted as pixels (as in HTML width/height attributes).
        /// </summary>
        public static int? ParseLengthToTwips(string value)
        {
            var points = ParseLengthToPoints(value, DefaultFontSizePt);
            return points.HasValue ? (int)Math.Round(points.Value * 20) : null;
        }

        public static double? ParseLengthToPixels(string value)
        {
            var points = ParseLengthToPoints(value, DefaultFontSizePt);
            return points.HasValue ? points.Value / 0.75 : null;
        }

        public static double? ParsePercentage(string value)
        {
            if (value == null)
            {
                return null;
            }
            value = value.Trim();
            if (value.EndsWith('%') && TryParseNumber(value[..^1], out var percent))
            {
                return percent;
            }
            return null;
        }

        /// <summary>
        /// Converts a CSS font-size to points. Relative sizes (em, %, smaller, larger) are relative to <paramref name="parentSizePt"/>.
        /// </summary>
        public static double? ParseFontSize(string value, double? parentSizePt)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            var parent = parentSizePt ?? DefaultFontSizePt;
            value = value.Trim().ToLowerInvariant();
            switch (value)
            {
                case "xx-small": return 7;
                case "x-small": return 7.5;
                case "small": return 10;
                case "medium": return 12;
                case "large": return 13.5;
                case "x-large": return 18;
                case "xx-large": return 24;
                case "xxx-large": return 36;
                case "smaller": return parent / 1.2;
                case "larger": return parent * 1.2;
            }
            if (value.EndsWith('%') && TryParseNumber(value[..^1], out var percent))
            {
                return parent * percent / 100;
            }
            return ParseLengthToPoints(value, parent);
        }

        /// <summary>
        /// Maps the legacy <c>&lt;font size="1..7"&gt;</c> attribute to points.
        /// </summary>
        public static double? ParseLegacyFontSize(string value)
        {
            if (!int.TryParse(value?.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var size))
            {
                return null;
            }
            if (value.TrimStart().StartsWith('+') || value.TrimStart().StartsWith('-'))
            {
                size = 3 + size;
            }
            return Math.Clamp(size, 1, 7) switch
            {
                1 => 7.5,
                2 => 10,
                3 => 12,
                4 => 13.5,
                5 => 18,
                6 => 24,
                _ => 36,
            };
        }

        private static double? ParseLengthToPoints(string value, double fontSizePt)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            value = value.Trim().ToLowerInvariant();
            if (value == "0")
            {
                return 0;
            }

            var unitStart = value.Length;
            while (unitStart > 0 && char.IsLetter(value[unitStart - 1]))
            {
                unitStart--;
            }
            var unit = value[unitStart..];
            if (!TryParseNumber(value[..unitStart], out var number))
            {
                return null;
            }
            return unit switch
            {
                "" or "px" => number * 0.75,
                "pt" => number,
                "pc" => number * 12,
                "in" => number * 72,
                "cm" => number * 72 / 2.54,
                "mm" => number * 72 / 25.4,
                "em" or "rem" => number * fontSizePt,
                _ => null,
            };
        }

        private static bool TryParseNumber(string value, out double number)
        {
            return double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number);
        }

        private static bool IsHex(string value)
        {
            foreach (var c in value)
            {
                if (!Uri.IsHexDigit(c))
                {
                    return false;
                }
            }
            return true;
        }
    }
}

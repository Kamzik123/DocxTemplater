using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html
{
    /// <summary>
    /// A resolved border of one side. <see cref="IsNone"/> means "explicitly no border".
    /// </summary>
    /// <param name="Style">Word border style</param>
    /// <param name="Size">width in eighths of a point (Word's unit for borders, 2..96)</param>
    /// <param name="Color">hex color or "auto"</param>
    internal sealed record CssBorder(BorderValues Style, uint Size, string Color)
    {
        public static readonly CssBorder None = new(BorderValues.Nil, 0, "auto");

        public bool IsNone => Style == BorderValues.Nil;

        /// <summary>
        /// A solid border with a width in pixels, as drawn by the html <c>border</c> attribute.
        /// </summary>
        public static CssBorder Solid(double widthPx)
        {
            return widthPx <= 0 ? None : new CssBorder(BorderValues.Single, ToEighthPoints(widthPx), "auto");
        }

        public static uint ToEighthPoints(double widthPx)
        {
            // 1px = 0.75pt = 6 eighths of a point
            return (uint)Math.Clamp(Math.Round(widthPx * 6), 2, 96);
        }
    }

    /// <summary>
    /// The borders of the four sides of a box. A null side was not specified.
    /// </summary>
    internal sealed record CssBoxBorders(CssBorder Top, CssBorder Right, CssBorder Bottom, CssBorder Left)
    {
        public bool IsEmpty => Top == null && Right == null && Bottom == null && Left == null;
    }

    /// <summary>
    /// Parses the css box properties used for tables: <c>border*</c> and <c>padding*</c>.
    /// </summary>
    internal static class CssBoxParser
    {
        private static readonly string[] Sides = ["top", "right", "bottom", "left"];

        public static CssBoxBorders ParseBorders(Dictionary<string, string> css)
        {
            var sides = new BorderSide[4];
            for (var i = 0; i < 4; i++)
            {
                sides[i] = new BorderSide();
            }

            // shorthands first, then the more specific properties (css cascade order is not tracked)
            if (css.TryGetValue("border", out var border))
            {
                foreach (var side in sides)
                {
                    side.ApplyShorthand(border);
                }
            }
            ApplyPerSide(css, "border-style", sides, (side, value) => side.Style = value);
            ApplyPerSide(css, "border-width", sides, (side, value) => side.SetWidth(value));
            ApplyPerSide(css, "border-color", sides, (side, value) => side.Color = value);
            for (var i = 0; i < 4; i++)
            {
                var name = "border-" + Sides[i];
                if (css.TryGetValue(name, out var shorthand))
                {
                    sides[i].ApplyShorthand(shorthand);
                }
                if (css.TryGetValue(name + "-style", out var style))
                {
                    sides[i].Specified = true;
                    sides[i].Style = style.Trim();
                }
                if (css.TryGetValue(name + "-width", out var width))
                {
                    sides[i].SetWidth(width);
                }
                if (css.TryGetValue(name + "-color", out var color))
                {
                    sides[i].Specified = true;
                    sides[i].Color = color.Trim();
                }
            }

            return new CssBoxBorders(sides[0].Resolve(), sides[1].Resolve(), sides[2].Resolve(), sides[3].Resolve());
        }

        /// <summary>
        /// Parses <c>padding</c> and <c>padding-top/right/bottom/left</c> to twips (top, right, bottom, left). Null = not specified.
        /// </summary>
        public static int?[] ParsePadding(Dictionary<string, string> css)
        {
            var result = new int?[4];
            if (css.TryGetValue("padding", out var padding))
            {
                var values = ExpandSides(padding);
                for (var i = 0; i < 4; i++)
                {
                    result[i] = CssParser.ParseLengthToTwips(values[i]);
                }
            }
            for (var i = 0; i < 4; i++)
            {
                if (css.TryGetValue("padding-" + Sides[i], out var side))
                {
                    result[i] = CssParser.ParseLengthToTwips(side);
                }
            }
            return result;
        }

        /// <summary>
        /// Expands the css 1-4 value syntax to (top, right, bottom, left).
        /// </summary>
        public static string[] ExpandSides(string value)
        {
            var tokens = Tokenize(value);
            return tokens.Count switch
            {
                0 => [null, null, null, null],
                1 => [tokens[0], tokens[0], tokens[0], tokens[0]],
                2 => [tokens[0], tokens[1], tokens[0], tokens[1]],
                3 => [tokens[0], tokens[1], tokens[2], tokens[1]],
                _ => [tokens[0], tokens[1], tokens[2], tokens[3]],
            };
        }

        private static void ApplyPerSide(Dictionary<string, string> css, string name, BorderSide[] sides, Action<BorderSide, string> apply)
        {
            if (!css.TryGetValue(name, out var value))
            {
                return;
            }
            var values = ExpandSides(value);
            for (var i = 0; i < 4; i++)
            {
                if (values[i] != null)
                {
                    sides[i].Specified = true;
                    apply(sides[i], values[i]);
                }
            }
        }

        /// <summary>
        /// Splits on whitespace, keeping functions like <c>rgb(1, 2, 3)</c> together.
        /// </summary>
        private static List<string> Tokenize(string value)
        {
            var tokens = new List<string>();
            var current = new StringBuilder();
            var depth = 0;
            foreach (var c in value ?? string.Empty)
            {
                if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth = Math.Max(0, depth - 1);
                }

                if (char.IsWhiteSpace(c) && depth == 0)
                {
                    if (current.Length > 0)
                    {
                        tokens.Add(current.ToString());
                        current.Clear();
                    }
                }
                else
                {
                    current.Append(c);
                }
            }
            if (current.Length > 0)
            {
                tokens.Add(current.ToString());
            }
            return tokens;
        }

        private static BorderValues? ParseBorderStyle(string value)
        {
            return value?.Trim().ToLowerInvariant() switch
            {
                "none" or "hidden" => BorderValues.Nil,
                "solid" => BorderValues.Single,
                "dashed" => BorderValues.Dashed,
                "dotted" => BorderValues.Dotted,
                "double" => BorderValues.Double,
                "groove" => BorderValues.ThreeDEngrave,
                "ridge" => BorderValues.ThreeDEmboss,
                "inset" => BorderValues.Inset,
                "outset" => BorderValues.Outset,
                _ => null
            };
        }

        private static double? ParseBorderWidthPx(string value)
        {
            return value?.Trim().ToLowerInvariant() switch
            {
                "thin" => 1,
                "medium" => 3,
                "thick" => 5,
                _ => CssParser.ParseLengthToPixels(value)
            };
        }

        private sealed class BorderSide
        {
            public bool Specified { get; set; }
            public string Style { get; set; }
            public double? WidthPx { get; private set; }
            public string Color { get; set; }

            public void SetWidth(string value)
            {
                Specified = true;
                WidthPx = ParseBorderWidthPx(value);
            }

            /// <summary>
            /// A shorthand resets all three properties, like in css.
            /// </summary>
            public void ApplyShorthand(string value)
            {
                Specified = true;
                Style = null;
                WidthPx = null;
                Color = null;
                foreach (var token in Tokenize(value))
                {
                    if (ParseBorderStyle(token) != null)
                    {
                        Style = token;
                    }
                    else if (ParseBorderWidthPx(token) is { } width)
                    {
                        WidthPx = width;
                    }
                    else if (CssParser.ParseColor(token) != null)
                    {
                        Color = token;
                    }
                }
            }

            public CssBorder Resolve()
            {
                if (!Specified)
                {
                    return null;
                }
                // like css: without a style there is no border
                var style = ParseBorderStyle(Style);
                if (style is null || style == BorderValues.Nil || WidthPx is <= 0)
                {
                    return CssBorder.None;
                }
                return new CssBorder(style.Value, CssBorder.ToEighthPoints(WidthPx ?? 3), CssParser.ParseColor(Color) ?? "auto");
            }
        }
    }
}

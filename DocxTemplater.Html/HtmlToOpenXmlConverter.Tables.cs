using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html
{
    /// <summary>
    /// HTML tables: layout (colspan / rowspan), column widths, borders, padding, spacing and row heights.
    /// </summary>
    /// <remarks>
    /// Borders follow two modes. If the html contains no border information at all, the table uses the template
    /// table style (or a simple grid if there is none) - this is what most templates want. As soon as the html
    /// contains border information (the <c>border</c>, <c>frame</c> or <c>rules</c> attribute, or css borders on the
    /// table, a row or a cell), the html defines every line like a browser would, overriding the table style.
    /// So <c>border="0"</c> really means "no lines".
    /// </remarks>
    internal sealed partial class HtmlToOpenXmlConverter
    {
        private const int MinColumnWidthTwips = 300;
        private const int MinGridColumnTwips = 20;
        private const int TwipsPerPixel = 15;

        // Word's default left/right cell margin (0.08") when the html sets no padding
        private const int DefaultCellMarginTwips = 108;

        // average character width of a proportional font in em - deliberately a little generous, so an
        // estimated column rather wraps a bit later than too early
        private const double AverageCharWidthEm = 0.55;

        /// <summary>
        /// A grid position of a table row. <see cref="Cell"/> is null for the continuation of a rowspan
        /// (then <see cref="Origin"/> is the spanning cell) and for padding cells of short rows.
        /// </summary>
        private sealed record CellSlot(IElement Cell, IElement Origin, int Column, int ColumnSpan, bool MergeRestart, bool MergeContinue, bool MergeLast)
        {
            public IElement Source => Cell ?? Origin;

            /// <summary>
            /// The top edge of the (possibly merged) cell is in this row.
            /// </summary>
            public bool IsTopEdge => !MergeContinue;

            /// <summary>
            /// The bottom edge of the (possibly merged) cell is in this row.
            /// </summary>
            public bool IsBottomEdge => !MergeRestart && (!MergeContinue || MergeLast);
        }

        /// <summary>
        /// The borders defined by the html for the whole table. Null if the html defines no borders.
        /// </summary>
        private sealed record TableBorderModel(CssBoxBorders Frame, CssBorder InsideHorizontal, CssBorder InsideVertical);

        private void WriteTable(IElement table, Dictionary<string, string> css)
        {
            FinishParagraph();

            var headRows = new List<IElement>();
            var bodyRows = new List<IElement>();
            var footRows = new List<IElement>();
            IElement caption = null;
            foreach (var child in table.Children)
            {
                switch (child.LocalName.ToLowerInvariant())
                {
                    case "caption":
                        caption = child;
                        break;
                    case "thead":
                        headRows.AddRange(child.Children.Where(IsRow));
                        break;
                    case "tfoot":
                        footRows.AddRange(child.Children.Where(IsRow));
                        break;
                    case "tbody":
                        bodyRows.AddRange(child.Children.Where(IsRow));
                        break;
                    case "tr":
                        bodyRows.Add(child);
                        break;
                }
            }

            if (caption != null)
            {
                HandleCaption(caption);
            }

            var rows = headRows.Concat(bodyRows).Concat(footRows).ToList();
            if (rows.Count == 0)
            {
                return;
            }

            var layout = LayoutTable(rows, out var columnCount);
            if (columnCount == 0)
            {
                return;
            }

            var rowStyles = rows.Select(r => CssParser.ParseStyleAttribute(r.GetAttribute("style"))).ToList();
            var borderModel = CreateBorderModel(table, css, rows, rowStyles);
            var columnWidths = ComputeColumnWidths(table, css, layout, columnCount, out var tableWidth, out var autoColumns);

            var wordTable = new Table();
            wordTable.AppendChild(CreateTableProperties(table, css, borderModel, tableWidth));

            var grid = new TableGrid();
            foreach (var width in columnWidths)
            {
                grid.AppendChild(new GridColumn { Width = width.ToString() });
            }
            wordTable.AppendChild(grid);

            for (var r = 0; r < rows.Count; r++)
            {
                var row = new TableRow();
                var rowProperties = CreateRowProperties(rows[r], rowStyles[r], layout[r], isHeader: r < headRows.Count);
                if (rowProperties != null)
                {
                    row.AppendChild(rowProperties);
                }

                foreach (var slot in layout[r])
                {
                    row.AppendChild(CreateCell(slot, r, rows.Count, columnCount, columnWidths, autoColumns, rows[r], rowStyles[r], borderModel));
                }
                wordTable.AppendChild(row);
            }

            // Word merges two tables without a paragraph in between into one table
            if (m_container.LastChild is Table)
            {
                m_container.AppendChild(CreateTableSeparator());
            }
            m_container.AppendChild(wordTable);
        }

        /// <summary>
        /// A paragraph that is as small as possible (1pt, no spacing). Word joins two tables that directly follow each
        /// other into one table, so adjacent tables need a paragraph between them to stay separate like in a browser.
        /// </summary>
        internal static Paragraph CreateTableSeparator()
        {
            return new Paragraph(new ParagraphProperties
            {
                SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0", Line = "20", LineRule = LineSpacingRuleValues.Exact },
                ParagraphMarkRunProperties = new ParagraphMarkRunProperties(new FontSize { Val = "2" }, new FontSizeComplexScript { Val = "2" })
            });
        }

        private void HandleCaption(IElement caption)
        {
            var css = CssParser.ParseStyleAttribute(caption.GetAttribute("style"));
            var block = m_blockFormats.Peek() with { Alignment = JustificationValues.Center, IsPlain = false };
            WriteBlock(caption, css, block, CurrentRunFormat);
        }

        private static bool IsRow(IElement element)
        {
            return element.LocalName.Equals("tr", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCell(IElement element)
        {
            return element.LocalName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
                   element.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase);
        }

        #region Layout

        /// <summary>
        /// Places the html cells on a grid, resolving colspan and rowspan. Word models a rowspan
        /// as a vertically merged cell that is continued in every following row it spans.
        /// </summary>
        private static List<List<CellSlot>> LayoutTable(List<IElement> rows, out int columnCount)
        {
            var layout = new List<List<CellSlot>>();
            // per grid column: remaining rows of a rowspan, the colspan and the spanning cell
            var pending = new List<(int Remaining, int ColumnSpan, IElement Origin)>();
            columnCount = 0;

            for (var r = 0; r < rows.Count; r++)
            {
                var slots = new List<CellSlot>();
                var cells = new Queue<IElement>(rows[r].Children.Where(IsCell));
                var column = 0;
                while (true)
                {
                    if (column < pending.Count && pending[column].Remaining > 0)
                    {
                        var (remaining, span, origin) = pending[column];
                        slots.Add(new CellSlot(null, origin, column, span, false, true, remaining == 1));
                        pending[column] = (remaining - 1, span, origin);
                        column += span;
                        continue;
                    }

                    if (cells.Count > 0)
                    {
                        var cell = cells.Dequeue();
                        var columnSpan = Math.Clamp(ParseInt(cell.GetAttribute("colspan")) ?? 1, 1, 63);
                        var rowSpan = ParseInt(cell.GetAttribute("rowspan")) ?? 1;
                        var rowsLeft = rows.Count - r;
                        rowSpan = rowSpan <= 0 ? rowsLeft : Math.Min(rowSpan, rowsLeft);
                        slots.Add(new CellSlot(cell, null, column, columnSpan, rowSpan > 1, false, false));
                        if (rowSpan > 1)
                        {
                            while (pending.Count <= column)
                            {
                                pending.Add((0, 1, null));
                            }
                            pending[column] = (rowSpan - 1, columnSpan, cell);
                        }
                        column += columnSpan;
                        continue;
                    }

                    // no more html cells - fill gaps up to rowspans continuing further right
                    var nextPending = -1;
                    for (var c = column; c < pending.Count; c++)
                    {
                        if (pending[c].Remaining > 0)
                        {
                            nextPending = c;
                            break;
                        }
                    }
                    if (nextPending < 0)
                    {
                        break;
                    }
                    while (column < nextPending)
                    {
                        slots.Add(new CellSlot(null, null, column, 1, false, false, false));
                        column++;
                    }
                }
                columnCount = Math.Max(columnCount, column);
                layout.Add(slots);
            }

            // pad short rows so every row covers the full grid
            foreach (var slots in layout)
            {
                var covered = slots.Sum(x => x.ColumnSpan);
                for (var c = covered; c < columnCount; c++)
                {
                    slots.Add(new CellSlot(null, null, c, 1, false, false, false));
                }
            }
            return layout;
        }

        /// <summary>
        /// Column widths in twips from <c>&lt;col&gt;</c> / <c>&lt;colgroup&gt;</c> and the widths of single-column cells.
        /// Columns without a width share the remaining space.
        /// </summary>
        private int[] ComputeColumnWidths(IElement table, Dictionary<string, string> css, List<List<CellSlot>> layout, int columnCount,
            out TableWidth tableWidth, out bool[] autoColumns)
        {
            var absolute = new double?[columnCount];
            var percent = new double?[columnCount];

            void SetColumn(int column, string width)
            {
                if (column >= columnCount || absolute[column] != null || percent[column] != null)
                {
                    return;
                }
                if (CssParser.ParsePercentage(width) is { } p && p > 0)
                {
                    percent[column] = p;
                }
                else if (CssParser.ParseLengthToTwips(width) is { } twips && twips > 0)
                {
                    absolute[column] = twips;
                }
            }

            // 1. <col> and <colgroup>
            var column = 0;
            foreach (var child in table.Children)
            {
                var name = child.LocalName.ToLowerInvariant();
                if (name == "colgroup")
                {
                    var cols = child.Children.Where(x => x.LocalName.Equals("col", StringComparison.OrdinalIgnoreCase)).ToList();
                    if (cols.Count == 0)
                    {
                        // <colgroup span="2" width="100"> without <col> children
                        var span = Math.Clamp(ParseInt(child.GetAttribute("span")) ?? 1, 1, 1000);
                        var width = GetWidth(child);
                        for (var i = 0; i < span; i++)
                        {
                            SetColumn(column++, width);
                        }
                    }
                    foreach (var col in cols)
                    {
                        var span = Math.Clamp(ParseInt(col.GetAttribute("span")) ?? 1, 1, 1000);
                        var width = GetWidth(col) ?? GetWidth(child);
                        for (var i = 0; i < span; i++)
                        {
                            SetColumn(column++, width);
                        }
                    }
                }
                else if (name == "col")
                {
                    var span = Math.Clamp(ParseInt(child.GetAttribute("span")) ?? 1, 1, 1000);
                    var width = GetWidth(child);
                    for (var i = 0; i < span; i++)
                    {
                        SetColumn(column++, width);
                    }
                }
            }

            // 2. single-column cells, first row first (like browsers, explicit <col> widths win)
            foreach (var slot in layout.SelectMany(x => x).Where(x => x.Cell != null && x.ColumnSpan == 1))
            {
                SetColumn(slot.Column, GetWidth(slot.Cell));
            }

            var (tableTwips, tablePercent) = GetTableWidth(table, css);
            var allAbsolute = absolute.All(x => x != null);
            var fitContent = m_configuration.TablesWithoutWidth == HtmlTableWidth.FitContent;
            var shrinkToContent = false;
            tableWidth = null;
            autoColumns = new bool[columnCount];
            double total;
            if (tableTwips.HasValue)
            {
                total = tableTwips.Value;
                tableWidth = new TableWidth { Type = TableWidthUnitValues.Dxa, Width = tableTwips.Value.ToString() };
            }
            else if (tablePercent.HasValue)
            {
                total = m_availableWidthTwips * tablePercent.Value / 100;
                tableWidth = new TableWidth { Type = TableWidthUnitValues.Pct, Width = ((int)(tablePercent.Value * 50)).ToString() };
            }
            else if (allAbsolute)
            {
                // the columns define the table width
                total = Math.Min(absolute.Sum(x => x!.Value), m_availableWidthTwips);
                tableWidth = new TableWidth { Type = TableWidthUnitValues.Dxa, Width = ((int)total).ToString() };
            }
            else
            {
                // no width in the html: like a browser, the table is only as wide as its content (at most the text width)
                total = m_availableWidthTwips;
                shrinkToContent = fitContent;
            }

            var widths = new double[columnCount];
            var known = 0.0;
            var unknownColumns = new List<int>();
            for (var i = 0; i < columnCount; i++)
            {
                if (absolute[i] is { } a)
                {
                    widths[i] = a;
                }
                else if (percent[i] is { } p)
                {
                    widths[i] = total * p / 100;
                }
                else
                {
                    unknownColumns.Add(i);
                    continue;
                }
                known += widths[i];
            }

            if (unknownColumns.Count > 0 && !fitContent)
            {
                var share = Math.Max(MinColumnWidthTwips, (total - known) / unknownColumns.Count);
                foreach (var i in unknownColumns)
                {
                    widths[i] = share;
                }
            }
            else if (unknownColumns.Count > 0)
            {
                // columns without a width are sized by their content (min = longest word, max = longest line)
                var (min, max) = EstimateContentWidths(table, layout, columnCount);
                var space = Math.Max(0, total - known);
                var sumMin = unknownColumns.Sum(i => min[i]);
                var sumMax = unknownColumns.Sum(i => max[i]);
                foreach (var i in unknownColumns)
                {
                    if (sumMax <= space && shrinkToContent)
                    {
                        // everything fits: each column as wide as its content, Word's autofit takes it from there
                        widths[i] = max[i];
                        autoColumns[i] = true;
                    }
                    else if (sumMax <= space)
                    {
                        // a fixed table width with room to spare: grow the columns in proportion to their content
                        widths[i] = sumMax > 0 ? max[i] + ((space - sumMax) * max[i] / sumMax) : space / unknownColumns.Count;
                    }
                    else if (sumMin < space)
                    {
                        // not enough room for every line: wrap, but give each column at least its longest word
                        widths[i] = min[i] + ((max[i] - min[i]) * (space - sumMin) / (sumMax - sumMin));
                    }
                    else
                    {
                        widths[i] = sumMin > 0 ? min[i] * space / sumMin : space / unknownColumns.Count;
                    }
                }
                if (shrinkToContent && sumMax <= space)
                {
                    tableWidth = new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" };
                }
            }
            else if (known > 0 && Math.Abs(known - total) > 1)
            {
                // all columns are defined but do not add up to the table width - scale them proportionally
                var factor = total / known;
                for (var i = 0; i < columnCount; i++)
                {
                    widths[i] *= factor;
                }
            }

            tableWidth ??= new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" };
            return widths.Select(x => Math.Max(MinGridColumnTwips, (int)Math.Round(x))).ToArray();
        }

        /// <summary>
        /// Estimates the minimum (longest word) and maximum (longest line) content width of every column in twips,
        /// from the single-column cells. Word re-measures autofit tables itself, the estimate is what other
        /// applications (e.g. LibreOffice) use and what decides wrapping when the table does not fit.
        /// </summary>
        private (double[] Min, double[] Max) EstimateContentWidths(IElement table, List<List<CellSlot>> layout, int columnCount)
        {
            var min = new double[columnCount];
            var max = new double[columnCount];
            var cellPadding = ParseInt(table.GetAttribute("cellpadding"));
            foreach (var slot in layout.SelectMany(x => x).Where(x => x.Cell != null && x.ColumnSpan == 1))
            {
                var (cellMin, cellMax) = EstimateCellWidth(slot.Cell, cellPadding);
                min[slot.Column] = Math.Max(min[slot.Column], cellMin);
                max[slot.Column] = Math.Max(max[slot.Column], cellMax);
            }
            return (min, max);
        }

        private (double Min, double Max) EstimateCellWidth(IElement cell, int? tableCellPadding)
        {
            var css = CssParser.ParseStyleAttribute(cell.GetAttribute("style"));
            var fontSize = CssParser.ParseFontSize(css.TryGetValue("font-size", out var size) ? size : null, TemplateFontSize) ?? TemplateFontSize;
            var noWrap = cell.HasAttribute("nowrap")
                         || (css.TryGetValue("white-space", out var whiteSpace) && whiteSpace.Trim().ToLowerInvariant() is "nowrap" or "pre");

            var lines = new List<string>();
            var line = new System.Text.StringBuilder();
            var imageTwips = 0.0;

            void Flush()
            {
                var text = string.Join(' ', line.ToString().Split((char[])[' ', '\t', '\r', '\n', '\f'], StringSplitOptions.RemoveEmptyEntries));
                if (text.Length > 0)
                {
                    lines.Add(text);
                }
                line.Clear();
            }

            void Walk(INode node)
            {
                foreach (var child in node.ChildNodes)
                {
                    if (child is IText text)
                    {
                        line.Append(text.Data);
                        continue;
                    }
                    if (child is not IElement element)
                    {
                        continue;
                    }
                    var tag = element.LocalName.ToLowerInvariant();
                    if (SkippedElements.Contains(tag))
                    {
                        continue;
                    }
                    if (tag == "br")
                    {
                        Flush();
                    }
                    else if (tag == "img")
                    {
                        var imageCss = CssParser.ParseStyleAttribute(element.GetAttribute("style"));
                        var width = CssParser.ParseLengthToPixels(imageCss.TryGetValue("width", out var w) ? w : element.GetAttribute("width"));
                        imageTwips = Math.Max(imageTwips, (width ?? 0) * TwipsPerPixel);
                    }
                    else
                    {
                        var isBlock = IsBlockLevel(tag) || tag is "tr" or "li";
                        if (isBlock)
                        {
                            Flush();
                        }
                        Walk(element);
                        if (isBlock)
                        {
                            Flush();
                        }
                    }
                }
            }

            Walk(cell);
            Flush();

            var charTwips = fontSize * 20 * AverageCharWidthEm;
            var maxChars = lines.Count > 0 ? lines.Max(x => x.Length) : 0;
            var minChars = noWrap ? maxChars : lines.SelectMany(x => x.Split(' ')).Select(x => x.Length).DefaultIfEmpty(0).Max();

            var padding = CssBoxParser.ParsePadding(css);
            var horizontalPadding = padding[1] != null || padding[3] != null
                ? (padding[1] ?? 0) + (padding[3] ?? 0)
                : tableCellPadding.HasValue ? tableCellPadding.Value * TwipsPerPixel * 2 : DefaultCellMarginTwips * 2;

            return (Math.Max(minChars * charTwips, imageTwips) + horizontalPadding,
                    Math.Max(maxChars * charTwips, imageTwips) + horizontalPadding);
        }

        private static string GetWidth(IElement element)
        {
            var css = CssParser.ParseStyleAttribute(element.GetAttribute("style"));
            return css.TryGetValue("width", out var width) ? width : element.GetAttribute("width");
        }

        private (int? Twips, double? Percent) GetTableWidth(IElement table, Dictionary<string, string> css)
        {
            var width = css.TryGetValue("width", out var cssWidth) ? cssWidth : table.GetAttribute("width");
            if (CssParser.ParsePercentage(width) is { } percent)
            {
                return (null, Math.Clamp(percent, 1, 100));
            }
            if (CssParser.ParseLengthToTwips(width) is { } twips && twips > 0)
            {
                return (Math.Min(twips, m_availableWidthTwips), null);
            }
            return (null, null);
        }

        #endregion

        #region Table, row and cell properties

        private TableProperties CreateTableProperties(IElement table, Dictionary<string, string> css, TableBorderModel borderModel, TableWidth tableWidth)
        {
            var properties = new TableProperties();
            var tableStyle = FindTableStyle();
            if (tableStyle != null)
            {
                properties.TableStyle = new TableStyle { Val = tableStyle };
            }
            properties.TableWidth = tableWidth;

            var marginLeft = css.TryGetValue("margin-left", out var ml) ? ml.Trim() : null;
            var marginRight = css.TryGetValue("margin-right", out var mr) ? mr.Trim() : null;
            if (css.TryGetValue("margin", out var margin))
            {
                var sides = CssBoxParser.ExpandSides(margin);
                marginLeft ??= sides[3];
                marginRight ??= sides[1];
            }
            var tableAlign = marginLeft == "auto" && marginRight == "auto" ? "center" : table.GetAttribute("align");
            if (ParseAlignment(tableAlign) is { } alignment && alignment != JustificationValues.Both)
            {
                properties.TableJustification = new TableJustification
                {
                    Val = alignment == JustificationValues.Center ? TableRowAlignmentValues.Center
                        : alignment == JustificationValues.Right ? TableRowAlignmentValues.Right
                        : TableRowAlignmentValues.Left
                };
            }
            else if (marginLeft != "auto" && CssParser.ParseLengthToTwips(marginLeft) is { } indent && indent > 0)
            {
                properties.TableIndentation = new TableIndentation { Width = indent, Type = TableWidthUnitValues.Dxa };
            }

            var spacingPx = ParseInt(table.GetAttribute("cellspacing"))
                            ?? (css.TryGetValue("border-spacing", out var borderSpacing) ? (int?)CssParser.ParseLengthToPixels(CssBoxParser.ExpandSides(borderSpacing)[0]) : null);
            if (spacingPx is > 0)
            {
                properties.TableCellSpacing = new TableCellSpacing { Width = (spacingPx.Value * TwipsPerPixel).ToString(), Type = TableWidthUnitValues.Dxa };
            }

            if (borderModel != null)
            {
                properties.TableBorders = new TableBorders(
                    ToBorder<TopBorder>(borderModel.Frame.Top),
                    ToBorder<LeftBorder>(borderModel.Frame.Left),
                    ToBorder<BottomBorder>(borderModel.Frame.Bottom),
                    ToBorder<RightBorder>(borderModel.Frame.Right),
                    ToBorder<InsideHorizontalBorder>(borderModel.InsideHorizontal),
                    ToBorder<InsideVerticalBorder>(borderModel.InsideVertical));
            }
            else if (tableStyle == null)
            {
                // no table style and no borders in the html - a simple grid
                var grid = CssBorder.Solid(0.67);
                properties.TableBorders = new TableBorders(
                    ToBorder<TopBorder>(grid), ToBorder<LeftBorder>(grid), ToBorder<BottomBorder>(grid), ToBorder<RightBorder>(grid),
                    ToBorder<InsideHorizontalBorder>(grid), ToBorder<InsideVerticalBorder>(grid));
            }

            var background = CssParser.ParseColor(table.GetAttribute("bgcolor"))
                             ?? (css.TryGetValue("background-color", out var bg) || css.TryGetValue("background", out bg) ? CssParser.ParseColor(bg) : null);
            if (background != null)
            {
                properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = background };
            }

            if (css.TryGetValue("table-layout", out var tableLayout) && tableLayout.Trim().Equals("fixed", StringComparison.OrdinalIgnoreCase))
            {
                properties.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };
            }

            if (ParseInt(table.GetAttribute("cellpadding")) is { } paddingPx && paddingPx >= 0)
            {
                var padding = (short)Math.Min(paddingPx * TwipsPerPixel, short.MaxValue);
                properties.TableCellMarginDefault = new TableCellMarginDefault(
                    new TopMargin { Width = padding.ToString(), Type = TableWidthUnitValues.Dxa },
                    new TableCellLeftMargin { Width = padding, Type = TableWidthValues.Dxa },
                    new BottomMargin { Width = padding.ToString(), Type = TableWidthUnitValues.Dxa },
                    new TableCellRightMargin { Width = padding, Type = TableWidthValues.Dxa });
            }
            return properties;
        }

        private static TableRowProperties CreateRowProperties(IElement row, Dictionary<string, string> rowCss, List<CellSlot> slots, bool isHeader)
        {
            var properties = new TableRowProperties();
            if (isHeader)
            {
                properties.AppendChild(new TableHeader());
            }

            // the row is at least as high as its highest height declaration
            var heights = slots.Where(x => x.Cell != null).Select(x => GetHeight(x.Cell)).Append(GetHeight(row, rowCss));
            var height = heights.Max();
            if (height is > 0)
            {
                properties.AppendChild(new TableRowHeight { Val = (uint)height.Value, HeightType = HeightRuleValues.AtLeast });
            }

            if (IsBreakAvoided(rowCss, "page-break-inside", "break-inside"))
            {
                properties.AppendChild(new CantSplit());
            }
            return properties.HasChildren ? properties : null;
        }

        private static int? GetHeight(IElement element, Dictionary<string, string> css = null)
        {
            css ??= CssParser.ParseStyleAttribute(element.GetAttribute("style"));
            var height = css.TryGetValue("height", out var cssHeight) ? cssHeight : element.GetAttribute("height");
            return CssParser.ParseLengthToTwips(height);
        }

        private TableCell CreateCell(CellSlot slot, int rowIndex, int rowCount, int columnCount, int[] columnWidths, bool[] autoColumns,
            IElement row, Dictionary<string, string> rowCss, TableBorderModel borderModel)
        {
            var cell = new TableCell();
            var properties = new TableCellProperties();
            var source = slot.Source;
            var css = source != null ? CssParser.ParseStyleAttribute(source.GetAttribute("style")) : new Dictionary<string, string>();

            var width = columnWidths.Skip(slot.Column).Take(slot.ColumnSpan).Sum();
            properties.TableCellWidth = autoColumns.Skip(slot.Column).Take(slot.ColumnSpan).All(x => x)
                ? new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = "0" }
                : new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() };

            if (slot.ColumnSpan > 1)
            {
                properties.GridSpan = new GridSpan { Val = slot.ColumnSpan };
            }
            if (slot.MergeRestart)
            {
                properties.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart };
            }
            else if (slot.MergeContinue)
            {
                properties.VerticalMerge = new VerticalMerge();
            }

            if (borderModel != null)
            {
                properties.TableCellBorders = CreateCellBorders(slot, rowIndex, rowCount, columnCount, css, rowCss, borderModel);
            }

            var background = source != null
                ? CssParser.ParseColor(source.GetAttribute("bgcolor")) ?? (css.TryGetValue("background-color", out var bg) || css.TryGetValue("background", out bg) ? CssParser.ParseColor(bg) : null)
                : null;
            background ??= CssParser.ParseColor(row.GetAttribute("bgcolor"))
                           ?? (rowCss.TryGetValue("background-color", out var rowBg) || rowCss.TryGetValue("background", out rowBg) ? CssParser.ParseColor(rowBg) : null);
            if (background != null)
            {
                properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = background };
            }

            // html nowrap is deliberately NOT mapped to Word's noWrap: a page cannot scroll sideways like a browser,
            // so noWrap only ever cuts text off. nowrap is used for the column width estimate instead
            // (EstimateCellWidth), and Word's autofit keeps the line unbroken whenever there is room.

            var padding = CssBoxParser.ParsePadding(css);
            if (padding.Any(x => x != null))
            {
                var margin = new TableCellMargin();
                if (padding[0] is { } top)
                {
                    margin.TopMargin = new TopMargin { Width = top.ToString(), Type = TableWidthUnitValues.Dxa };
                }
                if (padding[3] is { } left)
                {
                    margin.LeftMargin = new LeftMargin { Width = left.ToString(), Type = TableWidthUnitValues.Dxa };
                }
                if (padding[2] is { } bottom)
                {
                    margin.BottomMargin = new BottomMargin { Width = bottom.ToString(), Type = TableWidthUnitValues.Dxa };
                }
                if (padding[1] is { } right)
                {
                    margin.RightMargin = new RightMargin { Width = right.ToString(), Type = TableWidthUnitValues.Dxa };
                }
                properties.TableCellMargin = margin;
            }

            var verticalAlign = (css.TryGetValue("vertical-align", out var cssVerticalAlign) ? cssVerticalAlign : source?.GetAttribute("valign"))
                                ?? (rowCss.TryGetValue("vertical-align", out var rowVerticalAlign) ? rowVerticalAlign : row.GetAttribute("valign"));
            TableVerticalAlignmentValues? cellAlignment = verticalAlign?.Trim().ToLowerInvariant() switch
            {
                "top" => TableVerticalAlignmentValues.Top,
                "middle" or "center" => TableVerticalAlignmentValues.Center,
                "bottom" => TableVerticalAlignmentValues.Bottom,
                _ => null
            };
            if (cellAlignment.HasValue)
            {
                properties.TableCellVerticalAlignment = new TableCellVerticalAlignment { Val = cellAlignment.Value };
            }
            cell.AppendChild(properties);

            if (slot.Cell != null)
            {
                var tag = slot.Cell.LocalName.ToLowerInvariant();
                // the row alignment is the default for its cells
                var rowAlign = rowCss.TryGetValue("text-align", out var rowTextAlign) ? rowTextAlign : row.GetAttribute("align");
                var block = new BlockFormat { IsPlain = false, Alignment = ParseAlignment(rowAlign) };
                // padding, margins and background belong to the cell, not to its paragraphs
                var paragraphCss = css.Where(x => !x.Key.StartsWith("padding", StringComparison.OrdinalIgnoreCase)
                                                  && !x.Key.StartsWith("margin", StringComparison.OrdinalIgnoreCase)
                                                  && !x.Key.StartsWith("background", StringComparison.OrdinalIgnoreCase))
                    .ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);
                block = ApplyBlockCss(slot.Cell, paragraphCss, block);
                var format = ApplyInlineCss(css, ApplyTagFormat(tag, slot.Cell, CurrentRunFormat), isBlock: true);
                RenderInto(cell, block, format, () => WalkChildren(slot.Cell));
            }

            // a table cell must end with a paragraph
            if (cell.LastChild is not Paragraph)
            {
                cell.AppendChild(new Paragraph());
            }
            return cell;
        }

        #endregion

        #region Borders

        /// <summary>
        /// Returns the border model if the html defines borders anywhere in the table, otherwise null (template style applies).
        /// </summary>
        private static TableBorderModel CreateBorderModel(IElement table, Dictionary<string, string> css, List<IElement> rows, List<Dictionary<string, string>> rowStyles)
        {
            var tableBorders = CssBoxParser.ParseBorders(css);
            var hasBorderAttribute = table.HasAttribute("border");
            var htmlDefinesBorders = hasBorderAttribute || table.HasAttribute("frame") || table.HasAttribute("rules") || !tableBorders.IsEmpty
                                     || rowStyles.Any(x => !CssBoxParser.ParseBorders(x).IsEmpty)
                                     || rows.SelectMany(r => r.Children.Where(IsCell))
                                         .Any(c => !CssBoxParser.ParseBorders(CssParser.ParseStyleAttribute(c.GetAttribute("style"))).IsEmpty);
            if (!htmlDefinesBorders)
            {
                return null;
            }

            // <table border="n">: an n pixel frame and 1 pixel lines between all cells; border="" means 1
            var borderWidth = hasBorderAttribute ? ParseInt(table.GetAttribute("border")) ?? 1 : 0;
            var frame = table.GetAttribute("frame")?.Trim().ToLowerInvariant() ?? (borderWidth > 0 ? "box" : "void");
            var rules = table.GetAttribute("rules")?.Trim().ToLowerInvariant() ?? (borderWidth > 0 ? "all" : "none");
            var frameBorder = CssBorder.Solid(Math.Max(1, borderWidth));
            var ruleBorder = CssBorder.Solid(1);

            var top = frame is "above" or "hsides" or "box" or "border" ? frameBorder : CssBorder.None;
            var bottom = frame is "below" or "hsides" or "box" or "border" ? frameBorder : CssBorder.None;
            var left = frame is "lhs" or "vsides" or "box" or "border" ? frameBorder : CssBorder.None;
            var right = frame is "rhs" or "vsides" or "box" or "border" ? frameBorder : CssBorder.None;

            return new TableBorderModel(
                new CssBoxBorders(tableBorders.Top ?? top, tableBorders.Right ?? right, tableBorders.Bottom ?? bottom, tableBorders.Left ?? left),
                rules is "all" or "rows" ? ruleBorder : CssBorder.None,
                rules is "all" or "cols" ? ruleBorder : CssBorder.None);
        }

        /// <summary>
        /// Every side of a cell: the cell's css, then the row's css, then the table frame (outer edges) or rules (inner edges).
        /// </summary>
        private static TableCellBorders CreateCellBorders(CellSlot slot, int rowIndex, int rowCount, int columnCount,
            Dictionary<string, string> css, Dictionary<string, string> rowCss, TableBorderModel model)
        {
            var cellBorders = CssBoxParser.ParseBorders(css);
            var rowBorders = CssBoxParser.ParseBorders(rowCss);
            var isFirstColumn = slot.Column == 0;
            var isLastColumn = slot.Column + slot.ColumnSpan >= columnCount;

            // inner edges of a vertically merged cell are not drawn by Word - keep them empty
            var top = slot.IsTopEdge
                ? cellBorders.Top ?? rowBorders.Top ?? (rowIndex == 0 ? model.Frame.Top : model.InsideHorizontal)
                : CssBorder.None;
            var bottom = slot.IsBottomEdge
                ? cellBorders.Bottom ?? rowBorders.Bottom ?? (rowIndex == rowCount - 1 ? model.Frame.Bottom : model.InsideHorizontal)
                : CssBorder.None;
            var left = cellBorders.Left ?? (isFirstColumn ? rowBorders.Left : null) ?? (isFirstColumn ? model.Frame.Left : model.InsideVertical);
            var right = cellBorders.Right ?? (isLastColumn ? rowBorders.Right : null) ?? (isLastColumn ? model.Frame.Right : model.InsideVertical);

            return new TableCellBorders(
                ToBorder<TopBorder>(top),
                ToBorder<LeftBorder>(left),
                ToBorder<BottomBorder>(bottom),
                ToBorder<RightBorder>(right));
        }

        private static T ToBorder<T>(CssBorder border)
            where T : BorderType, new()
        {
            if (border == null || border.IsNone)
            {
                return new T { Val = BorderValues.Nil };
            }
            return new T { Val = border.Style, Size = border.Size, Space = 0, Color = border.Color };
        }

        #endregion

        private string FindTableStyle()
        {
            var style = (m_configuration.TableStyle != null ? m_mainDocumentPart.FindTableStyleByName(m_configuration.TableStyle) : null)
                        ?? m_mainDocumentPart.FindTableStyleByName("Table Grid");
            return style?.StyleId?.Value;
        }

        private static int? ParseInt(string value)
        {
            return int.TryParse(value?.Trim(), out var result) ? result : null;
        }
    }
}

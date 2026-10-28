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

        // the estimated text width is deliberately a little generous (fonts differ from the built-in metrics),
        // so a column is rather a bit too wide than wrapping a word or a number
        private const double TextWidthSafety = 1.1;
        private const double BoldWidthFactor = 1.06;

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

        /// <summary>
        /// A table row with its parsed style, as needed for its cells.
        /// </summary>
        private sealed record RowContext(IElement Row, Dictionary<string, string> Css, CssBoxBorders Borders, int Index, int RowCount, int ColumnCount);

        private readonly Dictionary<IElement, CssBoxBorders> m_cellBorders = new();

        private void WriteTable(IElement table, Dictionary<string, string> css)
        {
            FinishParagraph();
            m_tableDepth++;
            try
            {
                WriteTableContent(table, css);
            }
            finally
            {
                if (--m_tableDepth == 0)
                {
                    m_cssCache.Clear();
                    m_cellBorders.Clear();
                }
            }
        }

        private void WriteTableContent(IElement table, Dictionary<string, string> css)
        {

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

            // without <thead>, leading rows of only <th> cells are the header - repeated on every page like <thead>
            var headerRowCount = headRows.Count;
            if (headerRowCount == 0)
            {
                while (headerRowCount < rows.Count - 1 && IsHeaderRow(rows[headerRowCount]))
                {
                    headerRowCount++;
                }
            }

            var rowStyles = rows.Select(GetCss).ToList();
            var rowBorders = rowStyles.Select(CssBoxParser.ParseBorders).ToList();
            var borderModel = CreateBorderModel(table, css, rows, rowBorders);
            var columnWidths = ComputeColumnWidths(table, css, layout, columnCount, out var tableWidth, out var autoColumns, out var fixedLayout);

            var wordTable = new Table();
            var tableProperties = CreateTableProperties(table, css, borderModel, tableWidth, fixedLayout);
            wordTable.AppendChild(tableProperties);
            // with cell spacing every cell is drawn as its own box - keep the borders on each cell
            var bordersOnEveryCell = tableProperties.TableCellSpacing != null;

            var grid = new TableGrid();
            foreach (var width in columnWidths)
            {
                grid.AppendChild(new GridColumn { Width = width.ToString() });
            }
            wordTable.AppendChild(grid);

            for (var r = 0; r < rows.Count; r++)
            {
                var row = new TableRow();
                var rowProperties = CreateRowProperties(rows[r], rowStyles[r], layout[r], isHeader: r < headerRowCount);
                if (rowProperties != null)
                {
                    row.AppendChild(rowProperties);
                }

                var rowContext = new RowContext(rows[r], rowStyles[r], rowBorders[r], r, rows.Count, columnCount);
                foreach (var slot in layout[r])
                {
                    row.AppendChild(CreateCell(slot, rowContext, columnWidths, autoColumns, borderModel, bordersOnEveryCell, ParseInt(table.GetAttribute("cellpadding"))));
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
            var css = GetCss(caption);
            var block = m_blockFormats.Peek() with { Alignment = JustificationValues.Center, IsPlain = false };
            WriteBlock(caption, css, block, CurrentRunFormat);
        }

        private static bool IsRow(IElement element)
        {
            return element.LocalName.Equals("tr", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHeaderRow(IElement row)
        {
            var cells = row.Children.Where(IsCell).ToList();
            return cells.Count > 0 && cells.All(x => x.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase));
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
        /// <remarks>
        /// With <see cref="HtmlTableLayout.Fixed"/> the widths are final: like a browser, no column is narrower than its
        /// longest word or number (whatever width the html asks for), and a table that does not fit the text width takes
        /// the space from the columns with wrappable text first. Word's autofit would otherwise wrap numbers like "1.350"
        /// in narrow columns, and applications without Word's autofit could not reproduce the columns at all.
        /// </remarks>
        private int[] ComputeColumnWidths(IElement table, Dictionary<string, string> css, List<List<CellSlot>> layout, int columnCount,
            out TableWidth tableWidth, out bool[] autoColumns, out bool fixedLayout)
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

            var htmlFixedLayout = css.TryGetValue("table-layout", out var tableLayout) && tableLayout.Trim().Equals("fixed", StringComparison.OrdinalIgnoreCase);
            var finalWidths = m_configuration.TableLayout == HtmlTableLayout.Fixed;
            fixedLayout = finalWidths || htmlFixedLayout;
            List<MeasuredCell> measuredCells = null;
            (double[] Min, double[] Max)? contentWidths = null;
            (double[] Min, double[] Max) ContentWidths()
            {
                return contentWidths ??= GetContentWidths(measuredCells ??= MeasureCells(table, layout), columnCount);
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
                var (min, max) = ContentWidths();
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

            if (finalWidths)
            {
                if (!htmlFixedLayout)
                {
                    var (min, _) = ContentWidths();
                    for (var i = 0; i < columnCount; i++)
                    {
                        widths[i] = Math.Max(widths[i], min[i]);
                    }
                    // like in a browser the table grows beyond its html width when its content needs it - up to the text width
                    FitToWidth(widths, min, m_availableWidthTwips);
                    if (unknownColumns.Count > 0)
                    {
                        DistributeByContent(widths, unknownColumns, measuredCells ??= MeasureCells(table, layout), layout.Count);
                    }
                }
                var result = widths.Select(x => Math.Max(MinGridColumnTwips, (int)Math.Round(x))).ToArray();
                tableWidth = new TableWidth { Type = TableWidthUnitValues.Dxa, Width = result.Sum().ToString() };
                autoColumns = new bool[columnCount];
                return result;
            }

            tableWidth ??= new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" };
            return widths.Select(x => Math.Max(MinGridColumnTwips, (int)Math.Round(x))).ToArray();
        }

        /// <summary>
        /// Shrinks the columns to <paramref name="limit"/>: first the space columns have beyond their longest word
        /// (<paramref name="min"/>), then the widest columns - narrow columns of numbers keep their width as long as possible.
        /// </summary>
        private static void FitToWidth(double[] widths, double[] min, double limit)
        {
            var excess = widths.Sum() - limit;
            if (excess <= 1)
            {
                return;
            }

            var slack = widths.Select((w, i) => Math.Max(0, w - min[i])).ToArray();
            var totalSlack = slack.Sum();
            if (totalSlack >= excess)
            {
                for (var i = 0; i < widths.Length; i++)
                {
                    widths[i] -= slack[i] * excess / totalSlack;
                }
                return;
            }

            for (var i = 0; i < widths.Length; i++)
            {
                widths[i] -= slack[i];
            }

            // still too wide: cap the widest columns at the width that makes the table fit
            var sorted = widths.OrderBy(x => x).ToArray();
            var rest = limit;
            var cap = 0.0;
            for (var k = 0; k < sorted.Length; k++)
            {
                var remaining = sorted.Length - k;
                if (sorted[k] * remaining > rest)
                {
                    cap = rest / remaining;
                    break;
                }
                rest -= sorted[k];
            }
            for (var i = 0; i < widths.Length; i++)
            {
                widths[i] = Math.Min(widths[i], cap);
            }
        }

        /// <summary>
        /// A single-column cell with its measured text.
        /// </summary>
        private sealed record MeasuredCell(int Row, int Column, bool IsHeader, CellText Text);

        /// <summary>
        /// The minimum (longest word) and maximum (longest line) content width of every column in twips,
        /// from the single-column cells. This decides the final column widths (<see cref="HtmlTableLayout.Fixed"/>);
        /// with autofit it is what other applications (e.g. LibreOffice) use and what decides wrapping when the table does not fit.
        /// </summary>
        private static (double[] Min, double[] Max) GetContentWidths(List<MeasuredCell> cells, int columnCount)
        {
            var min = new double[columnCount];
            var max = new double[columnCount];
            foreach (var cell in cells)
            {
                min[cell.Column] = Math.Max(min[cell.Column], cell.Text.MinWidth);
                max[cell.Column] = Math.Max(max[cell.Column], cell.Text.MaxWidth);
            }
            return (min, max);
        }

        /// <summary>
        /// Splits the width of the <paramref name="columns"/> (the columns without a width in the html) by their content,
        /// so the table gets as few lines as possible (see <see cref="ColumnWidthOptimizer"/>). The table width stays.
        /// </summary>
        /// <remarks>
        /// A table that fits on one line including its header cells keeps every cell on one line. Otherwise header cells
        /// (<c>&lt;th&gt;</c>) are only measured in columns without other cells: they occur once and may wrap, the body
        /// rows decide the height of the table.
        /// </remarks>
        private static void DistributeByContent(double[] widths, List<int> columns, List<MeasuredCell> cells, int rowCount)
        {
            var space = columns.Sum(i => widths[i]);
            var position = new int[widths.Length];
            Array.Fill(position, -1);
            for (var k = 0; k < columns.Count; k++)
            {
                position[columns[k]] = k;
            }

            var (_, max) = GetContentWidths(cells, widths.Length);
            var oneLine = columns.Select(i => max[i]).ToArray();
            var sumOneLine = oneLine.Sum();
            double[] result;
            if (sumOneLine <= space && sumOneLine > 0)
            {
                result = oneLine.Select(x => x + ((space - sumOneLine) * x / sumOneLine)).ToArray();
            }
            else
            {
                var withBody = cells.Where(x => !x.IsHeader).Select(x => x.Column).ToHashSet();
                var measured = cells
                    .Where(x => position[x.Column] >= 0 && (!x.IsHeader || !withBody.Contains(x.Column)))
                    .Select(x => (x.Row, position[x.Column], x.Text))
                    .ToList();
                result = ColumnWidthOptimizer.Distribute(measured, columns.Count, rowCount, space);
            }
            if (result == null)
            {
                return;
            }
            for (var k = 0; k < columns.Count; k++)
            {
                widths[columns[k]] = result[k];
            }
        }

        private List<MeasuredCell> MeasureCells(IElement table, List<List<CellSlot>> layout)
        {
            var result = new List<MeasuredCell>();
            var cellPadding = ParseInt(table.GetAttribute("cellpadding"));
            var tableFontSize = CurrentRunFormat.FontSizePt ?? TemplateFontSize;
            var tableCss = GetCss(table);
            if (tableCss.TryGetValue("font-size", out var tableSize))
            {
                tableFontSize = CssParser.ParseFontSize(tableSize, tableFontSize) ?? tableFontSize;
            }
            var tableBold = IsBold(tableCss, CurrentRunFormat.Bold == true);

            IElement lastRow = null;
            var rowFontSize = tableFontSize;
            var rowBold = tableBold;
            for (var r = 0; r < layout.Count; r++)
            {
                foreach (var slot in layout[r].Where(x => x.Cell != null && x.ColumnSpan == 1))
                {
                    if (slot.Cell.ParentElement != lastRow)
                    {
                        lastRow = slot.Cell.ParentElement;
                        var rowCss = lastRow != null ? GetCss(lastRow) : EmptyCss;
                        rowFontSize = rowCss.TryGetValue("font-size", out var rowSize) ? CssParser.ParseFontSize(rowSize, tableFontSize) ?? tableFontSize : tableFontSize;
                        rowBold = IsBold(rowCss, tableBold);
                    }
                    var isHeader = slot.Cell.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase);
                    result.Add(new MeasuredCell(r, slot.Column, isHeader, MeasureCell(slot.Cell, cellPadding, rowFontSize, rowBold)));
                }
            }
            return result;
        }

        /// <summary>
        /// Measures the words and lines of a cell with built-in font metrics (<see cref="CharWidthEm"/>).
        /// </summary>
        private CellText MeasureCell(IElement cell, int? tableCellPadding, double rowFontSize, bool rowBold)
        {
            var css = GetCss(cell);
            var noWrap = cell.HasAttribute("nowrap")
                         || (css.TryGetValue("white-space", out var whiteSpace) && whiteSpace.Trim().ToLowerInvariant() is "nowrap" or "pre");
            var text = new CellTextBuilder();

            void AddText(string data, double fontSize, bool bold)
            {
                var twipsPerEm = fontSize * 20 * TextWidthSafety * (bold ? BoldWidthFactor : 1);
                var spaceWidth = CharWidthEm(' ') * twipsPerEm;
                foreach (var c in data)
                {
                    if (c is ' ' or '\t' or '\r' or '\n' or '\f')
                    {
                        text.AddSpace();
                        continue;
                    }
                    text.AddChar(CharWidthEm(c) * twipsPerEm, spaceWidth);
                }
            }

            void Walk(INode node, double fontSize, bool bold)
            {
                foreach (var child in node.ChildNodes)
                {
                    if (child is IText textNode)
                    {
                        AddText(textNode.Data, fontSize, bold);
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
                    var elementCss = GetCss(element);
                    if (tag == "br")
                    {
                        text.EndLine();
                    }
                    else if (tag == "img")
                    {
                        var width = CssParser.ParseLengthToPixels(elementCss.TryGetValue("width", out var w) ? w : element.GetAttribute("width"));
                        text.AddBox((width ?? 0) * TwipsPerPixel);
                    }
                    else
                    {
                        var childSize = tag switch
                        {
                            "small" => fontSize / 1.2,
                            "big" => fontSize * 1.2,
                            // Word draws sub- and superscript at about two thirds of the size
                            "sub" or "sup" => fontSize * 0.66,
                            _ => fontSize
                        };
                        if (elementCss.TryGetValue("font-size", out var size))
                        {
                            childSize = CssParser.ParseFontSize(size, childSize) ?? childSize;
                        }
                        var childBold = IsBold(elementCss, bold || tag is "b" or "strong" or "th");
                        var isBlock = IsBlockLevel(tag) || tag is "tr" or "li";
                        if (isBlock)
                        {
                            text.EndLine();
                        }
                        Walk(element, childSize, childBold);
                        if (isBlock)
                        {
                            text.EndLine();
                        }
                    }
                }
            }

            var fontSize = css.TryGetValue("font-size", out var cellSize) ? CssParser.ParseFontSize(cellSize, rowFontSize) ?? rowFontSize : rowFontSize;
            var isHeaderCell = cell.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase);
            Walk(cell, fontSize, IsBold(css, rowBold || isHeaderCell));

            return text.Build(GetHorizontalPadding(CssBoxParser.ParsePadding(css), tableCellPadding), noWrap);
        }

        /// <summary>
        /// Left plus right inner margin of a cell in twips: the cell's css padding, the table's cellpadding or the default margin.
        /// </summary>
        private int GetHorizontalPadding(int?[] padding, int? tableCellPadding)
        {
            return padding[1] != null || padding[3] != null
                ? (padding[1] ?? 0) + (padding[3] ?? 0)
                : tableCellPadding.HasValue ? tableCellPadding.Value * TwipsPerPixel * 2 : DefaultHorizontalCellMargin * 2;
        }

        /// <summary>
        /// Left (and right) margin of cells without padding in the html: <see cref="HtmlFormatterConfiguration.TableCellHorizontalMargin"/> or Word's default.
        /// </summary>
        private int DefaultHorizontalCellMargin => m_configuration.TableCellHorizontalMargin ?? DefaultCellMarginTwips;

        private static bool IsBold(Dictionary<string, string> css, bool inherited)
        {
            if (!css.TryGetValue("font-weight", out var weight))
            {
                return inherited;
            }
            weight = weight.Trim().ToLowerInvariant();
            if (weight is "bold" or "bolder")
            {
                return true;
            }
            if (weight is "normal" or "lighter")
            {
                return false;
            }
            return int.TryParse(weight, out var numeric) ? numeric >= 600 : inherited;
        }

        /// <summary>
        /// Width of a character in em, close to the metrics of common sans-serif fonts (Arial, Helvetica, Heebo).
        /// </summary>
        private static double CharWidthEm(char c)
        {
            if (c is >= '0' and <= '9')
            {
                return 0.556;
            }
            switch (c)
            {
                case ' ' or '.' or ',' or ':' or ';' or '!' or '\'' or '|':
                case 'f' or 't' or 'I':
                    return 0.278;
                case 'i' or 'j' or 'l':
                    return 0.222;
                case 'r' or '-' or '(' or ')' or '[' or ']' or '/' or '\\':
                    return 0.333;
                case 'm' or 'M':
                    return 0.833;
                case 'w':
                    return 0.722;
                case 'W':
                    return 0.944;
            }
            if (c >= 0x2E80)
            {
                // CJK and other full width characters
                return 1.0;
            }
            if (char.IsUpper(c))
            {
                return 0.667;
            }
            return char.IsLower(c) ? 0.556 : 0.6;
        }

        private string GetWidth(IElement element)
        {
            var css = GetCss(element);
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

        private TableProperties CreateTableProperties(IElement table, Dictionary<string, string> css, TableBorderModel borderModel, TableWidth tableWidth, bool fixedLayout)
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

            if (fixedLayout)
            {
                // Word uses the widths as written instead of re-running autofit
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
            else if (m_configuration.TableCellHorizontalMargin is { } cellMargin && cellMargin >= 0)
            {
                var horizontal = (short)Math.Min(cellMargin, short.MaxValue);
                properties.TableCellMarginDefault = new TableCellMarginDefault(
                    new TableCellLeftMargin { Width = horizontal, Type = TableWidthValues.Dxa },
                    new TableCellRightMargin { Width = horizontal, Type = TableWidthValues.Dxa });
            }
            return properties;
        }

        private TableRowProperties CreateRowProperties(IElement row, Dictionary<string, string> rowCss, List<CellSlot> slots, bool isHeader)
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

        private int? GetHeight(IElement element, Dictionary<string, string> css = null)
        {
            css ??= GetCss(element);
            var height = css.TryGetValue("height", out var cssHeight) ? cssHeight : element.GetAttribute("height");
            return CssParser.ParseLengthToTwips(height);
        }

        private TableCell CreateCell(CellSlot slot, RowContext row, int[] columnWidths, bool[] autoColumns, TableBorderModel borderModel,
            bool bordersOnEveryCell, int? tableCellPadding)
        {
            var cell = new TableCell();
            var properties = new TableCellProperties();
            var source = slot.Source;
            var css = source != null ? GetCss(source) : EmptyCss;
            var rowCss = row.Css;

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
                properties.TableCellBorders = CreateCellBorders(slot, row, source, borderModel, bordersOnEveryCell);
            }

            var background = source != null
                ? CssParser.ParseColor(source.GetAttribute("bgcolor")) ?? (css.TryGetValue("background-color", out var bg) || css.TryGetValue("background", out bg) ? CssParser.ParseColor(bg) : null)
                : null;
            background ??= CssParser.ParseColor(row.Row.GetAttribute("bgcolor"))
                           ?? (rowCss.TryGetValue("background-color", out var rowBg) || rowCss.TryGetValue("background", out rowBg) ? CssParser.ParseColor(rowBg) : null);
            if (background != null)
            {
                properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = background };
            }

            // html nowrap is deliberately NOT mapped to Word's noWrap: a page cannot scroll sideways like a browser,
            // so noWrap only ever cuts text off. nowrap is used for the column width estimate instead
            // (EstimateCellWidth), so the column is wide enough whenever there is room.

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
                                ?? (rowCss.TryGetValue("vertical-align", out var rowVerticalAlign) ? rowVerticalAlign : row.Row.GetAttribute("valign"));
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

            // the row alignment is the default for its cells
            var rowAlign = rowCss.TryGetValue("text-align", out var rowTextAlign) ? rowTextAlign : row.Row.GetAttribute("align");
            var block = new BlockFormat { IsPlain = false, Alignment = ParseAlignment(rowAlign) };
            var format = CurrentRunFormat;
            if (source != null)
            {
                // padding, margins and background belong to the cell, not to its paragraphs
                block = ApplyBlockCss(source, css, block, includeBox: false);
                format = ApplyInlineCss(css, ApplyTagFormat(source.LocalName.ToLowerInvariant(), source, format), isBlock: true);
            }

            // content inside the cell (nested tables, images) is limited to the cell
            var horizontalPadding = GetHorizontalPadding(padding, tableCellPadding);
            var previousAvailableWidth = m_availableWidthTwips;
            m_availableWidthTwips = Math.Max(MinColumnWidthTwips, width - horizontalPadding);
            try
            {
                RenderInto(cell, block, format, () =>
                {
                    if (slot.Cell != null)
                    {
                        WalkChildren(slot.Cell);
                    }
                    // a table cell must end with a paragraph - formatted like the text of the cell, so an empty
                    // cell is as high as a line of its text and not as a line of the template's Normal style
                    if (m_container.LastChild is not Paragraph)
                    {
                        EnsureParagraph();
                    }
                });
            }
            finally
            {
                m_availableWidthTwips = previousAvailableWidth;
            }
            return cell;
        }

        #endregion

        #region Borders

        /// <summary>
        /// Returns the border model if the html defines borders anywhere in the table, otherwise null (template style applies).
        /// </summary>
        private TableBorderModel CreateBorderModel(IElement table, Dictionary<string, string> css, List<IElement> rows, List<CssBoxBorders> rowBorders)
        {
            var tableBorders = CssBoxParser.ParseBorders(css);
            var hasBorderAttribute = table.HasAttribute("border");
            var htmlDefinesBorders = hasBorderAttribute || table.HasAttribute("frame") || table.HasAttribute("rules") || !tableBorders.IsEmpty
                                     || rowBorders.Any(x => !x.IsEmpty)
                                     || rows.SelectMany(r => r.Children.Where(IsCell)).Any(c => !GetCellBorders(c).IsEmpty);
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

        private CssBoxBorders GetCellBorders(IElement cell)
        {
            if (!m_cellBorders.TryGetValue(cell, out var borders))
            {
                borders = CssBoxParser.ParseBorders(GetCss(cell));
                m_cellBorders[cell] = borders;
            }
            return borders;
        }

        /// <summary>
        /// Every side of a cell: the cell's css, then the row's css, then the table frame (outer edges) or rules (inner edges).
        /// Only the sides that differ from the table borders (<c>tblBorders</c>) are written - for a table with the same lines
        /// everywhere the cells carry no borders at all. Null if no side differs.
        /// </summary>
        private TableCellBorders CreateCellBorders(CellSlot slot, RowContext row, IElement source, TableBorderModel model, bool bordersOnEveryCell)
        {
            var cellBorders = source != null ? GetCellBorders(source) : EmptyBorders;
            var rowBorders = row.Borders;
            var isFirstColumn = slot.Column == 0;
            var isLastColumn = slot.Column + slot.ColumnSpan >= row.ColumnCount;

            // what the table borders give this cell
            var tableTop = row.Index == 0 ? model.Frame.Top : model.InsideHorizontal;
            var tableBottom = row.Index == row.RowCount - 1 ? model.Frame.Bottom : model.InsideHorizontal;
            var tableLeft = isFirstColumn ? model.Frame.Left : model.InsideVertical;
            var tableRight = isLastColumn ? model.Frame.Right : model.InsideVertical;

            // inner edges of a vertically merged cell are not drawn by Word - keep them empty
            var top = slot.IsTopEdge ? cellBorders.Top ?? rowBorders.Top ?? tableTop : CssBorder.None;
            var bottom = slot.IsBottomEdge ? cellBorders.Bottom ?? rowBorders.Bottom ?? tableBottom : CssBorder.None;
            var left = cellBorders.Left ?? (isFirstColumn ? rowBorders.Left : null) ?? tableLeft;
            var right = cellBorders.Right ?? (isLastColumn ? rowBorders.Right : null) ?? tableRight;

            var borders = new TableCellBorders();
            if (bordersOnEveryCell || !SameBorder(top, tableTop))
            {
                borders.TopBorder = ToBorder<TopBorder>(top);
            }
            if (bordersOnEveryCell || !SameBorder(left, tableLeft))
            {
                borders.LeftBorder = ToBorder<LeftBorder>(left);
            }
            if (bordersOnEveryCell || !SameBorder(bottom, tableBottom))
            {
                borders.BottomBorder = ToBorder<BottomBorder>(bottom);
            }
            if (bordersOnEveryCell || !SameBorder(right, tableRight))
            {
                borders.RightBorder = ToBorder<RightBorder>(right);
            }
            return borders.HasChildren ? borders : null;
        }

        private static readonly CssBoxBorders EmptyBorders = new(null, null, null, null);

        private static bool SameBorder(CssBorder a, CssBorder b)
        {
            a ??= CssBorder.None;
            b ??= CssBorder.None;
            return (a.IsNone && b.IsNone) || a == b;
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

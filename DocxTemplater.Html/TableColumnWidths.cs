using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater.Html
{
    /// <summary>
    /// A word of a table cell: its width in twips, the width of the space before it (0 at the start of a line)
    /// and the widths of its characters (null for an image, which cannot be broken).
    /// </summary>
    internal readonly record struct CellWord(double Width, double SpaceBefore, double[] Chars);

    /// <summary>
    /// The measured text of a table cell: its lines (paragraphs and line breaks), each a list of words,
    /// plus the horizontal cell padding. Cells with the same measurement are equal, so the line counts of a column
    /// only need computing once per distinct text.
    /// </summary>
    internal sealed class CellText : IEquatable<CellText>
    {
        private readonly CellWord[][] m_lines;
        private readonly int m_hashCode;

        public CellText(CellWord[][] lines, double padding, double longestWord, double longestLine, bool noWrap)
        {
            m_lines = lines;
            Padding = padding;
            MinWidth = (noWrap ? longestLine : longestWord) + padding;
            MaxWidth = longestLine + padding;

            var hash = new HashCode();
            hash.Add(padding);
            foreach (var line in lines)
            {
                hash.Add(line.Length);
                foreach (var word in line)
                {
                    hash.Add(word.Width);
                }
            }
            m_hashCode = hash.ToHashCode();
        }

        public double Padding { get; }

        /// <summary>
        /// Width of the longest word (longest line for nowrap) including the padding.
        /// </summary>
        public double MinWidth { get; }

        /// <summary>
        /// Width of the longest line including the padding - the cell fits on one line from this width on.
        /// </summary>
        public double MaxWidth { get; }

        /// <summary>
        /// Number of lines in a column of <paramref name="width"/> twips: greedy word wrap like Word, a word longer
        /// than the line is broken by character.
        /// </summary>
        public int CountLines(double width)
        {
            var available = Math.Max(1, width - Padding);
            var total = 0;
            foreach (var line in m_lines)
            {
                var count = 1;
                var x = 0.0;
                foreach (var word in line)
                {
                    var needed = x > 0 ? x + word.SpaceBefore + word.Width : word.Width;
                    if (needed <= available)
                    {
                        x = needed;
                        continue;
                    }
                    if (x > 0)
                    {
                        count++;
                        x = 0;
                    }
                    if (word.Width <= available)
                    {
                        x = word.Width;
                    }
                    else if (word.Chars == null)
                    {
                        // an image wider than the column fills its line
                        x = available;
                    }
                    else
                    {
                        foreach (var c in word.Chars)
                        {
                            if (x > 0 && x + c > available)
                            {
                                count++;
                                x = 0;
                            }
                            x += c;
                        }
                    }
                }
                total += count;
            }
            return Math.Max(1, total);
        }

        public bool Equals(CellText other)
        {
            if (ReferenceEquals(this, other))
            {
                return true;
            }
            if (other == null || m_hashCode != other.m_hashCode || Padding != other.Padding || m_lines.Length != other.m_lines.Length)
            {
                return false;
            }
            for (var i = 0; i < m_lines.Length; i++)
            {
                var a = m_lines[i];
                var b = other.m_lines[i];
                if (a.Length != b.Length)
                {
                    return false;
                }
                for (var j = 0; j < a.Length; j++)
                {
                    if (a[j].Width != b[j].Width || a[j].SpaceBefore != b[j].SpaceBefore || !SameChars(a[j].Chars, b[j].Chars))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as CellText);
        }

        public override int GetHashCode()
        {
            return m_hashCode;
        }

        private static bool SameChars(double[] a, double[] b)
        {
            return a == null || b == null ? a == b : a.AsSpan().SequenceEqual(b);
        }
    }

    /// <summary>
    /// Collects the words of a table cell while its html is walked.
    /// </summary>
    internal sealed class CellTextBuilder
    {
        private readonly List<CellWord[]> m_lines = new();
        private readonly List<CellWord> m_line = new();
        private readonly List<double> m_chars = new();
        private double m_word;
        private double m_spaceBefore;
        private double m_lineWidth;
        private double m_longestWord;
        private double m_longestLine;
        private bool m_pendingSpace;

        public void AddSpace()
        {
            EndWord();
            m_pendingSpace = m_lineWidth > 0;
        }

        public void AddChar(double width, double spaceWidth)
        {
            if (m_pendingSpace)
            {
                m_lineWidth += spaceWidth;
                m_spaceBefore = spaceWidth;
                m_pendingSpace = false;
            }
            m_chars.Add(width);
            m_word += width;
            m_lineWidth += width;
        }

        /// <summary>
        /// Something that cannot be broken, e.g. an image.
        /// </summary>
        public void AddBox(double width)
        {
            EndWord();
            m_line.Add(new CellWord(width, 0, null));
            m_lineWidth += width;
            m_longestWord = Math.Max(m_longestWord, width);
        }

        public void EndWord()
        {
            if (m_chars.Count == 0)
            {
                return;
            }
            m_line.Add(new CellWord(m_word, m_spaceBefore, m_chars.ToArray()));
            m_longestWord = Math.Max(m_longestWord, m_word);
            m_chars.Clear();
            m_word = 0;
            m_spaceBefore = 0;
        }

        public void EndLine()
        {
            EndWord();
            m_longestLine = Math.Max(m_longestLine, m_lineWidth);
            if (m_line.Count > 0)
            {
                m_lines.Add(m_line.ToArray());
                m_line.Clear();
            }
            m_lineWidth = 0;
            m_pendingSpace = false;
        }

        public CellText Build(double padding, bool noWrap)
        {
            EndLine();
            return new CellText(m_lines.ToArray(), padding, m_longestWord, m_longestLine, noWrap);
        }
    }

    /// <summary>
    /// Splits the width of a fixed layout table between its columns so the table is as short as possible.
    /// </summary>
    /// <remarks>
    /// Every row is as high as its highest cell, so the height of a table is the sum over all rows of the line count
    /// of the row's highest cell. Starting from narrow columns, the optimiser repeatedly widens the column that saves
    /// the most row lines per twip, trying steps of 10, 20, 40, ... twips to get over plateaus (a column often only
    /// loses a line after a large widening). Once no widening saves a row line, it widens what saves the most cell
    /// lines (e.g. a number that is broken next to a higher text cell), and finally shares the rest by how far each
    /// column is from fitting its content on one line.
    /// The columns start at their longest word if all of them fit, otherwise at their cell margins. Equal widths are
    /// much worse for tables mixing one text column with many short numeric columns that do not fit unbroken.
    /// </remarks>
    internal static class ColumnWidthOptimizer
    {
        private const int MinTextWidthTwips = 10;
        private const int StepTwips = 10;

        /// <summary>
        /// Returns the column widths for <paramref name="space"/> twips, or null if the cell margins alone do not fit.
        /// </summary>
        /// <param name="cells">The cells used for the measurement: their layout row, column and text.</param>
        public static double[] Distribute(IReadOnlyList<(int Row, int Column, CellText Text)> cells, int columnCount, int rowCount, double space)
        {
            var texts = new List<CellText>[columnCount];
            var textIndex = new Dictionary<CellText, int>[columnCount];
            var entries = new List<(int Row, int Text)>[columnCount];
            for (var c = 0; c < columnCount; c++)
            {
                texts[c] = new List<CellText>();
                textIndex[c] = new Dictionary<CellText, int>();
                entries[c] = new List<(int Row, int Text)>();
            }
            foreach (var (row, column, text) in cells)
            {
                if (!textIndex[column].TryGetValue(text, out var index))
                {
                    index = texts[column].Count;
                    texts[column].Add(text);
                    textIndex[column][text] = index;
                }
                entries[column].Add((row, index));
            }

            var oneLine = texts.Select(x => x.Count > 0 ? x.Max(t => t.MaxWidth) : 0).ToArray();
            var sumOneLine = oneLine.Sum();
            if (sumOneLine <= space)
            {
                // everything fits on one line: each column its one-line width plus a share of the rest
                return sumOneLine > 0
                    ? oneLine.Select(x => x + ((space - sumOneLine) * x / sumOneLine)).ToArray()
                    : Enumerable.Repeat(space / columnCount, columnCount).ToArray();
            }

            var width = texts.Select(x => (int)Math.Ceiling(x.Count > 0 ? x.Max(t => t.Padding) : 0) + MinTextWidthTwips).ToArray();
            // if every word fits, no word is broken: e.g. a single row with a long text would otherwise break the
            // numbers next to it into one character per line, as that does not make the row any higher
            var longestWord = texts.Select((x, c) => Math.Max(width[c], (int)Math.Ceiling(x.Count > 0 ? x.Max(t => t.MinWidth) : 0))).ToArray();
            if (longestWord.Sum() <= space)
            {
                width = longestWord;
            }
            var left = (int)Math.Floor(space) - width.Sum();
            if (left < 0)
            {
                return null;
            }

            // line counts per column and width, for the distinct texts of the column
            var cache = new Dictionary<int, int[]>[columnCount];
            int[] LinesAt(int column, int w)
            {
                cache[column] ??= new Dictionary<int, int[]>();
                if (!cache[column].TryGetValue(w, out var result))
                {
                    result = texts[column].Select(t => t.CountLines(w)).ToArray();
                    cache[column][w] = result;
                }
                return result;
            }

            // current line counts (row x column), and per row the highest count, how many cells have it and the second highest
            var lines = new int[rowCount, columnCount];
            var max1 = new int[rowCount];
            var count1 = new int[rowCount];
            var max2 = new int[rowCount];
            void UpdateRow(int r)
            {
                int first = 0, count = 0, second = 0;
                for (var c = 0; c < columnCount; c++)
                {
                    var value = lines[r, c];
                    if (value > first)
                    {
                        second = first;
                        first = value;
                        count = 1;
                    }
                    else if (value == first)
                    {
                        count++;
                    }
                    else if (value > second)
                    {
                        second = value;
                    }
                }
                max1[r] = first;
                count1[r] = count;
                max2[r] = count > 1 ? first : second;
            }

            for (var c = 0; c < columnCount; c++)
            {
                var current = LinesAt(c, width[c]);
                foreach (var (row, text) in entries[c])
                {
                    lines[row, c] = current[text];
                }
            }
            var usedRows = entries.SelectMany(x => x.Select(e => e.Row)).Distinct().ToArray();
            foreach (var r in usedRows)
            {
                UpdateRow(r);
            }

            // the narrowest widening (in steps of 10 twips) with the same line counts as the given one, so a large step
            // that jumps a plateau does not take more width than the column needs; line counts only drop with the width
            int Trim(int column, int widening)
            {
                var target = LinesAt(column, width[column] + widening).Sum();
                int low = 0, high = (widening - 1) / StepTwips;
                while (low < high)
                {
                    var middle = (low + high) / 2;
                    if (LinesAt(column, width[column] + ((middle + 1) * StepTwips)).Sum() == target)
                    {
                        high = middle;
                    }
                    else
                    {
                        low = middle + 1;
                    }
                }
                return Math.Min(widening, (low + 1) * StepTwips);
            }

            while (left > 0)
            {
                (double Ratio, int Column, int Step) bestRow = (0, -1, 0);
                (double Ratio, int Column, int Step) bestCell = (0, -1, 0);
                for (var c = 0; c < columnCount; c++)
                {
                    if (width[c] >= oneLine[c])
                    {
                        continue;
                    }
                    for (var step = StepTwips; ; step *= 2)
                    {
                        var widening = Trim(c, Math.Min(step, left));
                        var next = LinesAt(c, width[c] + widening);
                        long rowSaved = 0, cellSaved = 0;
                        foreach (var (row, text) in entries[c])
                        {
                            var current = lines[row, c];
                            var value = next[text];
                            if (value == current)
                            {
                                continue;
                            }
                            cellSaved += current - value;
                            var others = current == max1[row] && count1[row] == 1 ? max2[row] : max1[row];
                            rowSaved += max1[row] - Math.Max(others, value);
                        }
                        if (rowSaved > 0 && (double)rowSaved / widening > bestRow.Ratio)
                        {
                            bestRow = ((double)rowSaved / widening, c, widening);
                        }
                        if (cellSaved > 0 && (double)cellSaved / widening > bestCell.Ratio)
                        {
                            bestCell = ((double)cellSaved / widening, c, widening);
                        }
                        if (step >= left || width[c] + step >= oneLine[c])
                        {
                            break;
                        }
                    }
                }

                var (_, column, chosenStep) = bestRow.Column >= 0 ? bestRow : bestCell;
                if (column < 0)
                {
                    break;
                }
                width[column] += chosenStep;
                left -= chosenStep;
                var updated = LinesAt(column, width[column]);
                foreach (var (row, text) in entries[column])
                {
                    lines[row, column] = updated[text];
                    UpdateRow(row);
                }
            }

            var result = width.Select(x => (double)x).ToArray();
            if (left > 0)
            {
                // nothing saves a line anymore: share the rest by how far each column is from its one-line width
                var deficit = result.Select((w, c) => Math.Max(0, oneLine[c] - w)).ToArray();
                var sumDeficit = deficit.Sum();
                for (var c = 0; c < columnCount; c++)
                {
                    result[c] += sumDeficit > 0 ? left * deficit[c] / sumDeficit : (double)left / columnCount;
                }
            }
            return result;
        }
    }
}

using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html.Test
{
    internal partial class HtmlFormatterTest
    {
        private static Styles TableGridStyle()
        {
            return new Styles(new Style(new StyleName { Val = "Table Grid" }) { Type = StyleValues.Table, StyleId = "TableGrid" });
        }

        private static Table RenderTable(string html, Styles styles = null)
        {
            return Render("{{ds}:html}", html, styles: styles ?? TableGridStyle()).Elements<Table>().Single();
        }

        private static BorderValues? Val(BorderType border)
        {
            return border?.Val?.Value;
        }

        [Test]
        public void TableWithoutBorderInformationUsesTemplateStyle()
        {
            var table = RenderTable("<table><tr><td>a</td><td>b</td></tr></table>");

            Assert.That(table.GetFirstChild<TableProperties>().TableStyle.Val.Value, Is.EqualTo("TableGrid"));
            Assert.That(table.GetFirstChild<TableProperties>().TableBorders, Is.Null, "the style defines the borders");
            Assert.That(table.Descendants<TableCellBorders>(), Is.Empty);
        }

        [Test]
        public void BorderZeroRemovesAllBordersEvenWithTableStyle()
        {
            var table = RenderTable("<table border=\"0\"><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>");

            var borders = table.GetFirstChild<TableProperties>().TableBorders;
            Assert.That(borders.ChildElements.Cast<BorderType>().Select(Val), Is.All.EqualTo(BorderValues.Nil));
            var cellBorders = table.Descendants<TableCellBorders>().ToList();
            Assert.That(cellBorders, Has.Count.EqualTo(4));
            Assert.That(cellBorders.SelectMany(x => x.ChildElements.Cast<BorderType>()).Select(Val), Is.All.EqualTo(BorderValues.Nil));
        }

        [Test]
        public void BorderAttributeDrawsFrameAndCellLines()
        {
            var table = RenderTable("<table border=\"2\"><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>");

            var borders = table.GetFirstChild<TableProperties>().TableBorders;
            Assert.That(borders.TopBorder.Size.Value, Is.EqualTo(12u), "2px frame = 12 eighths of a point");
            Assert.That(borders.InsideHorizontalBorder.Size.Value, Is.EqualTo(6u), "1px lines between cells");
            Assert.That(borders.InsideVerticalBorder.Val.Value, Is.EqualTo(BorderValues.Single));

            var firstCell = table.Descendants<TableCell>().First().TableCellProperties.TableCellBorders;
            Assert.That(firstCell.TopBorder.Size.Value, Is.EqualTo(12u), "outer edge uses the frame");
            Assert.That(firstCell.RightBorder.Size.Value, Is.EqualTo(6u), "inner edge uses the rules");
        }

        [Test]
        public void FrameAndRulesAttributes()
        {
            var table = RenderTable("<table frame=\"hsides\" rules=\"rows\"><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>");

            var borders = table.GetFirstChild<TableProperties>().TableBorders;
            Assert.That(Val(borders.TopBorder), Is.EqualTo(BorderValues.Single));
            Assert.That(Val(borders.BottomBorder), Is.EqualTo(BorderValues.Single));
            Assert.That(Val(borders.LeftBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(Val(borders.InsideHorizontalBorder), Is.EqualTo(BorderValues.Single));
            Assert.That(Val(borders.InsideVerticalBorder), Is.EqualTo(BorderValues.Nil));
        }

        [Test]
        public void CssBordersOnCellsRowsAndTable()
        {
            var html = """
                       <table style="border-collapse: collapse; border: 2px dashed rgb(0, 0, 255)">
                         <tr style="border-bottom: 1px solid #ccc"><td style="border: 1px solid red; border-left: none">a</td><td>b</td></tr>
                         <tr><td>c</td><td style="border-top: 3px double #00ff00">d</td></tr>
                       </table>
                       """;
            var table = RenderTable(html);
            var cells = table.Descendants<TableCell>().Select(c => c.TableCellProperties.TableCellBorders).ToList();

            Assert.That(table.GetFirstChild<TableProperties>().TableBorders.TopBorder.Val.Value, Is.EqualTo(BorderValues.Dashed));
            Assert.That(table.GetFirstChild<TableProperties>().TableBorders.TopBorder.Color.Value, Is.EqualTo("0000FF"));

            // cell a: own border, left side removed
            Assert.That(cells[0].TopBorder.Color.Value, Is.EqualTo("FF0000"));
            Assert.That(Val(cells[0].LeftBorder), Is.EqualTo(BorderValues.Nil));
            // cell b: top from the table frame, bottom from the row
            Assert.That(cells[1].TopBorder.Val.Value, Is.EqualTo(BorderValues.Dashed));
            Assert.That(cells[1].BottomBorder.Color.Value, Is.EqualTo("CCCCCC"));
            // between cells without own borders: no rules -> nothing (like a browser)
            Assert.That(Val(cells[2].RightBorder), Is.EqualTo(BorderValues.Nil));
            // cell d: double top border
            Assert.That(cells[3].TopBorder.Val.Value, Is.EqualTo(BorderValues.Double));
            Assert.That(cells[3].TopBorder.Size.Value, Is.EqualTo(18u));
        }

        [Test]
        public void ColumnWidthsFromColElements()
        {
            var table = RenderTable("<table><colgroup><col width=\"100\"><col style=\"width: 200px\"></colgroup><tr><td>a</td><td>b</td></tr></table>");

            var grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => x.Width.Value).ToList();
            Assert.That(grid, Is.EqualTo((string[])["1500", "3000"]));
            var tableWidth = table.GetFirstChild<TableProperties>().TableWidth;
            Assert.That(tableWidth.Type.Value, Is.EqualTo(TableWidthUnitValues.Dxa));
            Assert.That(tableWidth.Width.Value, Is.EqualTo("4500"), "the columns define the table width");
            var cellWidths = table.Descendants<TableCellWidth>().Select(x => x.Width.Value).ToList();
            Assert.That(cellWidths, Is.EqualTo((string[])["1500", "3000"]));
        }

        [Test]
        public void ColumnWidthsFromCellsShareTheRemainingSpace()
        {
            // 600px table = 9000 twips: 25% = 2250, 150px = 2250, the third column gets the rest
            var table = RenderTable("<table width=\"600\"><tr><td width=\"25%\">a</td><td style=\"width:150px\">b</td><td>c</td></tr><tr><td colspan=\"2\">wide</td><td>d</td></tr></table>");

            var grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => x.Width.Value).ToList();
            Assert.That(grid, Is.EqualTo((string[])["2250", "2250", "4500"]));
            var spanning = table.Elements<TableRow>().Last().Elements<TableCell>().First();
            Assert.That(spanning.TableCellProperties.TableCellWidth.Width.Value, Is.EqualTo("4500"));
        }

        [Test]
        public void PaddingSpacingHeightAndLayout()
        {
            var html = """
                       <table cellpadding="5" cellspacing="2" bgcolor="#eeeeee" style="table-layout: fixed; margin-left: 20px">
                         <tr height="40" style="break-inside: avoid" valign="bottom" align="right">
                           <td style="padding: 2px 4px" nowrap>a</td>
                           <td style="height: 60px" align="center">b</td>
                         </tr>
                       </table>
                       """;
            var table = RenderTable(html);
            var properties = table.GetFirstChild<TableProperties>();

            Assert.That(properties.TableCellMarginDefault.TopMargin.Width.Value, Is.EqualTo("75"), "cellpadding 5px");
            Assert.That(properties.TableCellMarginDefault.TableCellLeftMargin.Width.Value, Is.EqualTo((short)75));
            Assert.That(properties.TableCellSpacing.Width.Value, Is.EqualTo("30"), "cellspacing 2px");
            Assert.That(properties.Shading.Fill.Value, Is.EqualTo("EEEEEE"));
            Assert.That(properties.TableLayout.Type.Value, Is.EqualTo(TableLayoutValues.Fixed));
            Assert.That(properties.TableIndentation.Width.Value, Is.EqualTo(300));

            var row = table.Elements<TableRow>().Single();
            Assert.That(row.TableRowProperties.GetFirstChild<TableRowHeight>().Val.Value, Is.EqualTo(900u), "the highest of row (40px) and cell (60px)");
            Assert.That(row.TableRowProperties.GetFirstChild<CantSplit>(), Is.Not.Null);

            var cells = row.Elements<TableCell>().ToList();
            var margin = cells[0].TableCellProperties.TableCellMargin;
            Assert.That(margin.TopMargin.Width.Value, Is.EqualTo("30"));
            Assert.That(margin.LeftMargin.Width.Value, Is.EqualTo("60"));
            Assert.That(cells[0].TableCellProperties.NoWrap, Is.Not.Null);
            Assert.That(cells[0].TableCellProperties.TableCellVerticalAlignment.Val.Value, Is.EqualTo(TableVerticalAlignmentValues.Bottom), "valign from the row");
            Assert.That(cells[0].Descendants<Justification>().Single().Val.Value, Is.EqualTo(JustificationValues.Right), "align from the row");
            Assert.That(cells[1].Descendants<Justification>().Single().Val.Value, Is.EqualTo(JustificationValues.Center), "the cell overrides the row");
            Assert.That(cells[0].Descendants<Indentation>(), Is.Empty, "cell padding is not a paragraph indent");
        }

        [Test]
        public void RowspanCellBordersOnlyOnTheOuterEdges()
        {
            var table = RenderTable("<table border=\"1\"><tr><td rowspan=\"3\">merged</td><td>a</td></tr><tr><td>b</td></tr><tr><td>c</td></tr></table>");

            var firstColumn = table.Elements<TableRow>().Select(r => r.Elements<TableCell>().First().TableCellProperties.TableCellBorders).ToList();
            Assert.That(Val(firstColumn[0].TopBorder), Is.EqualTo(BorderValues.Single));
            Assert.That(Val(firstColumn[0].BottomBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(Val(firstColumn[1].TopBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(Val(firstColumn[2].BottomBorder), Is.EqualTo(BorderValues.Single));
            Assert.That(firstColumn.All(x => Val(x.LeftBorder) == BorderValues.Single));
        }

        [Test]
        public void TableWithoutWidthFitsItsContent()
        {
            // a bullet-like layout table: a dash, an empty spacer column and the text right next to it
            var html = """
                       <table border="0" cellpadding="0" cellspacing="0">
                         <tr><td style="font-size: 10pt"><b>-</b></td><td style="font-size: 10pt"></td><td style="font-size: 10pt"><b>Con coeficientes de combinaci&#243;n</b></td></tr>
                       </table>
                       """;
            var table = RenderTable(html);

            var tableWidth = table.GetFirstChild<TableProperties>().TableWidth;
            Assert.That(tableWidth.Type.Value, Is.EqualTo(TableWidthUnitValues.Auto), "Word sizes the table to its content");
            Assert.That(table.Descendants<TableCellWidth>().Select(x => x.Type.Value), Is.All.EqualTo(TableWidthUnitValues.Auto));

            var grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => int.Parse(x.Width.Value)).ToList();
            Assert.That(grid[0], Is.LessThan(200), "the dash column is only as wide as the dash");
            Assert.That(grid[1], Is.LessThan(50), "the empty column almost disappears");
            Assert.That(grid[2], Is.InRange(2500, 4500), "the text column fits its text on one line");
        }

        [Test]
        public void TableWithoutWidthWrapsWhenContentIsWiderThanThePage()
        {
            var longText = string.Join(" ", Enumerable.Repeat("long words", 60));
            var table = RenderTable($"<table><tr><td>short</td><td>{longText}</td></tr></table>");

            Assert.That(table.GetFirstChild<TableProperties>().TableWidth.Type.Value, Is.EqualTo(TableWidthUnitValues.Pct), "full width when the content does not fit");
            var grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => int.Parse(x.Width.Value)).ToList();
            Assert.That(grid.Sum(), Is.InRange(9300, 9400), "the columns fill the text width");
            Assert.That(grid[0], Is.LessThan(grid[1]), "the short column keeps its content width, the long one wraps");
        }

        [Test]
        public void TableWithoutWidthFullWidthOption()
        {
            var body = Render("{{ds}:html}", "<table><tr><td>-</td><td>text</td></tr></table>", styles: TableGridStyle(),
                configuration: new HtmlFormatterConfiguration { TablesWithoutWidth = HtmlTableWidth.FullWidth });
            var table = body.Elements<Table>().Single();

            Assert.That(table.GetFirstChild<TableProperties>().TableWidth.Width.Value, Is.EqualTo("5000"));
            var grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => x.Width.Value).Distinct();
            Assert.That(grid.Count(), Is.EqualTo(1), "equal columns");
        }

        [Test]
        public void TableWithoutStyleAndWithoutBorderInformationGetsAGrid()
        {
            var table = RenderTable("<table><tr><td>a</td></tr></table>", styles: new Styles());

            var borders = table.GetFirstChild<TableProperties>().TableBorders;
            Assert.That(Val(borders.InsideHorizontalBorder), Is.EqualTo(BorderValues.Single));
            Assert.That(table.GetFirstChild<TableProperties>().TableStyle, Is.Null);
        }
    }
}

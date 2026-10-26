using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Html.Test
{
    internal partial class HtmlFormatterTest
    {
        private static Styles TableGridStyle()
        {
            return new Styles(new Style(new StyleName { Val = "Table Grid" }) { Type = StyleValues.Table, StyleId = "TableGrid" });
        }

        private static Table RenderTable(string html, Styles styles = null, HtmlFormatterConfiguration configuration = null)
        {
            return Render("{{ds}:html}", html, styles: styles ?? TableGridStyle(), configuration: configuration).Elements<Table>().Single();
        }

        private static HtmlFormatterConfiguration AutoFit => new() { TableLayout = HtmlTableLayout.AutoFit };

        private static List<int> GridWidths(Table table)
        {
            return table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => int.Parse(x.Width.Value)).ToList();
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
            Assert.That(borders.ChildElements, Has.Count.EqualTo(6));
            Assert.That(borders.ChildElements.Cast<BorderType>().Select(Val), Is.All.EqualTo(BorderValues.Nil));
            Assert.That(table.Descendants<TableCellBorders>(), Is.Empty, "the table borders apply to every cell");
        }

        [Test]
        public void BorderAttributeDrawsFrameAndCellLines()
        {
            var table = RenderTable("<table border=\"2\"><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>");

            var borders = table.GetFirstChild<TableProperties>().TableBorders;
            Assert.That(borders.TopBorder.Size.Value, Is.EqualTo(12u), "2px frame = 12 eighths of a point");
            Assert.That(borders.InsideHorizontalBorder.Size.Value, Is.EqualTo(6u), "1px lines between cells");
            Assert.That(borders.InsideVerticalBorder.Val.Value, Is.EqualTo(BorderValues.Single));

            Assert.That(table.Descendants<TableCellBorders>(), Is.Empty, "same lines everywhere - written once on the table");
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

            var tableBorders = table.GetFirstChild<TableProperties>().TableBorders;
            // cell a: own border, left side removed
            Assert.That(cells[0].TopBorder.Color.Value, Is.EqualTo("FF0000"));
            Assert.That(Val(cells[0].LeftBorder), Is.EqualTo(BorderValues.Nil));
            // cell b: top from the table frame (not repeated on the cell), bottom from the row
            Assert.That(cells[1].TopBorder, Is.Null);
            Assert.That(cells[1].BottomBorder.Color.Value, Is.EqualTo("CCCCCC"));
            // between cells without own borders: no rules -> nothing (like a browser)
            Assert.That(Val(tableBorders.InsideVerticalBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(cells[2], Is.Null, "cell c has only the table borders");
            // cell d: double top border, the rest from the table
            Assert.That(cells[3].TopBorder.Val.Value, Is.EqualTo(BorderValues.Double));
            Assert.That(cells[3].TopBorder.Size.Value, Is.EqualTo(18u));
            Assert.That(cells[3].ChildElements, Has.Count.EqualTo(1));
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
            Assert.That(cells[0].TableCellProperties.NoWrap, Is.Null, "Word's noWrap would cut the text off instead of wrapping");
            Assert.That(cells[0].TableCellProperties.TableCellVerticalAlignment.Val.Value, Is.EqualTo(TableVerticalAlignmentValues.Bottom), "valign from the row");
            Assert.That(cells[0].Descendants<Justification>().Single().Val.Value, Is.EqualTo(JustificationValues.Right), "align from the row");
            Assert.That(cells[1].Descendants<Justification>().Single().Val.Value, Is.EqualTo(JustificationValues.Center), "the cell overrides the row");
            Assert.That(cells[0].Descendants<Indentation>(), Is.Empty, "cell padding is not a paragraph indent");
        }

        [Test]
        public void RowspanCellBordersOnlyOnTheOuterEdges()
        {
            var table = RenderTable("<table border=\"1\"><tr><td rowspan=\"3\">merged</td><td>a</td></tr><tr><td>b</td></tr><tr><td>c</td></tr></table>");

            // the outer edges come from the table borders; only the inner edges of the merged cell are removed per cell
            var firstColumn = table.Elements<TableRow>().Select(r => r.Elements<TableCell>().First().TableCellProperties.TableCellBorders).ToList();
            Assert.That(firstColumn[0].TopBorder, Is.Null);
            Assert.That(Val(firstColumn[0].BottomBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(Val(firstColumn[1].TopBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(Val(firstColumn[1].BottomBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(Val(firstColumn[2].TopBorder), Is.EqualTo(BorderValues.Nil));
            Assert.That(firstColumn[2].BottomBorder, Is.Null);
            Assert.That(firstColumn.All(x => x.LeftBorder == null));
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
            var table = RenderTable(html, configuration: AutoFit);

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
            var table = RenderTable($"<table><tr><td>short</td><td>{longText}</td></tr></table>", configuration: AutoFit);

            Assert.That(table.GetFirstChild<TableProperties>().TableWidth.Type.Value, Is.EqualTo(TableWidthUnitValues.Pct), "full width when the content does not fit");
            var grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => int.Parse(x.Width.Value)).ToList();
            Assert.That(grid.Sum(), Is.InRange(9300, 9400), "the columns fill the text width");
            Assert.That(grid[0], Is.LessThan(grid[1]), "the short column keeps its content width, the long one wraps");
        }

        [Test]
        public void AdjacentTablesAreKeptApart()
        {
            // "Donde:" table directly followed by the definitions table - Word would merge them into one table
            var html = """
                       <table border="0" cellpadding="0" cellspacing="0"><tr><td>-</td><td></td><td>Donde:</td></tr></table>
                       <table border="0" cellspacing="0">
                         <tr><td>G<sub>k</sub></td><td style="white-space: nowrap">Acci&#243;n permanente</td></tr>
                         <tr><td>&#947;<sub>Q,i</sub></td><td style="white-space: nowrap">Coeficiente parcial de seguridad de las acciones variables de acompa&#241;amiento</td></tr>
                       </table>
                       """;
            var body = Render("{{ds}:html}", html, styles: TableGridStyle());

            var elements = body.ChildElements.Where(x => x is not SectionProperties).ToList();
            Assert.That(elements.Select(x => x.GetType().Name), Is.EqualTo((string[])["Table", "Paragraph", "Table"]));
            var separator = (Paragraph)elements[1];
            Assert.That(separator.InnerText, Is.Empty);
            Assert.That(separator.ParagraphProperties.SpacingBetweenLines.Line.Value, Is.EqualTo("20"), "a 1pt line, nearly invisible");

            var definitions = (Table)elements[2];
            Assert.That(definitions.Descendants<NoWrap>(), Is.Empty);
            var grid = definitions.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(x => int.Parse(x.Width.Value)).ToList();
            Assert.That(grid[1], Is.GreaterThan(grid[0] * 5), "the nowrap definitions get the space");
        }

        [Test]
        public void HtmlTableNextToTemplateTableIsKeptApart()
        {
            var templateTable = new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "5000" }),
                new TableRow(new TableCell(new Paragraph(new Run(new Text("template table"))))));
            var body = RenderDocument(new Body(new Paragraph(new Run(new Text("{{ds}:html}"))), templateTable),
                "<table><tr><td>html table</td></tr></table>", styles: TableGridStyle()).Body;

            Assert.That(body.ChildElements.Select(x => x.GetType().Name), Is.EqualTo((string[])["Table", "Paragraph", "Table"]));
        }

        [Test]
        public void TableWithoutWidthFullWidthOption()
        {
            var body = Render("{{ds}:html}", "<table><tr><td>-</td><td>text</td></tr></table>", styles: TableGridStyle(),
                configuration: new HtmlFormatterConfiguration { TablesWithoutWidth = HtmlTableWidth.FullWidth, TableLayout = HtmlTableLayout.AutoFit });
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

        [Test]
        public void FixedLayoutWritesFinalColumnWidths()
        {
            var table = RenderTable("<table><tr><td>a</td><td>some longer text</td></tr></table>");

            var properties = table.GetFirstChild<TableProperties>();
            Assert.That(properties.TableLayout.Type.Value, Is.EqualTo(TableLayoutValues.Fixed));
            Assert.That(properties.TableWidth.Type.Value, Is.EqualTo(TableWidthUnitValues.Dxa));
            var grid = GridWidths(table);
            Assert.That(properties.TableWidth.Width.Value, Is.EqualTo(grid.Sum().ToString()), "the table is as wide as its columns");
            Assert.That(table.Descendants<TableCellWidth>().Select(x => int.Parse(x.Width.Value)), Is.EqualTo(grid));
            Assert.That(table.Descendants<TableCellWidth>().Select(x => x.Type.Value), Is.All.EqualTo(TableWidthUnitValues.Dxa));
        }

        [Test]
        public void NumbersDoNotWrapInNarrowColumns()
        {
            // the html asks for 20pt columns in a 225pt table, "1.350" at 10pt needs about 36pt with the cell margins
            var columns = string.Concat(Enumerable.Repeat("<col width=\"27\">", 8));
            var cells = string.Concat(Enumerable.Repeat("<td>1.350</td>", 8));
            var table = RenderTable($"<table width=\"300\" style=\"font-size: 10pt\"><colgroup>{columns}</colgroup><tr>{cells}</tr></table>");

            var grid = GridWidths(table);
            Assert.That(grid, Is.All.InRange(700, 800), "each column holds the number on one line");
            Assert.That(grid.Sum(), Is.GreaterThan(4500), "the table grows beyond its html width, like in a browser");
        }

        [Test]
        public void TooWideTableTakesTheSpaceFromTheTextColumns()
        {
            var text = string.Join(" ", Enumerable.Repeat("description", 40));
            var numbers = string.Concat(Enumerable.Repeat("<td>12.345</td>", 6));
            var table = RenderTable($"<table><tr><td>{text}</td>{numbers}</tr></table>");

            var grid = GridWidths(table);
            Assert.That(grid.Sum(), Is.InRange(9300, 9400), "the table fits the text width");
            Assert.That(grid.Skip(1).Distinct().Count(), Is.EqualTo(1));
            Assert.That(grid[1], Is.InRange(850, 1000), "the numbers keep their width (11pt)");
            Assert.That(grid[0], Is.GreaterThan(grid[1]), "the text column wraps");
        }

        [Test]
        public void NumericColumnsShrinkLastWhenNothingFits()
        {
            var words = string.Concat(Enumerable.Repeat("<td>Pneumonoultramicroscopicsilicovolcanoconiosis</td>", 3));
            var table = RenderTable($"<table><tr>{words}<td>1.350</td></tr></table>");

            var grid = GridWidths(table);
            Assert.That(grid.Sum(), Is.InRange(9300, 9400));
            Assert.That(grid[3], Is.InRange(700, 900), "the narrow column keeps its width, the long words are cut");
        }

        [Test]
        public void HeaderRowsWithoutTheadRepeatOnEveryPage()
        {
            var table = RenderTable("<table><tr><th>a</th><th>b</th></tr><tr><th>row</th><td>1</td></tr></table>");

            var rows = table.Elements<TableRow>().ToList();
            Assert.That(rows[0].TableRowProperties.GetFirstChild<TableHeader>(), Is.Not.Null);
            Assert.That(rows[1].TableRowProperties?.GetFirstChild<TableHeader>(), Is.Null, "a row with data is no header");
        }

        [Test]
        public void NestedTableFitsItsCell()
        {
            var inner = string.Join(" ", Enumerable.Repeat("nested", 50));
            var table = RenderTable($"<table width=\"300\"><tr><td><table><tr><td>{inner}</td></tr></table></td></tr></table>");

            var nested = table.Descendants<Table>().Single();
            var outerCell = int.Parse(table.Elements<TableRow>().Single().Elements<TableCell>().Single().TableCellProperties.TableCellWidth.Width.Value);
            Assert.That(GridWidths(nested).Sum(), Is.LessThanOrEqualTo(outerCell - (2 * 108)));
        }

        [Test]
        public void EmptyCellsGetTheFormattingOfTheirText()
        {
            var template = new Paragraph(new Run(new RunProperties(new RunFonts { Ascii = "Heebo", HighAnsi = "Heebo" }, new FontSize { Val = "20" }), new Text("{{ds}:html}")));
            var (body, document) = RenderDocument(new Body(template), "<table><tr><td>1</td><td></td></tr></table>", styles: TableGridStyle());

            var paragraphs = body.Descendants<TableCell>().Select(c => c.Elements<Paragraph>().Single()).ToList();
            var styleIds = paragraphs.Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value).ToList();
            Assert.That(styleIds, Is.All.Not.Null);
            Assert.That(styleIds.Distinct().Count(), Is.EqualTo(1), "the empty cell's paragraph mark is formatted like the text");
            var style = document.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().Single(x => x.StyleId == styleIds[0]);
            Assert.That(style.StyleRunProperties.RunFonts.Ascii.Value, Is.EqualTo("Heebo"));
            Assert.That(style.StyleRunProperties.FontSize.Val.Value, Is.EqualTo("20"));
            Assert.That(paragraphs[0].Elements<Run>().Single().RunProperties, Is.Null);
        }
    }
}

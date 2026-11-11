using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// Creates the drawing elements for page pictures (inline in the text flow or anchored to the page) and the
    /// placeholder overlay shape.
    /// </summary>
    internal static class PdfDrawings
    {
        private const string PictureUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private const string ShapeUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        private const uint RelativeHeightBase = 251659264;
        private const long OverlayInsetEmu = 108000; // 3 mm
        private const double OverlayTitleFontPt = 16;
        private const double OverlayLineFontPt = 12;

        public static Drawing CreateInline(string relationshipId, uint id, long cx, long cy, string altText, bool border)
        {
            var docProperties = CreateDocProperties(id, $"PDF page {id}", altText);
            return new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    docProperties,
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(new A.GraphicData(CreatePicture(relationshipId, id, cx, cy, border)) { Uri = PictureUri }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,
                });
        }

        /// <summary>
        /// A picture anchored to the page at the given offset, in front of the text, taking no space in the text flow.
        /// </summary>
        public static Drawing CreatePageAnchor(string relationshipId, uint id, long offsetX, long offsetY, long cx, long cy, string altText, bool border)
        {
            return new Drawing(
                new DW.Anchor(
                    new DW.SimplePosition { X = 0L, Y = 0L },
                    new DW.HorizontalPosition(new DW.PositionOffset(offsetX.ToString(System.Globalization.CultureInfo.InvariantCulture)))
                    {
                        RelativeFrom = DW.HorizontalRelativePositionValues.Page,
                    },
                    new DW.VerticalPosition(new DW.PositionOffset(offsetY.ToString(System.Globalization.CultureInfo.InvariantCulture)))
                    {
                        RelativeFrom = DW.VerticalRelativePositionValues.Page,
                    },
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.WrapNone(),
                    CreateDocProperties(id, $"PDF page {id}", altText),
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(new A.GraphicData(CreatePicture(relationshipId, id, cx, cy, border)) { Uri = PictureUri }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,
                    SimplePos = false,
                    RelativeHeight = RelativeHeightBase + (id * 2),
                    BehindDoc = false,
                    Locked = false,
                    LayoutInCell = true,
                    AllowOverlap = true,
                });
        }

        /// <summary>
        /// A text box centered on the page, drawn on top of the page picture: white, slightly transparent, red border.
        /// The first line is the title.
        /// </summary>
        public static Drawing CreateOverlay(uint id, IReadOnlyList<string> lines, long pageWidthEmu, string description)
        {
            var (cx, cy) = EstimateOverlaySize(lines, pageWidthEmu);
            var content = new TextBoxContent();
            for (var i = 0; i < lines.Count; i++)
            {
                var isTitle = i == 0;
                var fontSize = ((int)((isTitle ? OverlayTitleFontPt : OverlayLineFontPt) * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture);
                var runProperties = new RunProperties(
                    new RunFonts { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" },
                    new Color { Val = isTitle ? "C00000" : "000000" },
                    new FontSize { Val = fontSize },
                    new FontSizeComplexScript { Val = fontSize });
                if (isTitle)
                {
                    runProperties.Bold = new Bold();
                }
                content.AppendChild(new Paragraph(
                    new ParagraphProperties
                    {
                        SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = isTitle ? "120" : "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },
                        Justification = new Justification { Val = JustificationValues.Center },
                    },
                    new Run(runProperties, new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve })));
            }

            var shape = new Wps.WordprocessingShape(
                new Wps.NonVisualDrawingShapeProperties { TextBox = true },
                new Wps.ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                    new A.SolidFill(new A.RgbColorModelHex(new A.Alpha { Val = 80000 }) { Val = "FFFFFF" }),
                    new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "C00000" })) { Width = 19050 }),
                new Wps.TextBoxInfo2(content),
                new Wps.TextBodyProperties(new A.ShapeAutoFit())
                {
                    Rotation = 0,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    LeftInset = (int)OverlayInsetEmu,
                    TopInset = (int)OverlayInsetEmu,
                    RightInset = (int)OverlayInsetEmu,
                    BottomInset = (int)OverlayInsetEmu,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = false,
                });

            return new Drawing(
                new DW.Anchor(
                    new DW.SimplePosition { X = 0L, Y = 0L },
                    new DW.HorizontalPosition(new DW.HorizontalAlignment("center")) { RelativeFrom = DW.HorizontalRelativePositionValues.Page },
                    new DW.VerticalPosition(new DW.VerticalAlignment("center")) { RelativeFrom = DW.VerticalRelativePositionValues.Page },
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.WrapNone(),
                    CreateDocProperties(id, $"PDF placeholder {id}", description),
                    new DW.NonVisualGraphicFrameDrawingProperties(),
                    new A.Graphic(new A.GraphicData(shape) { Uri = ShapeUri }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,
                    SimplePos = false,
                    RelativeHeight = RelativeHeightBase + (id * 2) + 1,
                    BehindDoc = false,
                    Locked = false,
                    LayoutInCell = true,
                    AllowOverlap = true,
                });
        }

        /// <summary>
        /// Size of the overlay text box from the text (about 0.55 em per character). Word grows the height to the
        /// text anyway (<c>spAutoFit</c>), the estimate keeps the box readable in other applications.
        /// </summary>
        private static (long Cx, long Cy) EstimateOverlaySize(IReadOnlyList<string> lines, long pageWidthEmu)
        {
            double width = 0;
            double height = 0;
            for (var i = 0; i < lines.Count; i++)
            {
                var fontPt = i == 0 ? OverlayTitleFontPt : OverlayLineFontPt;
                width = Math.Max(width, lines[i].Length * fontPt * 0.55);
                height += (fontPt * 1.25) + (i == 0 ? 6 : 0);
            }
            var cx = OpenXmlHelper.PointsToEmu(width) + (2 * OverlayInsetEmu);
            var maxWidth = (long)(pageWidthEmu * 0.9);
            if (cx > maxWidth)
            {
                // the text wraps - make room for the extra lines
                height *= Math.Ceiling((double)cx / maxWidth);
                cx = maxWidth;
            }
            var cy = OpenXmlHelper.PointsToEmu(height) + (2 * OverlayInsetEmu);
            return (Math.Max(cx, 360000), Math.Max(cy, 360000));
        }

        private static DW.DocProperties CreateDocProperties(uint id, string name, string description)
        {
            var docProperties = new DW.DocProperties { Id = id, Name = name };
            if (!string.IsNullOrEmpty(description))
            {
                docProperties.Description = description;
            }
            return docProperties;
        }

        private static PIC.Picture CreatePicture(string relationshipId, uint id, long cx, long cy, bool border)
        {
            var shapeProperties = new PIC.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });
            if (border)
            {
                shapeProperties.AppendChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "808080" })) { Width = 6350 });
            }
            return new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = $"PdfPage{id}.png" },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                    new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                    new A.Stretch(new A.FillRectangle())),
                shapeProperties);
        }

        /// <summary>
        /// Paragraph properties for a picture paragraph in the text flow: based on the template paragraph (alignment,
        /// style), without section break, list numbering and indentation, and with single line spacing - an exact line
        /// height would clip the inline picture to a strip.
        /// </summary>
        public static ParagraphProperties CreateFlowParagraphProperties(ParagraphProperties template, bool pageBreakBefore)
        {
            var properties = (ParagraphProperties)template?.CloneNode(true) ?? new ParagraphProperties();
            properties.SectionProperties = null;
            properties.NumberingProperties = null;
            properties.ParagraphPropertiesChange = null;
            properties.Indentation = new Indentation { Left = "0", Right = "0", FirstLine = "0" };
            properties.SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            properties.SnapToGrid = new SnapToGrid { Val = false };
            properties.KeepLines = new KeepLines();
            properties.PageBreakBefore = pageBreakBefore ? new PageBreakBefore() : null;
            return properties;
        }

        public static Paragraph CreateFlowParagraph(ParagraphProperties template, bool pageBreak, Drawing drawing)
        {
            var paragraph = new Paragraph(CreateFlowParagraphProperties(template, false));
            if (pageBreak)
            {
                // inside the picture paragraph, before the picture: a break paragraph of its own (or a break after the
                // picture) leaves an empty paragraph mark that Word moves to a blank page
                paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));
            }
            paragraph.AppendChild(new Run(drawing));
            return paragraph;
        }

        /// <summary>
        /// The tiny paragraph (1 pt) that holds anchored page pictures. Pages after the first of a section start with a page break.
        /// </summary>
        public static Paragraph CreateAnchorParagraph(bool pageBreakBefore, IEnumerable<Drawing> drawings)
        {
            var properties = new ParagraphProperties
            {
                KeepLines = new KeepLines(),
                PageBreakBefore = pageBreakBefore ? new PageBreakBefore() : null,
                SnapToGrid = new SnapToGrid { Val = false },
                SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0", Line = "20", LineRule = LineSpacingRuleValues.Exact },
                Indentation = new Indentation { Left = "0", Right = "0", FirstLine = "0" },
                ParagraphMarkRunProperties = new ParagraphMarkRunProperties(new FontSize { Val = "2" }, new FontSizeComplexScript { Val = "2" }),
            };
            var paragraph = new Paragraph(properties);
            foreach (var drawing in drawings)
            {
                paragraph.AppendChild(new Run(new RunProperties(new FontSize { Val = "2" }, new FontSizeComplexScript { Val = "2" }), drawing));
            }
            return paragraph;
        }

        /// <summary>
        /// An empty paragraph of 1 pt height, used to end a section where no paragraph is available.
        /// </summary>
        public static Paragraph CreateMinimalParagraph()
        {
            return CreateAnchorParagraph(false, Enumerable.Empty<Drawing>());
        }
    }
}

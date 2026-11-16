using System;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// How the pages of a PDF are sized and placed.
    /// </summary>
    public enum PdfFit
    {
        /// <summary>Scale each page into the text area of the current section (width and height).</summary>
        Page,

        /// <summary>Scale each page to the width of the text area. Tall pages may be cut off - only for short pages.</summary>
        Width,

        /// <summary>Real PDF size, only scaled down if the page is larger than the text area.</summary>
        None,

        /// <summary>
        /// True size: every PDF page gets a page of exactly its own size in a section of its own, without margins and
        /// (by default) without header and footer. For plans and drawings that must print to scale.
        /// </summary>
        Section,
    }

    /// <summary>
    /// Headers and footers of the sections created for <see cref="PdfFit.Section"/>.
    /// </summary>
    public enum PdfHeaderFooter
    {
        /// <summary>Empty header and footer, nothing is drawn over the PDF page.</summary>
        None,

        /// <summary>The headers and footers of the document are kept (drawn on top of the PDF page).</summary>
        Inherit,
    }

    /// <summary>
    /// What happens with <see cref="PdfFit.Section"/> pages that are larger than Word allows (22 inch, 558.8 mm).
    /// </summary>
    public enum PdfOversizePolicy
    {
        /// <summary>
        /// The page is scaled down until it fits and gets an overlay that names the original size, so it can be
        /// replaced by the original page after the export to PDF.
        /// </summary>
        ScaleWithOverlay,

        /// <summary>An <see cref="OpenXmlTemplateException"/> is thrown.</summary>
        Error,
    }

    /// <summary>
    /// When <see cref="PdfFit.Section"/> pages get the placeholder overlay.
    /// </summary>
    public enum PdfOverlayMode
    {
        /// <summary>Only pages that had to be scaled down (see <see cref="PdfOversizePolicy.ScaleWithOverlay"/>).</summary>
        Auto,

        /// <summary>Every page in a section of its own, also those at true size.</summary>
        Always,

        /// <summary>Never.</summary>
        Never,
    }

    /// <summary>
    /// How <see cref="PdfFit.Section"/> treats PDF pages that have the same size as the pages of the current section.
    /// </summary>
    public enum PdfSamePageSize
    {
        /// <summary>Inserted into the current section like <see cref="PdfFit.Page"/>, with the margins of the document.</summary>
        Inline,

        /// <summary>Also get a section of their own without margins.</summary>
        Section,
    }

    /// <summary>
    /// What is drawn on the pages created for a PDF.
    /// </summary>
    public enum PdfContent
    {
        /// <summary>The rendered PDF page.</summary>
        Render,

        /// <summary>
        /// Nothing: an empty page of the PDF page's size in a section of its own, to be replaced by the original page
        /// after the export to PDF. The PDF is only read for its page count and page sizes, nothing is rendered.
        /// </summary>
        Blank,
    }

    public class PdfFormatterConfiguration
    {
        /// <summary>
        /// Largest page width or height Word supports: 22 inch in twips.
        /// </summary>
        public const int WordMaxPageSizeTwips = 31680;

        public static readonly PdfFormatterConfiguration Default = new();

        /// <summary>Resolution of the rendered pages (72-600). Affects the file size, not the displayed size. Template argument <c>dpi</c>.</summary>
        public int Dpi { get; set; } = 150;

        /// <summary>Image format of the rendered pages. Template argument <c>fmt</c> (<c>png</c>, <c>jpg</c>).</summary>
        public PdfImageFormat ImageFormat { get; set; } = PdfImageFormat.Png;

        /// <summary>JPEG quality (1-100). Template argument <c>q</c>.</summary>
        public int JpegQuality { get; set; } = 85;

        /// <summary>
        /// Selection of 1-based pages, e.g. <c>1-3,5,8-</c>; null or empty for all pages. Template argument <c>pages</c>.
        /// </summary>
        public string Pages { get; set; }

        /// <summary>How pages are sized and placed. Template argument <c>fit</c>.</summary>
        public PdfFit Fit { get; set; } = PdfFit.Page;

        /// <summary>
        /// What the pages show. Template argument <c>content</c> (<c>render</c>, <c>blank</c>). <see cref="PdfContent.Blank"/>
        /// always lays out the pages like <see cref="PdfFit.Section"/> (whatever <see cref="Fit"/> says), gives every page
        /// a section of its own (<see cref="SamePageSize"/> is ignored) and ignores the render settings.
        /// </summary>
        public PdfContent Content { get; set; } = PdfContent.Render;

        /// <summary>
        /// Machine-readable text written on every page in a section of its own (<see cref="PdfFit.Section"/>,
        /// <see cref="PdfContent.Blank"/>), invisible (white, 1 pt), so a tool can find the page in the exported PDF and
        /// replace it with the original page. Null or empty: no marker. Tokens as in the overlay lines plus <c>{key}</c>.
        /// Line breaks are replaced with spaces. Template argument <c>marker</c>: the template syntax does not allow braces
        /// in arguments, so tokens are written <c>%key%</c> there, e.g. <c>marker:'[[PDF|%key%|%page%|%pages%]]'</c>.
        /// </summary>
        public string MarkerText { get; set; }

        /// <summary>
        /// Explicit maximum width of a page picture in EMU (360 000 per cm); overrides <see cref="Fit"/> (except
        /// <see cref="PdfFit.Section"/>). Template argument <c>w</c>, e.g. <c>w:16cm</c>.
        /// </summary>
        public long? MaxWidthEmu { get; set; }

        /// <summary>Explicit maximum height of a page picture in EMU. Template argument <c>h</c>.</summary>
        public long? MaxHeightEmu { get; set; }

        /// <summary>
        /// Page break before the first page (pages after the first always start on a new page). Template argument <c>pb</c>.
        /// </summary>
        public bool PageBreakBefore { get; set; }

        /// <summary>Thin border around each page picture. Template argument <c>border</c>.</summary>
        public bool Border { get; set; }

        /// <summary>
        /// Part of the text height that is kept free below a page scaled with <see cref="PdfFit.Page"/>, so the
        /// paragraph mark of the picture paragraph still fits on the page (otherwise Word adds blank pages).
        /// </summary>
        public double HeightSafetyMargin { get; set; } = 0.03;

        /// <summary>
        /// Name used in the alt text and the overlay if neither the template (<c>label</c>) nor the value
        /// (<see cref="PdfFile.Name"/>) gives one.
        /// </summary>
        public string DefaultLabel { get; set; } = "Document";

        /// <summary>Name of the document, overrides <see cref="PdfFile.Name"/>. Template argument <c>label</c>.</summary>
        public string Label { get; set; }

        /// <summary>
        /// Stable id of the document for <c>{key}</c>, overrides <see cref="PdfFile.Key"/>. Template argument <c>key</c>.
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Password for encrypted PDFs. Template argument <c>pwd</c> (prefer the configuration - templates are no secret store).
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Returns the password for an encrypted PDF; called with the name of the document if <see cref="Password"/> is not set.
        /// </summary>
        public Func<string, string> PasswordProvider { get; set; }

        /// <summary>Maximum number of pages per placeholder; more selected pages throw an exception.</summary>
        public int MaxPages { get; set; } = 200;

        /// <summary>Maximum number of pixels of a rendered page (protects against huge page sizes).</summary>
        public long MaxPixelsPerPage { get; set; } = 100_000_000;

        /// <summary>
        /// If a page would exceed <see cref="MaxPixelsPerPage"/>, render it with a lower dpi (true, default) or throw (false).
        /// </summary>
        public bool ReduceDpiForLargePages { get; set; } = true;

        /// <summary>Render annotations (stamps, comments, revision clouds).</summary>
        public bool WithAnnotations { get; set; } = true;

        /// <summary>Render the content of form fields.</summary>
        public bool WithFormFields { get; set; } = true;

        /// <summary>
        /// Page margins of the sections created for <see cref="PdfFit.Section"/> in twips (1/20 pt). Anything other
        /// than 0 scales the page into the remaining area and gives up the true scale. Template argument <c>margin</c>.
        /// </summary>
        public int SectionMarginTwips { get; set; }

        /// <summary>Headers and footers of the sections created for <see cref="PdfFit.Section"/>. Template argument <c>hf</c>.</summary>
        public PdfHeaderFooter HeaderFooter { get; set; } = PdfHeaderFooter.None;

        /// <summary>
        /// Pages larger than Word allows with <see cref="PdfFit.Section"/>. Template argument <c>oversize</c> (<c>scale</c>, <c>error</c>).
        /// </summary>
        public PdfOversizePolicy Oversize { get; set; } = PdfOversizePolicy.ScaleWithOverlay;

        /// <summary>Which <see cref="PdfFit.Section"/> pages get the placeholder overlay. Template argument <c>overlay</c>.</summary>
        public PdfOverlayMode Overlay { get; set; } = PdfOverlayMode.Auto;

        /// <summary>PDF pages with the page size of the current section and <see cref="PdfFit.Section"/>. Template argument <c>same</c>.</summary>
        public PdfSamePageSize SamePageSize { get; set; } = PdfSamePageSize.Inline;

        /// <summary>Largest page width or height of a section in twips, see <see cref="WordMaxPageSizeTwips"/>.</summary>
        public int MaxPageSizeTwips { get; set; } = WordMaxPageSizeTwips;

        /// <summary>
        /// Adds <c>w:doNotAutoCompressPictures</c> to the document settings, so Word does not downsample the page
        /// pictures when the document is saved (only matters above 220 dpi).
        /// </summary>
        public bool DisableImageCompression { get; set; }

        /// <summary>First line of the overlay. An empty text drops the line.</summary>
        public string OverlayTitle { get; set; } = "PLACEHOLDER - replace with original";

        /// <summary>
        /// Size line of the overlay. Tokens: <c>{sheet}</c>, <c>{orientation}</c>, <c>{width}</c>, <c>{height}</c>
        /// (mm), <c>{scale}</c>, <c>{name}</c>, <c>{key}</c>, <c>{page}</c>, <c>{pages}</c>. An empty text drops the line.
        /// </summary>
        public string OverlaySizeLine { get; set; } = "Target size: {sheet} {orientation}, {width} × {height} mm (shown at {scale})";

        /// <summary>Size line of the overlay for page sizes that are not a known sheet size.</summary>
        public string OverlaySizeLineUnknownSheet { get; set; } = "Target size: {width} × {height} mm (shown at {scale})";

        /// <summary>Line of the overlay naming the source.</summary>
        public string OverlaySourceLine { get; set; } = "Plan: {name}, page {page} of {pages}";

        /// <summary>Text for <c>{orientation}</c> of landscape pages.</summary>
        public string OrientationLandscape { get; set; } = "landscape";

        /// <summary>Text for <c>{orientation}</c> of portrait pages.</summary>
        public string OrientationPortrait { get; set; } = "portrait";

        /// <summary>Alt text of each page picture. Same tokens as the overlay lines.</summary>
        public string AltText { get; set; } = "{name}, page {page} of {pages}";

        public PdfFormatterConfiguration Clone()
        {
            return (PdfFormatterConfiguration)MemberwiseClone();
        }
    }
}

using System.Collections.Generic;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// Image format of rendered PDF pages.
    /// </summary>
    public enum PdfImageFormat
    {
        /// <summary>Lossless, best for text, line art and plans.</summary>
        Png,

        /// <summary>Lossy and much smaller for scans and photos.</summary>
        Jpeg,
    }

    /// <summary>
    /// Size of a PDF page in points (1/72 inch) as a viewer shows it: crop box with the page rotation applied.
    /// </summary>
    public sealed record PdfPageSize(double WidthPoints, double HeightPoints);

    /// <summary>
    /// Pages to render and how.
    /// </summary>
    /// <param name="PageIndices">0-based indices of the pages to render, in this order.</param>
    /// <param name="Dpi">Resolution of the rendered pictures. Only affects the pixel size, not the displayed size.</param>
    /// <param name="Format">Image format of the rendered pictures.</param>
    /// <param name="JpegQuality">JPEG quality 1-100, only used for <see cref="PdfImageFormat.Jpeg"/>.</param>
    /// <param name="WithAnnotations">Render annotations (stamps, comments, revision clouds).</param>
    /// <param name="WithFormFields">Render the content of form fields.</param>
    /// <param name="Password">Password of an encrypted PDF, null if none.</param>
    public sealed record PdfRenderRequest(
        IReadOnlyList<int> PageIndices,
        int Dpi,
        PdfImageFormat Format,
        int JpegQuality,
        bool WithAnnotations,
        bool WithFormFields,
        string Password);

    /// <summary>
    /// A rendered PDF page.
    /// </summary>
    /// <param name="PageIndex">0-based page index.</param>
    /// <param name="WidthPoints">Page width in points (after rotation). Used for the displayed size - not the pixel size.</param>
    /// <param name="HeightPoints">Page height in points (after rotation).</param>
    /// <param name="ImageBytes">The encoded picture.</param>
    /// <param name="Format">Format of <paramref name="ImageBytes"/>.</param>
    public sealed record RenderedPdfPage(
        int PageIndex,
        double WidthPoints,
        double HeightPoints,
        byte[] ImageBytes,
        PdfImageFormat Format);

    /// <summary>
    /// Renders PDF pages to pictures. <c>DocxTemplater.Pdf.Pdfium</c> contains an implementation based on PDFium.
    /// </summary>
    public interface IPdfRenderer
    {
        /// <summary>
        /// Returns the sizes of all pages; the number of entries is the page count.
        /// </summary>
        IReadOnlyList<PdfPageSize> GetPageSizes(byte[] pdf, string password);

        /// <summary>
        /// Renders the requested pages in the requested order. Implementations should open the document once and
        /// render the pages lazily, so large documents do not need all pictures in memory at once.
        /// </summary>
        IEnumerable<RenderedPdfPage> RenderPages(byte[] pdf, PdfRenderRequest request);
    }
}

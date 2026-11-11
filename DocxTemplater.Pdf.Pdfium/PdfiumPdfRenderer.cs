using System;
using System.Collections.Generic;
using System.Linq;
using PDFtoImage;
using SkiaSharp;

// PDFtoImage declares its supported platforms (Windows, Linux, macOS, Android, iOS, browser) - the calls below are
// guarded by a runtime check for Windows, Linux and macOS, which the platform analyzer does not recognize as a guard
#pragma warning disable CA1416

namespace DocxTemplater.Pdf.Pdfium
{
    /// <summary>
    /// Renders PDF pages with PDFium (Chrome's PDF engine) through the PDFtoImage package.
    /// </summary>
    /// <remarks>
    /// PDFium is not thread-safe; PDFtoImage serializes the calls into it, so rendering in parallel document
    /// generations is throttled. Native binaries: PDFium and SkiaSharp for Windows, Linux and macOS come with the
    /// package (only Windows is tested).
    /// </remarks>
    public sealed class PdfiumPdfRenderer : IPdfRenderer
    {
        public IReadOnlyList<PdfPageSize> GetPageSizes(byte[] pdf, string password)
        {
            ArgumentNullException.ThrowIfNull(pdf);
            if (!(OperatingSystem.IsWindows() || OperatingSystem.IsLinux() || OperatingSystem.IsMacOS()))
            {
                throw new PlatformNotSupportedException("PDFium rendering is supported on Windows, Linux and macOS");
            }
            return Conversion.GetPageSizes(pdf, password)
                .Select(x => new PdfPageSize(x.Width, x.Height))
                .ToList();
        }

        public IEnumerable<RenderedPdfPage> RenderPages(byte[] pdf, PdfRenderRequest request)
        {
            ArgumentNullException.ThrowIfNull(pdf);
            ArgumentNullException.ThrowIfNull(request);
            if (!(OperatingSystem.IsWindows() || OperatingSystem.IsLinux() || OperatingSystem.IsMacOS()))
            {
                throw new PlatformNotSupportedException("PDFium rendering is supported on Windows, Linux and macOS");
            }
            var sizes = Conversion.GetPageSizes(pdf, request.Password);
            return Render(pdf, request, sizes);
        }

        private static IEnumerable<RenderedPdfPage> Render(byte[] pdf, PdfRenderRequest request, IList<System.Drawing.SizeF> sizes)
        {
            if (!(OperatingSystem.IsWindows() || OperatingSystem.IsLinux() || OperatingSystem.IsMacOS()))
            {
                throw new PlatformNotSupportedException("PDFium rendering is supported on Windows, Linux and macOS");
            }
            var options = new RenderOptions(
                Dpi: request.Dpi,
                WithAnnotations: request.WithAnnotations,
                WithFormFill: request.WithFormFields,
                AntiAliasing: PdfAntiAliasing.All,
                BackgroundColor: SKColors.White);
            var format = request.Format == PdfImageFormat.Jpeg ? SKEncodedImageFormat.Jpeg : SKEncodedImageFormat.Png;
            var quality = request.Format == PdfImageFormat.Jpeg ? Math.Clamp(request.JpegQuality, 1, 100) : 100;
            var index = 0;
            // ToImages opens the document once for all pages and renders them one by one
            foreach (var bitmap in Conversion.ToImages(pdf, request.PageIndices, request.Password, options))
            {
                var pageIndex = request.PageIndices[index++];
                using (bitmap)
                {
                    using var data = bitmap.Encode(format, quality);
                    var size = sizes[pageIndex];
                    yield return new RenderedPdfPage(pageIndex, size.Width, size.Height, data.ToArray(), request.Format);
                }
            }
        }
    }
}

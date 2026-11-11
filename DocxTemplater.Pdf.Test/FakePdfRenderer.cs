using System.Globalization;
using System.Text;

namespace DocxTemplater.Pdf.Test
{
    /// <summary>
    /// A renderer without natives: a "PDF" is a text with the page sizes in points (see <see cref="Pdf"/>), every
    /// page renders to a 1x1 pixel PNG.
    /// </summary>
    internal sealed class FakePdfRenderer : IPdfRenderer
    {
        public const double A4Width = 595.28;
        public const double A4Height = 841.89;
        public const double A3Width = 841.89;
        public const double A3Height = 1190.55;

        private const string Prefix = "FAKEPDF;";

        // 1x1 pixel png
        private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

        public List<PdfRenderRequest> Requests { get; } = new();

        public int GetPageSizesCalls { get; private set; }

        public static byte[] Pdf(params (double Width, double Height)[] pages)
        {
            return Encoding.ASCII.GetBytes(Prefix + string.Join(";", pages.Select(p => string.Create(CultureInfo.InvariantCulture, $"{p.Width}x{p.Height}"))));
        }

        public static byte[] A4Pages(int count)
        {
            return Pdf(Enumerable.Repeat((A4Width, A4Height), count).ToArray());
        }

        public IReadOnlyList<PdfPageSize> GetPageSizes(byte[] pdf, string password)
        {
            GetPageSizesCalls++;
            var text = Encoding.ASCII.GetString(pdf);
            if (!text.StartsWith(Prefix, StringComparison.Ordinal))
            {
                throw new InvalidDataException("File not in PDF format or corrupted.");
            }
            return text[Prefix.Length..].Split(';', StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Split('x'))
                .Select(x => new PdfPageSize(double.Parse(x[0], CultureInfo.InvariantCulture), double.Parse(x[1], CultureInfo.InvariantCulture)))
                .ToList();
        }

        public IEnumerable<RenderedPdfPage> RenderPages(byte[] pdf, PdfRenderRequest request)
        {
            Requests.Add(request);
            var sizes = GetPageSizes(pdf, request.Password);
            foreach (var index in request.PageIndices)
            {
                yield return new RenderedPdfPage(index, sizes[index].WidthPoints, sizes[index].HeightPoints, Png, PdfImageFormat.Png);
            }
        }
    }
}

using System.Globalization;
using System.Text;

namespace DocxTemplater.Pdf.Pdfium.Test
{
    /// <summary>
    /// Writes minimal PDFs for the tests: each page has a colored background, a frame, a 10 cm scale bar and its
    /// page number - enough to see the size, orientation and order of the pages in the generated documents.
    /// </summary>
    internal static class TestPdf
    {
        public const double A4Width = 595.28;
        public const double A4Height = 841.89;
        public const double A3Width = 841.89;
        public const double A3Height = 1190.55;
        public const double A1Width = 1683.78;
        public const double A1Height = 2383.94;

        public static byte[] Create(params (double Width, double Height, int Rotate)[] pages)
        {
            var objects = new List<string>
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                $"<< /Type /Pages /Kids [{string.Join(" ", pages.Select((_, i) => $"{4 + (i * 2)} 0 R"))}] /Count {pages.Length} >>",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            };
            for (var i = 0; i < pages.Length; i++)
            {
                var (width, height, rotate) = pages[i];
                var cm = 72 / 2.54;
                var content = string.Join("\n",
                    "0.92 0.95 1 rg", $"0 0 {F(width)} {F(height)} re f",
                    "0 0 0.6 RG 4 w", $"10 10 {F(width - 20)} {F(height - 20)} re S",
                    "0 0 0 rg", $"{F(2 * cm)} {F(2 * cm)} {F(10 * cm)} {F(0.3 * cm)} re f",
                    "BT /F1 12 Tf", $"{F(2 * cm)} {F(2.6 * cm)} Td (10 cm) Tj ET",
                    "BT /F1 48 Tf", $"{F(width / 2 - 100)} {F(height / 2)} Td (Page {i + 1}) Tj ET");
                objects.Add($"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {F(width)} {F(height)}] /Rotate {rotate} /Resources << /Font << /F1 3 0 R >> >> /Contents {5 + (i * 2)} 0 R >>");
                objects.Add($"<< /Length {content.Length} >>\nstream\n{content}\nendstream");
            }

            var pdf = new StringBuilder("%PDF-1.4\n");
            var offsets = new List<int>();
            for (var i = 0; i < objects.Count; i++)
            {
                offsets.Add(pdf.Length);
                pdf.Append(CultureInfo.InvariantCulture, $"{i + 1} 0 obj\n{objects[i]}\nendobj\n");
            }
            var xref = pdf.Length;
            pdf.Append(CultureInfo.InvariantCulture, $"xref\n0 {objects.Count + 1}\n0000000000 65535 f \n");
            foreach (var offset in offsets)
            {
                pdf.Append(CultureInfo.InvariantCulture, $"{offset:D10} 00000 n \n");
            }
            pdf.Append(CultureInfo.InvariantCulture, $"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
            return Encoding.ASCII.GetBytes(pdf.ToString());
        }

        private static string F(double value)
        {
            return value.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}

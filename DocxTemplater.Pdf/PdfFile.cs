using System;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// A PDF file with a name, as a model value for <c>{{ds.Plan}:pdf}</c>. The name is used for the alt text of
    /// the inserted pages and for the placeholder overlay of plan pages. A plain <c>byte[]</c> works as well.
    /// </summary>
    /// <param name="Content">The bytes of the PDF file. Empty content removes the placeholder.</param>
    /// <param name="Name">A readable name, e.g. "Ground floor - Plan 102-A". May be null.</param>
    public sealed record PdfFile(byte[] Content, string Name)
    {
        public byte[] Content { get; init; } = Content ?? throw new ArgumentNullException(nameof(Content));
    }
}

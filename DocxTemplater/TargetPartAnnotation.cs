using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocxTemplater
{
    /// <summary>
    /// Marks a detached element tree with the part it will be inserted into.
    /// Formatters that run on content before it is attached to the document (e.g. placeholders inside
    /// HTML converted by DocxTemplater.Html) cannot find their part through the element's root - this
    /// annotation on the root of the detached tree tells them where referenced parts (images, ...) must go.
    /// </summary>
    public sealed class TargetPartAnnotation
    {
        public TargetPartAnnotation(OpenXmlPart part)
        {
            Part = part;
        }

        public OpenXmlPart Part { get; }

        /// <summary>
        /// The part that owns the element: the part of its root element, or - for a detached tree - the part
        /// annotated on its root. Null if neither is known.
        /// </summary>
        public static OpenXmlPart GetTargetPart(OpenXmlElement element)
        {
            var root = element.GetRoot();
            return (root as OpenXmlPartRootElement)?.OpenXmlPart ?? root.Annotation<TargetPartAnnotation>()?.Part;
        }
    }
}

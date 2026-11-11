using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;

namespace DocxTemplater.Pdf
{
    /// <summary>
    /// Inserts the pages of a PDF as pictures. Usage in the template: <c>{{ds.Contract}:pdf}</c>,
    /// <c>{{ds.Contract}:pdf(pages:'1-3,7', dpi:200)}</c>, <c>{{ds.FloorPlan}:pdf(fit:section)}</c> (true size on
    /// pages of their own). The value is a <c>byte[]</c> or a <see cref="PdfFile"/>.
    /// </summary>
    public class PdfFormatter : IFormatter
    {
        private readonly IPdfRenderer m_renderer;
        private readonly PdfFormatterConfiguration m_configuration;

        // per target part: rendered pages by PDF content + render options, so the same PDF (e.g. in a loop) is
        // rendered and stored only once
        private readonly ConditionalWeakTable<OpenXmlPart, Dictionary<string, Dictionary<int, PlacedPage>>> m_imageCache = new();
        private readonly ConditionalWeakTable<OpenXmlPart, StrongBox<uint>> m_lastIds = new();
        private readonly ConditionalWeakTable<MainDocumentPart, Tuple<string, string>> m_emptyHeaderFooters = new();

        public PdfFormatter(IPdfRenderer renderer, PdfFormatterConfiguration configuration = null)
        {
            m_renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));
            m_configuration = configuration ?? PdfFormatterConfiguration.Default;
        }

        public bool CanHandle(Type type, string prefix)
        {
            return (type == typeof(byte[]) || type == typeof(PdfFile)) && "PDF".Equals(prefix, StringComparison.OrdinalIgnoreCase);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext, Text target)
        {
            var (pdf, name) = formatterContext.Value switch
            {
                byte[] bytes => (bytes, null),
                PdfFile file => (file.Content, file.Name),
                _ => ((byte[])null, (string)null),
            };
            if (pdf == null)
            {
                return;
            }
            if (pdf.Length == 0)
            {
                target.Text = string.Empty;
                return;
            }

            var configuration = PdfArguments.Apply(m_configuration, formatterContext.Args);
            var label = !string.IsNullOrWhiteSpace(configuration.Label) ? configuration.Label
                : !string.IsNullOrWhiteSpace(name) ? name
                : configuration.DefaultLabel;

            if (target.GetFirstAncestor<Paragraph>() == null || PdfPlacement.IsInTextBox(target))
            {
                throw new OpenXmlTemplateException($"pdf placeholder '{formatterContext.Placeholder}' must be in a paragraph of the document body, a table cell, a header or a footer - text boxes and shapes are not supported");
            }
            if (configuration.Fit == PdfFit.Section)
            {
                // fail before rendering
                SectionLayout.EnsureSupported(target);
            }

            var mainDocumentPart = templateContext.MainDocumentPart;
            var part = TargetPartAnnotation.GetTargetPart(target) ?? mainDocumentPart;
            if (part is not MainDocumentPart and not HeaderPart and not FooterPart)
            {
                throw new OpenXmlTemplateException($"pdf pages can only be inserted into the document body, headers and footers, not into {part.GetType().Name}");
            }

            var password = configuration.Password ?? configuration.PasswordProvider?.Invoke(label);
            var pageSizes = Call(label, () => m_renderer.GetPageSizes(pdf, password));
            if (pageSizes == null || pageSizes.Count == 0)
            {
                throw new OpenXmlTemplateException($"PDF '{label}' has no pages");
            }
            var pageIndices = PdfArguments.ParsePageSelection(configuration.Pages, pageSizes.Count);
            if (pageIndices.Count == 0)
            {
                target.Text = string.Empty;
                return;
            }
            if (pageIndices.Count > configuration.MaxPages)
            {
                throw new OpenXmlTemplateException($"PDF '{label}': {pageIndices.Count} pages selected, at most {configuration.MaxPages} are allowed (PdfFormatterConfiguration.MaxPages)");
            }

            var pages = GetPages(part, pdf, password, label, pageIndices, pageSizes, configuration);

            var ids = m_lastIds.GetValue(part, _ => new StrongBox<uint>(0));
            var firstId = Math.Max(part.GetMaxDocPropertyId(), ids.Value) + 1;
            var context = new PdfInsertContext(configuration, mainDocumentPart, label, pageSizes.Count, formatterContext.Culture, firstId,
                () => GetEmptyHeaderFooter(mainDocumentPart));

            if (configuration.Fit == PdfFit.Section)
            {
                SectionLayout.Insert(target, pages, context);
            }
            else
            {
                FlowLayout.Insert(target, pages, context);
            }
            ids.Value = Math.Max(ids.Value, context.LastId);

            if (configuration.DisableImageCompression)
            {
                DisableImageCompression(mainDocumentPart);
            }
        }

        private List<PlacedPage> GetPages(OpenXmlPart part, byte[] pdf, string password, string label, List<int> pageIndices, IReadOnlyList<PdfPageSize> pageSizes, PdfFormatterConfiguration configuration)
        {
            var dpiPerPage = pageIndices.Distinct().ToDictionary(x => x, x => GetDpi(x, pageSizes[x], configuration, label));
            var key = string.Join("|",
                Convert.ToHexString(SHA256.HashData(pdf)),
                configuration.Dpi,
                configuration.MaxPixelsPerPage,
                configuration.ImageFormat,
                configuration.ImageFormat == PdfImageFormat.Jpeg ? configuration.JpegQuality : 0,
                configuration.WithAnnotations,
                configuration.WithFormFields);
            var cache = m_imageCache.GetValue(part, _ => new Dictionary<string, Dictionary<int, PlacedPage>>());
            if (!cache.TryGetValue(key, out var placed))
            {
                placed = new Dictionary<int, PlacedPage>();
                cache[key] = placed;
            }

            // render the missing pages, one call per run of pages with the same dpi (the renderer opens the document once per call)
            var missing = pageIndices.Distinct().Where(x => !placed.ContainsKey(x)).ToList();
            foreach (var (runPages, runDpi) in SplitByDpi(missing, dpiPerPage))
            {
                var request = new PdfRenderRequest(runPages, runDpi, configuration.ImageFormat, configuration.JpegQuality,
                    configuration.WithAnnotations, configuration.WithFormFields, password);
                var rendered = Call(label, () => m_renderer.RenderPages(pdf, request)?.ToList() ?? []);
                if (rendered.Count != runPages.Count)
                {
                    throw new OpenXmlTemplateException($"PDF '{label}': the renderer returned {rendered.Count} instead of {runPages.Count} pages");
                }
                foreach (var page in rendered)
                {
                    if (page.ImageBytes == null || page.ImageBytes.Length == 0)
                    {
                        throw new OpenXmlTemplateException($"PDF '{label}': page {page.PageIndex + 1} could not be rendered");
                    }
                    var size = pageSizes[page.PageIndex];
                    var relationshipId = AddImagePart(part, page.ImageBytes, page.Format);
                    placed[page.PageIndex] = new PlacedPage(page.PageIndex, size.WidthPoints, size.HeightPoints, relationshipId);
                }
            }
            return pageIndices.Select(x => placed[x]).ToList();
        }

        private static IEnumerable<(List<int> Pages, int Dpi)> SplitByDpi(List<int> pages, Dictionary<int, int> dpiPerPage)
        {
            var current = new List<int>();
            var currentDpi = 0;
            foreach (var page in pages)
            {
                if (current.Count > 0 && dpiPerPage[page] != currentDpi)
                {
                    yield return (current, currentDpi);
                    current = new List<int>();
                }
                currentDpi = dpiPerPage[page];
                current.Add(page);
            }
            if (current.Count > 0)
            {
                yield return (current, currentDpi);
            }
        }

        private static int GetDpi(int pageIndex, PdfPageSize size, PdfFormatterConfiguration configuration, string label)
        {
            var dpi = Math.Clamp(configuration.Dpi, 10, 1200);
            var pixels = size.WidthPoints / 72.0 * dpi * (size.HeightPoints / 72.0 * dpi);
            if (configuration.MaxPixelsPerPage <= 0 || pixels <= configuration.MaxPixelsPerPage)
            {
                return dpi;
            }
            if (!configuration.ReduceDpiForLargePages)
            {
                throw new OpenXmlTemplateException($"PDF '{label}': page {pageIndex + 1} would have {pixels / 1e6:0.#} megapixels at {dpi} dpi, at most {configuration.MaxPixelsPerPage / 1e6:0.#} are allowed (PdfFormatterConfiguration.MaxPixelsPerPage)");
            }
            return Math.Max(10, (int)Math.Floor(dpi * Math.Sqrt(configuration.MaxPixelsPerPage / pixels)));
        }

        private static T Call<T>(string label, Func<T> action)
        {
            try
            {
                return action();
            }
            catch (OpenXmlTemplateException)
            {
                throw;
            }
            catch (Exception e)
            {
                throw new OpenXmlTemplateException($"PDF '{label}' could not be read: {e.Message}", e);
            }
        }

        private static string AddImagePart(OpenXmlPart part, byte[] imageBytes, PdfImageFormat format)
        {
            var partType = format == PdfImageFormat.Jpeg ? ImagePartType.Jpeg : ImagePartType.Png;
            var imagePart = part switch
            {
                MainDocumentPart mainDocumentPart => mainDocumentPart.AddImagePart(partType),
                HeaderPart headerPart => headerPart.AddImagePart(partType),
                FooterPart footerPart => footerPart.AddImagePart(partType),
                _ => throw new OpenXmlTemplateException($"pdf pages can not be inserted into {part.GetType().Name}"),
            };
            using (var stream = new MemoryStream(imageBytes, false))
            {
                imagePart.FeedData(stream);
            }
            return part.GetIdOfPart(imagePart);
        }

        /// <summary>
        /// One empty header and footer per document, shared by all sections created for <c>fit:section</c>.
        /// </summary>
        private (string HeaderId, string FooterId) GetEmptyHeaderFooter(MainDocumentPart mainDocumentPart)
        {
            var ids = m_emptyHeaderFooters.GetValue(mainDocumentPart, main =>
            {
                var headerPart = main.AddNewPart<HeaderPart>();
                headerPart.Header = new Header(CreateEmptyParagraph());
                var footerPart = main.AddNewPart<FooterPart>();
                footerPart.Footer = new Footer(CreateEmptyParagraph());
                return Tuple.Create(main.GetIdOfPart(headerPart), main.GetIdOfPart(footerPart));
            });
            return (ids.Item1, ids.Item2);
        }

        private static Paragraph CreateEmptyParagraph()
        {
            // 1 pt, so the empty header and footer reserve no space
            return new Paragraph(new ParagraphProperties
            {
                SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0", Line = "20", LineRule = LineSpacingRuleValues.Exact },
                ParagraphMarkRunProperties = new ParagraphMarkRunProperties(new FontSize { Val = "2" }, new FontSizeComplexScript { Val = "2" }),
            });
        }

        private static void DisableImageCompression(MainDocumentPart mainDocumentPart)
        {
            var settingsPart = mainDocumentPart.DocumentSettingsPart ?? mainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings ??= new Settings();
            if (settingsPart.Settings.GetFirstChild<DoNotAutoCompressPictures>() == null)
            {
                settingsPart.Settings.AddChild(new DoNotAutoCompressPictures());
            }
        }
    }
}

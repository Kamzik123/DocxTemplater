# Writing DocxTemplater Modules

Notes from building `DocxTemplater.Html`, our replacement for the commercial HTML module. The goal is to make the next module (or a rework of this one) quick to write without re-reading the core.

- [What was done](#what-was-done)
- [How a module plugs into the core](#how-a-module-plugs-into-the-core)
- [Anatomy of a block-producing formatter](#anatomy-of-a-block-producing-formatter)
- [OpenXML pitfalls we hit](#openxml-pitfalls-we-hit)
- [Build, analyzers and tests](#build-analyzers-and-tests)
- [Packaging and replacing the official packages](#packaging-and-replacing-the-official-packages)
- [Checklist for a new module](#checklist-for-a-new-module)
- [Ideas for future work](#ideas-for-future-work)

---

## What was done

### Background
- The core library and the Markdown/Images modules are MIT licensed (`LICENSE`, and `<license type="expression">MIT</license>` in the nuspec of `DocxTemplater.Markdown 2.9.7`). The Markdown package was built from commit `94ed068`, which is exactly the base of this work. Its source is in `DocxTemplater.Markdown/`, so no decompiling was needed.
- The core already contains a free `html` formatter (`DocxTemplater/Formatter/HtmlFormatter.cs`). It stores the raw HTML as an `AlternativeFormatImportPart` and inserts a `w:altChunk`, and Word converts it when the document is opened. Consequences:
  - LibreOffice, PDF converters and OpenXML consumers do not render it.
  - Template styles are not applied, and placeholders inside the HTML are not processed.
  - The content is not real document content until Word saves the file again.
- There is no `DocxTemplater.Html` package on nuget.org (checked 2026-09-30), so the package ID was free.

### Core changes (small, kept minimal)
| File | Change | Why |
|------|--------|-----|
| `DocxTemplater/Formatter/VariableReplacer.cs` | `RegisterFormatter` inserts registered formatters **before** the built-in ones (registration order among registered formatters is kept) | Formatters are matched first-come-first-served via `CanHandle`. Built-ins were first, so a module could never override the `html` prefix |
| `DocxTemplater/InternalsVisibleTo.cs` | Added `DocxTemplater.Html` and `DocxTemplater.Html.Test` | The module needs `XmlNodeTemplate` (nested placeholders) and `HelperFunctions.ParseArguments`, both internal |

All 477 existing core tests still pass on net8/9/10 after these changes.

### New projects
| Project | Content |
|---------|---------|
| `DocxTemplater.Html` (net10.0) | `HtmlFormatter` (IFormatter + insertion into the template), `HtmlToOpenXmlConverter` (DOM walk → OpenXML; tables in `HtmlToOpenXmlConverter.Tables.cs`), `ListNumbering` (numbering.xml management), `CssParser` (inline styles, colors, units), `CssBoxParser` (css borders and padding), `HtmlFormatterConfiguration` |
| `DocxTemplater.Html.Test` (net10.0) | 25 tests + an `[Explicit]` sample document generator (`SampleDocumentTest`, set `DOCX_TEMPLATER_HTML_SAMPLE=<path>`) |

The only dependency is **AngleSharp** (MIT, `[1.8.2, 2.0.0)`), an HTML5-spec parser. It fixes broken HTML the way a browser does (unclosed tags, `<p>` inside `<b>`, stray `<li>`, ...), which matters because HTML from rich-text editors and users is often malformed.

The module targets **only net10.0** (per decision), overriding the repo-wide `netstandard2.0;net8.0;net9.0;net10.0` from `Directory.Build.props`. This is why it has its own test project: `DocxTemplater.Test` runs on net8/net9 as well and cannot reference a net10-only assembly.

---

## How a module plugs into the core

There is no module registry. A module is one or more classes implementing public core interfaces, registered by the user.

### `IFormatter` - the main extension point
```csharp
public interface IFormatter
{
    bool CanHandle(Type type, string prefix);   // prefix = formatter name from {{x}:prefix(args)}
    void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext, Text target);
}
```
- Registered with `DocxTemplate.RegisterFormatter(formatter)`.
- `CanHandle` is called with the runtime type of the value and the prefix. Compare case-insensitively (`md`, `MD`, `Md` all appear in the wild).
- `FormatterContext` holds `Value`, `Args` (raw strings from `(...)`), `Culture`, `Placeholder`.
- `target` is the `w:t` element that held the placeholder. Write the result into it (inline formatters just set `target.Text`) or replace its surroundings (block formatters).
- A value can default to a formatter via `[ModelProperty(DefaultFormatter = "html")]` or `new ValueWithMetadata(value, new ValueMetadata("html"))`, so the template does not need the prefix.
- Arguments: `HelperFunctions.ParseArguments(args)` (internal) parses `key:value` pairs such as `ts:'Style'`.

### Other extension points
| Interface | Registered with | Used by |
|-----------|-----------------|---------|
| `IImageServiceProvider` (implemented by a formatter) | `RegisterFormatter` detects it and calls `Context.SetImageService(...)` | `DocxTemplater.Images`. Other modules get the service through `templateContext.ImageService` (**null if the Images module is not registered**) |
| `ITemplateProcessorExtension` | `DocxTemplate.RegisterExtension` | Hooks into template processing |
| `IImageMetadataReader` | `new ImageFormatter(reader)` | `DocxTemplater.Images.ImageSharp` |

### Useful public helpers (`OpenXmlHelper`, all extension methods)
- `SplitAfterElement` / `SplitBeforeElement`: split a paragraph at a descendant. The first part is the original element; a second part exists only if there is content after the element.
- `GetFirstAncestor<T>`, `GetRoot`, `RemoveWithEmptyParent(preserveParagraphInCell)`, `HasOnlyPropertyChildren`
- `FindTableStyleByName` (matches style id **or** name), `GetMaxDocPropertyId`, `PixelsToEmu`
- `CreateNewAbstractNumbering` / `CreateNewNumberingInstance`: **careful**, see [pitfalls](#openxml-pitfalls-we-hit).

### Internals a block module usually needs
Grant access in `DocxTemplater/InternalsVisibleTo.cs`.
- `XmlNodeTemplate(OpenXmlCompositeElement, ITemplateProcessingContextAccess).Process()` runs the full template engine on a detached element. This is how `{{...}}` inside Markdown/HTML gets replaced (including loops and conditions).
- `HelperFunctions.ParseArguments`

---

## Anatomy of a block-producing formatter

Markdown and HTML both follow the same four steps. See `DocxTemplater.Html/HtmlFormatter.cs`.

1. **Capture the template context** before touching anything:
   - `target.GetFirstAncestor<Run>().RunProperties`: the placeholder's font, size and color. Clone it into every generated run so the content matches the surrounding text, but skip headings, where it would override the heading style.
   - `targetParagraph.ParagraphProperties`: clone it for "plain" generated paragraphs. **Remove `SectionProperties` from the clone.** If the placeholder paragraph ends a section, copying it would create a new section per paragraph.
   - The part: `(target.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart`. This is `MainDocumentPart`, `HeaderPart` or `FooterPart`. Hyperlink and image relationships must be added to **this** part, not the main document part, or links in headers break. Numbering and styles always live in `templateContext.MainDocumentPart`.
2. **Render into a detached `Body`** (a scratch container). This keeps the renderer simple and lets step 3 run on it.
3. **Replace nested placeholders** with `XmlNodeTemplate` on that body. Only do this if `InnerText` contains `{` (cheap check). Make it configurable, because it is a template-injection vector for untrusted input.
4. **Splice into the document**:
   - Split the template paragraph after `target`.
   - Merge the first generated paragraph's runs into the template paragraph after the placeholder run, so `Hello {{x}:html}` + `<b>World</b>` stays one paragraph.
   - Insert the remaining elements after it.
   - Append the template content after the placeholder to the last generated paragraph.
   - Only merge a generated paragraph with template text if it has no formatting of its own. The HTML module compares `pPr.OuterXml` with the template's `pPr`. Otherwise `Intro: {{x}:html}` + `<h1>` would turn "Intro:" into a heading (the Markdown module has this problem).
   - Detect "content before the placeholder" by looking at all descendants of the paragraph before `target`, **not** at run siblings. The core isolates a placeholder into its own `w:t` but often leaves it in the **same** `w:r` as the surrounding text.
   - Finally `target.RemoveWithEmptyParent()`. If the placeholder was in a table cell, make sure the cell still ends with a `w:p`.

Guard against recursion: a nested placeholder can produce HTML that contains placeholders that produce HTML... Both modules keep a nesting depth counter (max 3).

### Converter design (HTML specifics)
- **Two format stacks**: `RunFormat` (inline: bold, color, size, font, ... as nullable values; `null` = inherit) and `BlockFormat` (paragraph: alignment, indent, style, list item, preformatted, ...). Elements push a modified copy (`record ... with {}`) and pop it afterwards.
- **Lazy paragraphs**: a paragraph is created only when inline content arrives (`EnsureParagraph`), and block elements just close the current one (`FinishParagraph`). This avoids empty paragraphs from whitespace between tags and gives browser-like results for `<p></p>` (nothing) versus `<p><br></p>` (an empty line).
- **Whitespace collapsing** needs a single "last char was space" flag per paragraph. Leading spaces of a paragraph and spaces after `<br>` are dropped, and trailing spaces are trimmed in `FinishParagraph`. Only real HTML whitespace collapses (` ` must survive). *A bug we had:* resetting the flag inside `EnsureParagraph` swallowed the space after `</b>`, because the paragraph was created lazily **after** the text had been collapsed.
- **A trailing `<br>` in a block is dropped** (browsers ignore it as well).
- **List items**: a mutable `ListItemState` is shared by all paragraphs of one `<li>`. Only the first paragraph gets `numPr`, and later ones (e.g. text after a nested list) get an indent instead.
- **Tables**: first place the cells on a grid (`LayoutTable`). HTML `rowspan` becomes a `vMerge restart` cell plus `vMerge` continuation cells in each following row. `colspan` becomes `gridSpan`. Pad short rows, because Word tolerates ragged rows but LibreOffice renders them badly.
- **Table borders have two modes.** Without any border information in the html, the table keeps the template's table style (most templates want that). With *any* border information (`border`/`frame`/`rules` attributes or css borders on table, row or cell), the html defines every line, and each cell gets explicit `tcBorders`: cell css → row css → table frame (outer edges) / rules (inner edges), with `nil` for "no line". The explicit cell borders are needed because a table style can draw lines through conditional formatting (header row, banding) that `tblBorders` alone does not override. *The bug this fixed:* `border="0"` was ignored whenever the template had a "Table Grid" style.
- **Column widths** are resolved like a browser: `<col>`/`<colgroup>` first, then single-column cells; unknown columns share the rest. Cell widths are always the sum of their grid columns in dxa, so grid and cells never disagree. Cell `padding` goes to `tcMar` and is removed from the css used for the cell's paragraphs. Otherwise `padding-left` would also indent the text.
- **Headings**: find the style by the **built-in English name** (`heading 1`) first, then by id. Word localizes style *ids* (German `berschrift1`, Czech `Nadpis1`) but not the built-in *names*. If no style exists, use direct formatting plus `outlineLevel` (so the navigation pane and TOC still work) and `keepNext`.

---

## OpenXML pitfalls we hit

| Pitfall | Symptom | Fix |
|---------|---------|-----|
| Assigning `null` to a `StringValue` attribute in an object initializer (`new Hyperlink { Anchor = null }`) | The attribute is written **empty** and `Validate()` fails: *"The attribute value cannot be empty"* | Only assign when non-null |
| `IImageService.GetImage` returns `GetMaxDocPropertyId()` of the part, but images rendered into a detached body are not in the part yet | Several images in one value get the same `wp:docPr id`, and Word reports a corrupt file | Track the next id yourself: `id = max(returned + 1, m_nextId)` |
| `OpenXmlHelper.CreateNewNumberingInstance` uses `count + 1` as the id | Duplicate `w:numId` if ids are not contiguous | Use `max(existing ids) + 1` (see `ListNumbering.CreateInstance`) |
| `CreateNewAbstractNumbering` appends at the end if there is no `abstractNum` yet | `w:abstractNum` after `w:num` is a schema error (order: `numPicBullet*`, `abstractNum*`, `num*`, `numIdMacAtCleanup`) | Insert explicitly at the right position |
| Creating new `abstractNum`s on every call | numbering.xml grows with every placeholder in a loop | Tag them with `w:name` (`AbstractNumDefinitionName`) and reuse them |
| Numbering restart | Consecutive `<ol>`s continue each other's numbers | One `w:num` per ordered list with `w:lvlOverride/w:startOverride` |
| Element order inside `pPr` / `rPr` / `tblPr` / `tcPr` | Schema validation errors | Use the SDK's typed properties (`properties.Bold = ...`, `properties.Justification = ...`). They insert in schema order. Avoid `AppendChild` on property containers |
| Table cell ending with a table | Word "unreadable content" | A `w:tc` must end with a `w:p`. Append an empty paragraph |
| Copying the template `pPr` with a `sectPr` | Extra section breaks | Remove `SectionProperties` from clones |
| Test templates built in code | `Validate()` complains about the *template*, not your output | Build valid templates: `w:tbl` needs `tblPr` and `tblGrid` |

Always call `docTemplate.Validate()` in tests. It runs `OpenXmlValidator` on the whole package and catches most of the above.

---

## Build, analyzers and tests

- `Directory.Build.props` enables almost all .NET analyzers with `EnforceCodeStyleInBuild`, and **Release treats warnings as errors**. Always build new code with `-c Release` before calling it done. `dotnet format <project>` fixes most formatting (IDE0055) automatically. Rules that needed manual fixes: collection expressions (IDE0300/IDE0306), no expression-bodied local functions (IDE0061), static readonly arrays instead of inline array arguments (CA1861), deconstruction (IDE0042), and auto properties (IDE0032).
- In **Debug** builds the core prints the generated XML to the console (`#if DEBUG` in the Markdown formatter and core). That is handy but noisy, so filter the test output.
- To inspect a converter in isolation, grant `InternalsVisibleTo` to the test project and call it directly with `null` for parts/services.
- `TestHelper.SaveAsFileAndOpenInWord()` in the core tests opens results in Word when `DOCX_TEMPLATER_VISUAL_TESTING` is set.
- Run everything: `dotnet test DocxTemplater.slnx -c Release`.

---

## Packaging and replacing the official packages

- **Build the packages with `dotnet build`, not `dotnet pack`.** The projects set `GeneratePackageOnBuild`, and in that mode `dotnet pack` does not recompile. It packs the DLLs that are already in `bin/Release`, so the package gets the new version while the **DLLs inside keep the old assembly version** (we shipped `2.5.0.0` DLLs in a `2.9.7.1` package once, and the official Markdown DLL then failed with *Could not load file or assembly 'DocxTemplater, Version=2.9.7.0'*). Use:
  ```
  dotnet build <project> -c Release --no-incremental -p:Version=2.9.7.1 -p:PackageOutputPath=<folder>
  ```
  `-p:Version` overrides the default `2.5.0-local` from `Directory.Build.props`. SourceLink and `.snupkg` symbol packages are produced automatically. Check the DLL versions inside the `.nupkg` before shipping.
- **Never re-publish a version number with different content.** NuGet caches every restored package in `%UserProfile%\.nuget\packages\<id>\<version>` and never looks at the source again for that version. After a fix, bump the version (`2.9.7.2`, ...) or delete exactly that cache folder on every machine that restored it.
- **Use a version above the official one** (we use `2.9.7.1`, 4-part = "fork of 2.9.7"):
  - The same version (`2.9.7`) would be served from the global packages cache (`~/.nuget/packages/docxtemplater/2.9.7`) with the **official** content.
  - A lower version would conflict with the official `DocxTemplater.Markdown 2.9.7`, which depends on `DocxTemplater >= 2.9.7` (NU1605 downgrade error).
- **Keep the package IDs** (`DocxTemplater`, `DocxTemplater.Images`). Renaming the core would load two different `DocxTemplater.dll`s next to each other as soon as an official module is installed.
- The official `DocxTemplater.Markdown 2.9.7` works with our core `2.9.7.1`: same source commit, and the assemblies are not strong-named.

---

## Checklist for a new module

1. `DocxTemplater.<Name>/DocxTemplater.<Name>.csproj` with `GeneratePackageOnBuild`, a `ProjectReference` to `..\DocxTemplater\DocxTemplater.csproj`, and version ranges for dependencies (`[x.y.z, next-major)`).
2. Add the assembly (and its test assembly) to `DocxTemplater/InternalsVisibleTo.cs` if it needs internals.
3. Implement `IFormatter`. Use a case-insensitive prefix, check the value type, parse args with `HelperFunctions.ParseArguments`, and add a configuration class with `Clone()` for per-placeholder overrides.
4. For block content, follow the [four steps](#anatomy-of-a-block-producing-formatter), reusing `HtmlFormatter.InsertIntoTemplate` as a template.
5. Add a test project (NUnit, same package versions as `DocxTemplater.Test`), and call `Validate()` in every test.
6. Add both projects to `DocxTemplater.slnx` (`dotnet sln DocxTemplater.slnx add ...`).
7. `dotnet test DocxTemplater.slnx -c Release` must pass without warnings.
8. Document it in `README.md` (package table, quick reference, own section).

---

## Ideas for future work

- **Share the insertion logic.** `MarkdownFormatter.ApplyFormat` and `HtmlFormatter.InsertIntoTemplate` do the same splice. Moving an improved version (the HTML one, with the heading/merge rules) into the core as a public helper, e.g. `BlockContentInserter.Insert(target, Body)`, would make every future block module smaller and give Markdown the heading fix.
- **Fix `CreateNewNumberingInstance` / `CreateNewAbstractNumbering` in the core** (id and ordering issues above) instead of working around them.
- **Fix `GetImage` doc-property ids for detached content** in `ImageService`, which affects the Markdown module too (several images in one Markdown value).
- **HTML module gaps:**
  - `<style>` blocks / CSS classes (would need AngleSharp.Css).
  - Remote images through an opt-in resolver callback (`Func<Uri, byte[]>`). There is intentionally no network access by default.
  - `<a name>` / `id` → bookmarks, so `#anchor` links have targets.
  - `dir="rtl"` (bidi), `<q>` quotes, and definition lists with a hanging indent.
  - Line height and margins (`line-height`, `margin-top/bottom` → `w:spacing`).
- **Other module candidates** follow the same `IFormatter` pattern: QR/barcodes (image via `IImageService`), charts from data, RTF snippets, LaTeX/MathML → OMML equations.
- **Multi-targeting.** The HTML module is net10-only by choice. AngleSharp supports netstandard2.0, so adding the repo-wide targets back is a one-line change if older runtimes are ever needed. The code avoids net10-only APIs except collection expressions and `Math.Clamp`, which PolySharp/Polyfill cover.

# Tesserae.DocxEditor

**Tesserae.DocxEditor** is a [Tesserae](https://github.com/curiosity-ai/tesserae) (H5 C#-to-JavaScript)
wrapper around the [docx-editor](https://github.com/theolivenbaum/docx-editor) library.

It lets you drop a fully interactive **WYSIWYG DOCX editor** into a Tesserae app from C#, with no
JavaScript — and with **no JavaScript UI framework** (no React, no Vue). The editor engine
(ProseMirror) and the DOCX parser/serializer are bundled into this NuGet package as a browser-global
script (`window.docxeditor`) and embedded through `h5.json`, so there is **no preload step** —
referencing the package is enough.

Two faces are exposed:

- **`DocxEditor`** — a visual, editable WYSIWYG component (a Tesserae `IComponent`) with a toolbar
  composed entirely from Tesserae components. Open a `.docx`, edit it, save it back — round-tripping
  the real document model (fonts, styles, tables, …).
- **`DocxDocument`** — a headless, programmatic API (wrapping the library's `DocumentAgent`) for
  reading, editing, templating and saving documents with no UI.

## Usage — the visual editor

```csharp
using Tesserae.DocxEditor;
using static Tesserae.UI;

var editor = DOCX.Editor()
    .WithText("Type here, use the toolbar, then download as .docx.")
    .OnChange(() => Console.WriteLine("document changed"));

document.body.appendChild(
    Stack().S().Children(
        HStack().Children(
            Button("Download").OnClick(() => editor.Download("my.docx"))
        ),
        editor.Grow()
    ).Render()
);
```

`DocxEditor` is a regular `IComponent`, so the usual Tesserae sizing helpers
(`.W()`, `.H()`, `.WS()`, `.HS()`, `.S()`, `.Grow()`, …) apply.

### `DocxEditor` API

| Area        | Methods |
|-------------|---------|
| Setup       | `WithText(text)`, `WithDocx(arrayBuffer)`, `ReadOnly(bool)`, `NoToolbar()`, `OnChange(handler)`, `OnReady(handler)` |
| Documents   | `Load(arrayBuffer, onLoaded?)`, `NewDocument(text?)`, `Save(onSaved)`, `Download(filename)`, `GetText()`, `GetWordCount()` |
| Marks       | `ToggleBold`, `ToggleItalic`, `ToggleUnderline`, `ToggleStrikethrough`, `SetTextColor(rgbHex)`, `SetHighlight(color)`, `SetFontSize(pt)`, `SetFontFamily(name)` |
| Paragraph   | `SetAlignment("left"\|"center"\|"right"\|"justify")`, `ApplyStyle(styleId)`, `ToggleBulletList`, `ToggleNumberedList`, `IncreaseIndent`, `DecreaseIndent` |
| Insert      | `InsertTable(rows, cols)`, `InsertPageBreak()` |
| History     | `Undo()`, `Redo()` |
| Escape hatch| `Controller` — the raw `DocxEditorController` JS instance |

The built-in toolbar (Bold/Italic/Underline, alignment, lists, paragraph-style dropdown, table,
page break, undo/redo, live word count) is built from Tesserae `Button`/`Dropdown`/`Stack`
components and reflects the current selection's formatting. Pass `NoToolbar()` to supply your own.

## Usage — the headless API

```csharp
var doc = DOCX.NewDocument("Invoice for {customer}, total {amount}.");
Console.WriteLine(doc.GetWordCount());                 // 4
Console.WriteLine(string.Join(", ", doc.GetVariables())); // amount, customer

doc.InsertText(0, 0, "DRAFT — ").ApplyStyle(0, "Heading1");
doc.Download("invoice.docx");

// Load + template-fill an existing .docx (needs the original file bytes):
var loaded = await DOCX.LoadAsync(arrayBuffer);
var filled = await loaded.ApplyVariablesAsync(new Dictionary<string, string> {
    { "customer", "Jane Doe" }, { "amount", "$1,200" },
});
var bytes = await filled.ToBytesAsync();
```

## How the JavaScript is embedded

`build/bundle.mjs` uses [esbuild](https://esbuild.github.io/) to produce two browser-global IIFE
bundles (each in a readable `.js` and a minified `.min.js` variant — H5 picks the readable one for
Debug builds and the minified one for Release):

| File | Global | Contents |
|------|--------|----------|
| `assets/js/docx-editor-deps.js` | `window.docxeditordeps` | the pure-JS document libraries: jszip, pizzip, xml-js, docxtemplater |
| `assets/js/docx-editor.js`      | `window.docxeditor`     | the editor engine (ProseMirror) + the docx-editor core + the vanilla `DocxEditorController` (see `src/js/index.ts`) |

The heavy, rarely-changing third-party libraries are kept in the **separate** `docx-editor-deps`
bundle; the main bundle externalizes them and reads them from the global at runtime, so the docx
bundle itself stays lean and the vendor libs cache independently. `h5.json` lists the deps bundle
**before** the main bundle so it loads first.

**dompurify is not bundled.** Tesserae already ships it as the global `window.DOMPurify` (used by
Tesserae's own `SanitizeHTML`/`Markdown` helpers), so the docx bundle reuses that single copy instead
of shipping a second one. The csproj refreshes both bundles before each build via the
`BundleDocxEditorJs` MSBuild target.

No JavaScript UI framework (React/Vue) is bundled anywhere — only the ProseMirror editor engine and
the pure-JS document libraries above.

## Building

```bash
npm install      # once, to fetch esbuild (from the Tesserae.DocxEditor/ folder)
npm run bundle   # regenerate assets/js/docx-editor*.js
dotnet build     # builds the H5 package (also runs the bundle step)
```

## Scope note

This wrapper renders the editor with ProseMirror's own DOM (a clean, continuous-flow page), reusing
the exact DOCX schema, conversions and formatting commands the official React/Vue editors use, so
documents open and save faithfully. The page-exact paginated *painter* view that the React/Vue
adapters render is driven by framework-specific layout orchestration and is intentionally not part of
this pure-JS wrapper.

## License

MIT © Curiosity GmbH

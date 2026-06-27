# Tesserae.DocxEditor

**Tesserae.DocxEditor** is a [Tesserae](https://github.com/curiosity-ai/tesserae) (H5 C#-to-JavaScript) library for **editing `.docx` documents from C#** — a fully interactive WYSIWYG editor and a headless document API, with no hand-written JavaScript and **no JavaScript UI framework** (no React, no Vue).

It is a thin C# wrapper around the framework-agnostic core of the [docx-editor](https://github.com/theolivenbaum/docx-editor) library (originally [eigenpal/docx-editor](https://github.com/eigenpal/docx-editor), Apache-2.0). That core — the OOXML parser/serializer, the ProseMirror editing engine, and the formatting commands — is bundled into this package as a browser-global script (`window.docxeditor`) and embedded through `h5.json`, so there is **no preload step**: referencing the NuGet package is enough.

```bash
dotnet add package Tesserae.DocxEditor
```

## Two faces

| | Type | What it is |
| --- | --- | --- |
| **Visual** | `DocxEditor` | A WYSIWYG editor as a Tesserae `IComponent`, with a toolbar built entirely from Tesserae components. Open a `.docx`, edit it, save it back — round-tripping the real document model (fonts, theme colors, styles, tables, …). |
| **Headless** | `DocxDocument` | A programmatic API (wrapping the core's `DocumentAgent`) for reading, editing, templating and saving documents with no UI. |

Both are reached through the static `DOCX` factory, mirroring Tesserae's own `UI` entry point.

## The visual editor

```csharp
using Tesserae.DocxEditor;
using static Tesserae.UI;

var editor = DOCX.Editor()
    .WithText("Type here, use the toolbar, then download as .docx.")
    .OnChange(() => Console.WriteLine("document changed"));

document.body.appendChild(
    Stack().S().Children(
        HStack().Children(
            Button("Open .docx…").OnClick(OpenFile),
            Button("Download").OnClick(() => editor.Download("my.docx"))
        ),
        editor.Grow()
    ).Render()
);
```

`DocxEditor` is a regular `IComponent`, so the usual Tesserae sizing helpers (`.W()`, `.H()`, `.WS()`, `.HS()`, `.S()`, `.Grow()`, …) all apply. The built-in toolbar (bold/italic/underline, alignment, lists, a paragraph-style dropdown, table insert, page break, undo/redo and a live word count) is composed from Tesserae `Button`/`Dropdown`/`Stack` components and reflects the current selection's formatting. Pass `NoToolbar()` to supply your own.

### `DocxEditor` API

| Area | Methods |
| --- | --- |
| Setup | `WithText(text)`, `WithDocx(arrayBuffer)`, `ReadOnly(bool)`, `NoToolbar()`, `OnChange(handler)`, `OnReady(handler)` |
| Documents | `Load(arrayBuffer, onLoaded?)`, `NewDocument(text?)`, `Save(onSaved)`, `Download(filename)`, `GetText()`, `GetWordCount()` |
| Marks | `ToggleBold`, `ToggleItalic`, `ToggleUnderline`, `ToggleStrikethrough`, `SetTextColor(rgbHex)`, `SetHighlight(color)`, `SetFontSize(pt)`, `SetFontFamily(name)` |
| Paragraph | `SetAlignment("left"\|"center"\|"right"\|"justify")`, `ApplyStyle(styleId)`, `ToggleBulletList`, `ToggleNumberedList`, `IncreaseIndent`, `DecreaseIndent` |
| Insert | `InsertTable(rows, cols)`, `InsertPageBreak()` |
| History | `Undo()`, `Redo()` |
| Escape hatch | `Controller` — the raw `DocxEditorController` JS instance |

## The headless API

```csharp
var doc = DOCX.NewDocument("Invoice for {customer}, total {amount}.");
Console.WriteLine(doc.GetWordCount());                    // 4
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

The underlying agent is immutable: each editing method swaps in the new agent it returns and yields `this`, so calls chain naturally from C#.

### `DocxDocument` API

| Area | Methods |
| --- | --- |
| Construction | `DOCX.NewDocument(text?)`, `DOCX.LoadAsync(arrayBuffer)` |
| Reading | `GetText()`, `GetWordCount()`, `GetCharacterCount(includeSpaces?)`, `GetParagraphCount()`, `GetTableCount()`, `GetVariables()`, `GetTemplateTags()` |
| Editing | `InsertText(paragraphIndex, offset, text)`, `ApplyStyle(paragraphIndex, styleId)`, `InsertTable(paragraphIndex, offset, rows, cols)` |
| Templates | `SetVariable(name, value)`, `ApplyVariablesAsync(variables?)` |
| Export | `ToBytesAsync()`, `Download(filename)` |
| Escape hatch | `Agent` — the raw `DocumentAgent` JS instance |

## Sample app

[`Tesserae.DocxEditor.Sample`](Tesserae.DocxEditor.Sample) is a runnable Tesserae app demonstrating both faces: a live WYSIWYG editor with a Tesserae toolbar, file open/save/download, a read-only toggle, and a small headless `DocxDocument` showcase.

## How the JavaScript is embedded

`build/bundle.mjs` uses [esbuild](https://esbuild.github.io/) to produce two browser-global IIFE bundles (each as a readable `.js` and a minified `.min.js` — H5 picks the readable one for Debug builds and the minified one for Release):

| File | Global | Contents |
| --- | --- | --- |
| `assets/js/docx-editor-deps.js` | `window.docxeditordeps` | the pure-JS document libraries: jszip, pizzip, xml-js, docxtemplater |
| `assets/js/docx-editor.js` | `window.docxeditor` | the editor engine (ProseMirror) + the docx-editor core + the vanilla `DocxEditorController` (see `Tesserae.DocxEditor/src/js/index.ts`) |

The heavy, rarely-changing third-party libraries live in the **separate** `docx-editor-deps` bundle; the main bundle externalizes them and reads them from the global at runtime, so the docx bundle stays lean and the vendor libs cache independently. `h5.json` lists the deps bundle **before** the main bundle so it loads first. `dompurify` is not bundled — Tesserae already ships it as `window.DOMPurify`, so the docx bundle reuses that single copy. No JavaScript UI framework (React/Vue) is bundled anywhere — only the ProseMirror engine and the pure-JS document libraries.

## Building

```bash
npm install      # once, from Tesserae.DocxEditor/, to fetch esbuild
npm run bundle   # regenerate assets/js/docx-editor*.js
dotnet build     # builds the H5 package (also re-runs the bundle step)
```

The csproj refreshes both JS bundles before each build via the `BundleDocxEditorJs` MSBuild target, so the embedded JS always matches the current core sources. The step is no-op-safe: if node/npm are unavailable, the previously generated (and committed) bundle under `assets/js/` is used instead.

## Scope note

This wrapper renders the editor with ProseMirror's own DOM (a clean, continuous-flow page), reusing the exact DOCX schema, conversions and formatting commands the upstream React/Vue editors use, so documents open and save faithfully. The page-exact paginated *painter* view that those adapters render is driven by framework-specific layout orchestration and is intentionally not part of this pure-JS wrapper.

## Source library

The DOCX core, editing engine, and OOXML parser/serializer come from [docx-editor](https://github.com/theolivenbaum/docx-editor) (originally [eigenpal/docx-editor](https://github.com/eigenpal/docx-editor)). Its source — the `packages/core`, `packages/react`, `packages/vue` and related sources retained in this repository — is licensed Apache-2.0. See [LICENSE](LICENSE).

## License

The Tesserae.DocxEditor C# wrapper is MIT © Curiosity GmbH. The bundled docx-editor core is Apache-2.0 © its respective authors.

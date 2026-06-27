# Tesserae.DocxEditor

A [Tesserae](https://github.com/curiosity-ai/tesserae) (H5 C#-to-JavaScript) library for editing `.docx` documents from C# — a WYSIWYG editor component and a headless document API, with no JS UI framework. It wraps the framework-agnostic core of [docx-editor](https://github.com/theolivenbaum/docx-editor) (Apache-2.0), bundled into a browser-global script and embedded via `h5.json`.

## Layout

| Path | What it is |
| --- | --- |
| `Tesserae.DocxEditor/` | The C# wrapper (NuGet package). `src/Components/` — `DocxEditor` (visual `IComponent`) + `DocxDocument` (headless). `src/DOCX.cs` — the static factory. `src/js/index.ts` — the JS bundle entry (`window.docxeditor`). |
| `Tesserae.DocxEditor.Sample/` | Runnable Tesserae demo of both faces. |
| `packages/core/` | Vendored docx-editor core (OOXML parser/serializer, ProseMirror schema, commands). Bundled into the wrapper's JS — **not** consumed as a NuGet/npm dependency. |

## The JS bundle

`Tesserae.DocxEditor/src/js/index.ts` imports from `packages/core/src/...` and is bundled by `Tesserae.DocxEditor/build/bundle.mjs` (esbuild) into two browser-global IIFEs under `Tesserae.DocxEditor/assets/js/`:

- `docx-editor-deps.js` → `window.docxeditordeps` — vendor libs (jszip, pizzip, xml-js, docxtemplater). Loads first.
- `docx-editor.js` → `window.docxeditor` — ProseMirror engine + core + the `DocxEditorController`. Reads the deps from the global.

`dompurify` is not bundled — Tesserae provides `window.DOMPurify`. No React/Vue is bundled.

The committed `assets/js/*.js` (and `.min.js`) are what `h5.json` embeds, so **the dotnet build works from the committed bundles even without node/npm**. The csproj's `BundleDocxEditorJs` target regenerates them when node/npm are available (no-op-safe otherwise).

## Build

```bash
# Regenerate the JS bundle (needs node/npm + core deps):
bun install                                  # populates packages/core/node_modules
cd Tesserae.DocxEditor && npm install && npm run bundle

# Build the NuGet package:
dotnet build Tesserae.DocxEditor             # also re-runs the bundle step
```

`bundle.mjs` resolves prosemirror + the doc libs from `packages/core/node_modules`; the root workspace install (with the prosemirror version pins in the root `package.json` `overrides`/`resolutions`) must populate it. A single ProseMirror instance is required — the schema built by core must match the `EditorState` the controller creates.

## C# ↔ JS bridge

The C# components drive the bundled JS through H5 `Script.Write(...)`. `DocxEditor` lazily constructs `new globalThis.docxeditor.DocxEditorController(host, opts)` on mount (`DomObserver.WhenMounted`) and forwards toolbar/formatting calls to it; `DocxDocument` wraps the headless `DocumentAgent`. Keep the C# method surface in sync with the controller/agent methods in `src/js/index.ts` and `packages/core/src/agent/`.

## Security — untrusted DOCX input

A `.docx` is attacker-controlled zip+XML. The core sanitizes at the parse/trust boundary (see `packages/core/src/docx/*Parser.ts`, `utils/sanitizeHref.ts`). When touching parsing/rendering, preserve: no HTML-from-strings, hrefs through `sanitizeHref`, escaped CSS/XML, no DTD/external-entity resolution (XXE), zip-bomb/size caps, no zero-click external fetch (`TargetMode="External"`), recursion/element-count limits, and inert field codes / OLE.

## License

Wrapper: MIT © Curiosity GmbH. Bundled core: Apache-2.0. See `LICENSE`.

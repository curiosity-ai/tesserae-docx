/**
 * Bundles the framework-agnostic docx-editor core (plus a vanilla WYSIWYG editor
 * controller, see ../src/js/index.ts) into a single, self-contained, browser-global
 * JavaScript file that the H5/Tesserae DocxEditor wrapper embeds as a resource.
 *
 * Two artifacts are produced under Tesserae.DocxEditor/assets/js/:
 *   - docx-editor.js      non-minified IIFE (used by H5 Debug builds)
 *   - docx-editor.min.js  minified IIFE      (used by H5 Release builds)
 *
 * Both expose the library on `window.docxeditor` (e.g. `window.docxeditor.DocxEditorController`,
 * `window.docxeditor.DocumentAgent`). H5 picks the ".min.js" variant for Release and the
 * non-".min.js" variant for Debug, so the h5.json references both.
 *
 * No JavaScript UI framework (React/Vue) is bundled — only the ProseMirror editor engine
 * and pure-JS document libraries (jszip, pizzip, xml-js, docxtemplater, dompurify).
 *
 * Run with:  npm run bundle   (from the Tesserae.DocxEditor/ folder)
 */
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const entry = resolve(here, '../src/js/index.ts');
const outDir = resolve(here, '../assets/js');
// The editor's peer libs (prosemirror-*) are installed under the core package's
// node_modules. Point esbuild there so the entry's bare imports resolve to the
// SAME copy the core sources use — a single ProseMirror instance is required so
// the schema built by core matches the EditorState we create here.
const coreNodeModules = resolve(here, '../../packages/core/node_modules');

const banner = {
  js: '/* docx-editor — bundled for the Tesserae.DocxEditor wrapper. Exposed as window.docxeditor. */',
};

/** @type {import('esbuild').BuildOptions} */
const common = {
  entryPoints: [entry],
  bundle: true,
  format: 'iife',
  globalName: 'docxeditor',
  target: 'es2020',
  banner,
  legalComments: 'none',
  logLevel: 'info',
  // Resolve the entry's bare prosemirror-* imports from the core package's node_modules.
  nodePaths: [coreNodeModules],
  // The DOCX libs assume a browser; keep them as-is (no node shims needed at runtime).
  define: { 'process.env.NODE_ENV': '"production"' },
};

await build({ ...common, outfile: resolve(outDir, 'docx-editor.js'), minify: false });
await build({ ...common, outfile: resolve(outDir, 'docx-editor.min.js'), minify: true });

console.log('docx-editor bundled -> Tesserae.DocxEditor/assets/js/docx-editor.js (+ .min.js)');

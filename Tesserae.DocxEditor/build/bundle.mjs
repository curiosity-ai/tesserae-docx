/**
 * Bundles the framework-agnostic docx-editor core (plus a vanilla WYSIWYG editor
 * controller, see ../src/js/index.ts) into browser-global JavaScript files that the
 * H5/Tesserae DocxEditor wrapper embeds as resources.
 *
 * The third-party document libraries are kept in a SEPARATE vendor bundle so the docx
 * bundle itself stays lean and the heavy, rarely-changing libs cache independently:
 *
 *   assets/js/docx-editor-deps.js   jszip + pizzip + xml-js + docxtemplater
 *                                   -> window.docxeditordeps  (must load first)
 *   assets/js/docx-editor.js        the editor + headless API
 *                                   -> window.docxeditor      (references the deps global)
 *
 * (+ a minified ".min.js" of each; H5 picks the readable variant for Debug builds and the
 * minified one for Release builds.)
 *
 * dompurify is NOT bundled: Tesserae already ships it as the global `window.DOMPurify`
 * (see Tesserae's SanitizeHTML/Markdown helpers + purify.min.js), so the docx bundle reuses
 * that instead of shipping a second copy.
 *
 * No JavaScript UI framework (React/Vue) is bundled anywhere — only the ProseMirror editor
 * engine and the pure-JS document libraries above.
 *
 * Run with:  npm run bundle   (from the Tesserae.DocxEditor/ folder)
 */
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const entry = resolve(here, '../src/js/index.ts');
const depsEntry = resolve(here, 'deps-entry.mjs');
const outDir = resolve(here, '../assets/js');
// The peer libs (prosemirror-* and the document libs) are installed under the core
// package's node_modules. Point esbuild there so bare imports resolve to the SAME copy
// the core sources use — a single ProseMirror instance is required so the schema built by
// core matches the EditorState we create here.
const coreNodeModules = resolve(here, '../../packages/core/node_modules');

// Vendor packages that live in the separate deps bundle (window.docxeditordeps),
// plus dompurify which is provided by Tesserae's global. In the main bundle these are
// resolved to tiny shim modules that read from the globals instead of being inlined.
const DEPS_GLOBAL = 'globalThis.docxeditordeps';
const shimSource = (path) => {
  if (path === 'xml-js') {
    return `const m = ${DEPS_GLOBAL}.xmljs; export const xml2js = m.xml2js; export const js2xml = m.js2xml; export default (m.default || m);`;
  }
  if (path === 'dompurify') {
    // Reuse Tesserae's globally-loaded DOMPurify; fall back to a passthrough so the
    // bundle still loads in a (non-Tesserae) host that lacks it.
    return `export default (globalThis.DOMPurify || function(){ return { sanitize: function(h){ return h; }, addHook: function(){}, removeHook: function(){}, setConfig: function(){} }; });`;
  }
  const name = { jszip: 'JSZip', pizzip: 'PizZip', docxtemplater: 'Docxtemplater' }[path];
  return `export default ${DEPS_GLOBAL}.${name};`;
};

/** esbuild plugin: resolve the vendor packages to global-reading shims (main bundle only). */
const vendorGlobalsPlugin = {
  name: 'vendor-globals',
  setup(b) {
    const filter = /^(jszip|pizzip|docxtemplater|xml-js|dompurify)$/;
    b.onResolve({ filter }, (args) => ({ path: args.path, namespace: 'vendor-shim' }));
    b.onLoad({ filter: /.*/, namespace: 'vendor-shim' }, (args) => ({
      contents: shimSource(args.path),
      loader: 'js',
    }));
  },
};

const banner = {
  js: '/* docx-editor — bundled for the Tesserae.DocxEditor wrapper. Exposed as window.docxeditor (needs docx-editor-deps.js). */',
};
const depsBanner = {
  js: '/* docx-editor vendor libs (jszip, pizzip, xml-js, docxtemplater) for Tesserae.DocxEditor. Exposed as window.docxeditordeps. */',
};

const common = {
  bundle: true,
  format: 'iife',
  target: 'es2020',
  legalComments: 'none',
  logLevel: 'info',
  nodePaths: [coreNodeModules],
  define: { 'process.env.NODE_ENV': '"production"' },
};

// 1) Vendor deps bundle -> window.docxeditordeps
await build({
  ...common,
  entryPoints: [depsEntry],
  banner: depsBanner,
  outfile: resolve(outDir, 'docx-editor-deps.js'),
  minify: false,
});
await build({
  ...common,
  entryPoints: [depsEntry],
  banner: depsBanner,
  outfile: resolve(outDir, 'docx-editor-deps.min.js'),
  minify: true,
});

// 2) Main editor bundle -> window.docxeditor (vendor libs externalized to the globals)
await build({
  ...common,
  entryPoints: [entry],
  globalName: 'docxeditor',
  banner,
  plugins: [vendorGlobalsPlugin],
  outfile: resolve(outDir, 'docx-editor.js'),
  minify: false,
});
await build({
  ...common,
  entryPoints: [entry],
  globalName: 'docxeditor',
  banner,
  plugins: [vendorGlobalsPlugin],
  outfile: resolve(outDir, 'docx-editor.min.js'),
  minify: true,
});

console.log('docx-editor bundled -> assets/js/docx-editor-deps.js + docx-editor.js (+ .min.js)');

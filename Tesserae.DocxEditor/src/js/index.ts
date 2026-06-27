/*
 * Tesserae.DocxEditor — browser-global entry point (window.docxeditor).
 *
 * This bundles the FRAMEWORK-AGNOSTIC parts of @eigenpal/docx-editor-core into a
 * single, dependency-free* IIFE so the H5/Tesserae C# wrapper can drive a real
 * WYSIWYG DOCX editor without React or Vue.
 *
 *   *No JavaScript UI framework is pulled in. The only third-party code bundled is
 *    the editor engine (ProseMirror) and pure-JS document libs (jszip, pizzip,
 *    xml-js, docxtemplater, dompurify) — none of which is a UI framework.
 *
 * Two faces are exposed on `window.docxeditor`:
 *   1. The headless programmatic API (DocumentAgent, parseDocx, serializeDocx,
 *      content-controls, templates, …) — re-exported from core/headless.
 *   2. `DocxEditorController` — a vanilla controller that mounts a live, editable
 *      ProseMirror EditorView on a DOM element, backed by the exact same DOCX
 *      schema / conversions / commands the React & Vue editors use, so documents
 *      round-trip faithfully (open .docx → edit → save .docx).
 */

import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { undo, redo } from 'prosemirror-history';

// Re-export the ProseMirror selection classes so hosts can build programmatic
// selections against `controller.view` (e.g. range-format a span).
export { TextSelection, NodeSelection, AllSelection } from 'prosemirror-state';

import { singletonManager } from '../../../packages/core/src/prosemirror/schema';
import {
  toProseDoc,
  createEmptyDoc,
  fromProseDoc,
} from '../../../packages/core/src/prosemirror/conversion';
import {
  createDocumentStylesPlugin,
  createDocumentContextPlugin,
} from '../../../packages/core/src/prosemirror/plugins';
import { ensureParaIdsInState } from '../../../packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension';
import { extractSelectionState } from '../../../packages/core/src/prosemirror/selectionState';
import * as C from '../../../packages/core/src/prosemirror/commands';

import { parseDocx } from '../../../packages/core/src/docx/parser';
import { repackDocx, createDocx } from '../../../packages/core/src/docx/rezip';
import {
  createEmptyDocument,
  createDocumentWithText,
} from '../../../packages/core/src/utils/createDocument';
import { DocumentAgent } from '../../../packages/core/src/agent/DocumentAgent';

// Re-export the whole headless surface (DocumentAgent, parseDocx, serializeDocx,
// content-controls, templates, units/colors, types-as-runtime, …).
export * from '../../../packages/core/src/headless';

/** Wrapper build version (independent of the core library version). */
export const editorVersion = '1.0.0';

// ---------------------------------------------------------------------------
// Styles — injected once. We embed ProseMirror's tiny base stylesheets (so we
// don't ship a separate .css resource) plus a Word-like "page on a desk" look.
// ---------------------------------------------------------------------------

const EDITOR_CSS = `
.tss-docx-editor { box-sizing: border-box; height: 100%; width: 100%; overflow: auto; background: #f3f3f4; }
.tss-docx-editor .ProseMirror {
  background: #ffffff;
  max-width: 816px;
  margin: 24px auto;
  padding: 72px 72px;
  min-height: 1056px;
  box-shadow: 0 1px 8px rgba(0,0,0,0.18);
  outline: none;
  color: #111;
  font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
  font-size: 11pt;
  line-height: 1.4;
  cursor: text;
}
.tss-docx-editor.tss-docx-editor--readonly .ProseMirror { cursor: default; }
.tss-docx-editor .ProseMirror p { margin: 0 0 8px 0; }
.tss-docx-editor .ProseMirror table { border-collapse: collapse; }
.tss-docx-editor .ProseMirror td, .tss-docx-editor .ProseMirror th { border: 1px solid #bbb; padding: 4px 6px; vertical-align: top; }
/* ProseMirror base (vendored from prosemirror-view/style/prosemirror.css) */
.ProseMirror { position: relative; word-wrap: break-word; white-space: pre-wrap; white-space: break-spaces; -webkit-font-variant-ligatures: none; font-variant-ligatures: none; font-feature-settings: "liga" 0; }
.ProseMirror pre { white-space: pre-wrap; }
.ProseMirror li { position: relative; }
.ProseMirror-hideselection *::selection { background: transparent; }
.ProseMirror-hideselection *::-moz-selection { background: transparent; }
.ProseMirror-hideselection { caret-color: transparent; }
.ProseMirror [draggable][contenteditable=false] { user-select: text; }
.ProseMirror-selectednode { outline: 2px solid #8cf; }
li.ProseMirror-selectednode { outline: none; }
li.ProseMirror-selectednode:after { content: ""; position: absolute; left: -32px; right: -2px; top: -2px; bottom: -2px; border: 2px solid #8cf; pointer-events: none; }
/* ProseMirror gapcursor (vendored from prosemirror-gapcursor/style/gapcursor.css) */
.ProseMirror-gapcursor { display: none; pointer-events: none; position: absolute; }
.ProseMirror-gapcursor:after { content: ""; display: block; position: absolute; top: -2px; width: 20px; border-top: 1px solid black; animation: ProseMirror-cursor-blink 1.1s steps(2, start) infinite; }
@keyframes ProseMirror-cursor-blink { to { visibility: hidden; } }
.ProseMirror-focused .ProseMirror-gapcursor { display: block; }
/* toolbar active-state marker toggled by the C# wrapper */
.tss-tb-on, .tss-tb-on > * { background: #cfe3ff !important; }
`;

let stylesInjected = false;
function injectStyles(doc: any): void {
  if (stylesInjected) return;
  const d = doc || (typeof document !== 'undefined' ? document : null);
  if (!d || !d.head) return;
  const el = d.createElement('style');
  el.id = 'tss-docx-editor-styles';
  el.textContent = EDITOR_CSS;
  d.head.appendChild(el);
  stylesInjected = true;
}

// ---------------------------------------------------------------------------
// Controller
// ---------------------------------------------------------------------------

export interface DocxEditorControllerOptions {
  /** Pre-parsed Document to load on mount. */
  document?: any;
  /** Whether editing is disabled. */
  readOnly?: boolean;
  /** Invoked (no args) after every transaction that changes the document. */
  onChange?: () => void;
  /** Invoked (no args) after every transaction (selection or doc) — use to refresh toolbar state. */
  onStateChange?: () => void;
}

/**
 * Drives a single live editor mounted on a host element. Construct it, then
 * call `loadBuffer` / `newDocument` to populate it, run formatting commands via
 * the convenience methods, and `save()` to get DOCX bytes back.
 */
export class DocxEditorController {
  host: any;
  options: DocxEditorControllerOptions;
  view: any = null;
  baseDocument: any = null;

  constructor(host: any, options?: DocxEditorControllerOptions) {
    this.host = host;
    this.options = options || {};
    injectStyles(host && host.ownerDocument);
    this.baseDocument = this.options.document || null;
    this._mount(this.baseDocument);
  }

  // ---- lifecycle --------------------------------------------------------

  private _createState(documentModel: any): any {
    const styles =
      (documentModel && documentModel.package && documentModel.package.styles) || null;
    const pmDoc = documentModel ? toProseDoc(documentModel, { styles }) : createEmptyDoc();
    const theme =
      (documentModel && documentModel.package && documentModel.package.theme) || null;
    const defaultTableStyleId =
      (documentModel &&
        documentModel.package &&
        documentModel.package.settings &&
        documentModel.package.settings.defaultTableStyle) ||
      null;
    const plugins = [
      ...singletonManager.getPlugins(),
      createDocumentStylesPlugin(styles),
      createDocumentContextPlugin({ theme, defaultTableStyleId }),
    ];
    return ensureParaIdsInState(
      EditorState.create({ doc: pmDoc, schema: singletonManager.getSchema(), plugins })
    );
  }

  private _mount(documentModel: any): void {
    this.host.classList.add('tss-docx-editor');
    if (this.options.readOnly) this.host.classList.add('tss-docx-editor--readonly');
    const state = this._createState(documentModel);
    const self = this;
    this.view = new EditorView(this.host, {
      state,
      editable: () => !self.options.readOnly,
      dispatchTransaction(this: any, tr: any) {
        const newState = this.state.apply(tr);
        this.updateState(newState);
        if (tr.docChanged && self.options.onChange) {
          try {
            self.options.onChange();
          } catch (e) {
            /* host callback errors must not break editing */
          }
        }
        self._emitState();
      },
    });
    this._emitState();
  }

  private _emitState(): void {
    if (this.options.onStateChange) {
      try {
        this.options.onStateChange();
      } catch (e) {
        /* ignore host callback errors */
      }
    }
  }

  /** Replace the edited document. */
  setDocument(documentModel: any): void {
    this.baseDocument = documentModel;
    if (!this.view) {
      this._mount(documentModel);
      return;
    }
    this.view.updateState(this._createState(documentModel));
    this._emitState();
  }

  /** Parse DOCX bytes and load them. Returns a promise that resolves when loaded. */
  async loadBuffer(buffer: ArrayBuffer): Promise<boolean> {
    const doc = await parseDocx(buffer);
    this.setDocument(doc);
    return true;
  }

  /** Start a fresh document, optionally seeded with a line of text. */
  newDocument(text?: string): void {
    const doc = text && text.length ? createDocumentWithText(text) : createEmptyDocument();
    this.setDocument(doc);
  }

  // ---- command execution ------------------------------------------------

  private _run(command: any): boolean {
    if (!this.view) return false;
    const ok = command(this.view.state, this.view.dispatch, this.view);
    this.view.focus();
    return !!ok;
  }

  /** Run any ProseMirror Command object directly (advanced / escape hatch). */
  exec(command: any): boolean {
    return this._run(command);
  }

  toggleBold(): boolean {
    return this._run(C.toggleBold);
  }
  toggleItalic(): boolean {
    return this._run(C.toggleItalic);
  }
  toggleUnderline(): boolean {
    return this._run(C.toggleUnderline);
  }
  toggleStrike(): boolean {
    return this._run(C.toggleStrike);
  }
  toggleSuperscript(): boolean {
    return this._run(C.toggleSuperscript);
  }
  toggleSubscript(): boolean {
    return this._run(C.toggleSubscript);
  }

  setAlignment(alignment: string): boolean {
    return this._run(C.setAlignment(alignment as any));
  }
  setFontSize(size: number): boolean {
    return this._run(C.setFontSize(size));
  }
  setFontFamily(name: string): boolean {
    return this._run(C.setFontFamily(name));
  }
  setTextColor(rgbHex: string): boolean {
    return this._run(C.setTextColor({ rgb: rgbHex }));
  }
  clearTextColor(): boolean {
    return this._run(C.clearTextColor);
  }
  setHighlight(color: string): boolean {
    return this._run(C.setHighlight(color));
  }
  clearHighlight(): boolean {
    return this._run(C.clearHighlight);
  }
  setLineSpacing(value: number): boolean {
    return this._run(C.setLineSpacing(value));
  }
  increaseIndent(): boolean {
    return this._run(C.increaseIndent());
  }
  decreaseIndent(): boolean {
    return this._run(C.decreaseIndent());
  }

  toggleBulletList(): boolean {
    return this._run(C.toggleBulletList);
  }
  toggleNumberedList(): boolean {
    return this._run(C.toggleNumberedList);
  }

  /** Apply a paragraph style by id (e.g. "Heading1", "Normal", "Title"). */
  applyStyle(styleId: string): boolean {
    return this._run(C.applyStyle(styleId));
  }
  clearStyle(): boolean {
    return this._run(C.clearStyle);
  }

  insertTable(rows: number, cols: number): boolean {
    return this._run(C.insertTable(rows, cols));
  }
  insertPageBreak(): boolean {
    return this._run(C.insertPageBreak);
  }

  undo(): boolean {
    return this._run(undo);
  }
  redo(): boolean {
    return this._run(redo);
  }
  canUndo(): boolean {
    return this.view ? !!undo(this.view.state) : false;
  }
  canRedo(): boolean {
    return this.view ? !!redo(this.view.state) : false;
  }

  // ---- toolbar / query --------------------------------------------------

  private _sel(): any {
    return this.view ? extractSelectionState(this.view.state) : null;
  }

  isBold(): boolean {
    const s = this._sel();
    return !!(s && s.textFormatting && s.textFormatting.bold);
  }
  isItalic(): boolean {
    const s = this._sel();
    return !!(s && s.textFormatting && s.textFormatting.italic);
  }
  isUnderline(): boolean {
    const s = this._sel();
    return !!(s && s.textFormatting && s.textFormatting.underline);
  }
  isStrike(): boolean {
    const s = this._sel();
    return !!(s && s.textFormatting && s.textFormatting.strike);
  }
  getAlignment(): string {
    const s = this._sel();
    return (s && s.paragraphFormatting && s.paragraphFormatting.alignment) || 'left';
  }
  getStyleId(): string {
    const s = this._sel();
    return (s && s.styleId) || '';
  }
  getFontSize(): number {
    const s = this._sel();
    const v = s && s.textFormatting && s.textFormatting.fontSize;
    return typeof v === 'number' ? v : 0;
  }

  // ---- read / export ----------------------------------------------------

  private _currentDocument(): any {
    if (!this.view) return this.baseDocument;
    return fromProseDoc(this.view.state.doc, this.baseDocument);
  }

  /** Current document as a Document model (PM → Document). */
  getDocument(): any {
    return this._currentDocument();
  }

  getText(): string {
    const doc = this._currentDocument();
    if (!doc) return '';
    return DocumentAgent.fromDocument(doc).getText();
  }
  getWordCount(): number {
    const doc = this._currentDocument();
    if (!doc) return 0;
    return DocumentAgent.fromDocument(doc).getWordCount();
  }

  /** Serialize the current document to DOCX bytes (ArrayBuffer). */
  async save(): Promise<ArrayBuffer> {
    const doc = this._currentDocument();
    if (doc && doc.originalBuffer) {
      const repacked = await repackDocx(doc);
      doc.originalBuffer = repacked;
      this.baseDocument = doc;
      return repacked;
    }
    return createDocx(doc);
  }

  /** Serialize to a DOCX Blob (handy for browser downloads). */
  async saveBlob(): Promise<Blob> {
    const buf = await this.save();
    return new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  // ---- misc -------------------------------------------------------------

  setReadOnly(readOnly: boolean): void {
    this.options.readOnly = readOnly;
    if (this.host) this.host.classList.toggle('tss-docx-editor--readonly', !!readOnly);
    if (this.view) this.view.updateState(this.view.state); // re-evaluate editable()
  }

  focus(): void {
    if (this.view) this.view.focus();
  }

  destroy(): void {
    if (this.view) {
      this.view.destroy();
      this.view = null;
    }
  }
}

/** Convenience factory mirroring the headless `createAgent` style. */
export function createEditor(
  host: any,
  options?: DocxEditorControllerOptions
): DocxEditorController {
  return new DocxEditorController(host, options);
}

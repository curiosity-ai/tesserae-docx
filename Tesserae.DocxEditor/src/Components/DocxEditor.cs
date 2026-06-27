using System;
using System.Collections.Generic;
using H5;
using Tesserae;
using static H5.Core.dom;
using static Tesserae.UI;

namespace Tesserae.DocxEditor
{
    /// <summary>
    /// A Tesserae component that embeds a live, editable WYSIWYG DOCX editor built on the
    /// framework-agnostic core of the
    /// <see href="https://github.com/eigenpal/docx-editor">docx-editor</see> library — with
    /// <b>no JavaScript UI framework</b> (no React/Vue). The editor engine (ProseMirror) and the
    /// DOCX parser/serializer are bundled with this package as a browser-global script
    /// (<c>window.docxeditor</c>) and embedded via <c>h5.json</c>, so there is no preload step.
    ///
    /// The toolbar is composed entirely from Tesserae components; toolbar buttons drive the
    /// underlying editor through the bundled controller. The editor instance is created lazily
    /// once the component is mounted into the DOM.
    /// </summary>
    [H5.Name("docx.DocxEditor")]
    public class DocxEditor : IComponent
    {
        // Host element the ProseMirror EditorView mounts into.
        private readonly HTMLElement _host;

        // The live `DocxEditorController` JS instance (null until mounted).
        private object _ctrl;

        // Configuration captured before the controller exists.
        private bool _showToolbar = true;
        private bool _readOnly = false;
        private string _initialText; // seed an empty doc with a line of text
        private object _initialBuffer; // ArrayBuffer to load on mount

        private Action _onChange;
        private Action _onReady;

        // Toolbar element refs for active-state reflection.
        private HTMLElement _btnBold, _btnItalic, _btnUnderline, _btnStrike;
        private HTMLElement _btnLeft, _btnCenter, _btnRight, _btnJustify;
        private TextBlock _status;

        private HTMLElement _rendered;

        /// <summary>Gets the underlying <c>DocxEditorController</c> JS instance (or <c>null</c> before mount).</summary>
        public object Controller => _ctrl;

        /// <summary>Initializes a new, empty editor that fills its parent.</summary>
        public DocxEditor()
        {
            _host = Div(_("tss-docx-editor-host"));
            _host.style.flexGrow = "1";
            _host.style.minHeight = "0";
            _host.style.width = "100%";
            DomObserver.WhenMounted(_host, OnMounted);
        }

        // ── Configuration (call before the editor is mounted) ──────────────────────

        /// <summary>Hides the built-in Tesserae toolbar (on by default).</summary>
        public DocxEditor NoToolbar() { _showToolbar = false; return this; }

        /// <summary>Makes the editor read-only.</summary>
        public DocxEditor ReadOnly(bool readOnly = true)
        {
            _readOnly = readOnly;
            if (_ctrl != null) Script.Write("{0}.setReadOnly({1})", _ctrl, readOnly);
            return this;
        }

        /// <summary>Seeds a new, empty document with a single line of <paramref name="text"/> on mount.</summary>
        public DocxEditor WithText(string text) { _initialText = text; return this; }

        /// <summary>Loads a DOCX file (an <c>ArrayBuffer</c>) on mount (or immediately if already mounted).</summary>
        public DocxEditor WithDocx(object arrayBuffer)
        {
            if (_ctrl != null) Load(arrayBuffer);
            else _initialBuffer = arrayBuffer;
            return this;
        }

        /// <summary>Registers a callback invoked whenever the document content changes.</summary>
        public DocxEditor OnChange(Action handler) { _onChange = handler; return this; }

        /// <summary>Registers a callback invoked once the editor has mounted and is ready.</summary>
        public DocxEditor OnReady(Action handler)
        {
            _onReady = handler;
            if (_ctrl != null) handler?.Invoke();
            return this;
        }

        // ── Document I/O ───────────────────────────────────────────────────────────

        /// <summary>Loads a DOCX file (an <c>ArrayBuffer</c>) into the editor.</summary>
        public DocxEditor Load(object arrayBuffer, Action onLoaded = null)
        {
            if (_ctrl == null) { _initialBuffer = arrayBuffer; return this; }
            Action done = () => onLoaded?.Invoke();
            Script.Write(
                "{0}.loadBuffer({1}).then(function(){ {2}(); }).catch(function(e){ console.error('docx load failed', e); })",
                _ctrl, arrayBuffer, done);
            return this;
        }

        /// <summary>Starts a fresh, empty document (optionally seeded with a line of text).</summary>
        public DocxEditor NewDocument(string text = null)
        {
            if (_ctrl == null) { _initialText = text; _initialBuffer = null; return this; }
            if (string.IsNullOrEmpty(text)) Script.Write("{0}.newDocument()", _ctrl);
            else Script.Write("{0}.newDocument({1})", _ctrl, text);
            return this;
        }

        /// <summary>Serializes the current document to DOCX bytes and hands the <c>ArrayBuffer</c> to <paramref name="onSaved"/>.</summary>
        public DocxEditor Save(Action<object> onSaved)
        {
            if (_ctrl == null) return this;
            Action<object> cb = buf => onSaved?.Invoke(buf);
            Script.Write(
                "{0}.save().then(function(buf){ {1}(buf); }).catch(function(e){ console.error('docx save failed', e); })",
                _ctrl, cb);
            return this;
        }

        /// <summary>Serializes the current document and triggers a browser download as <paramref name="filename"/>.</summary>
        public DocxEditor Download(string filename = "document.docx")
        {
            if (_ctrl == null) return this;
            Script.Write(@"{0}.saveBlob().then(function(blob){
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url; a.download = {1};
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
            }).catch(function(e){ console.error('docx download failed', e); })", _ctrl, filename);
            return this;
        }

        /// <summary>Returns the plain-text content of the document (empty before mount).</summary>
        public string GetText() => _ctrl == null ? "" : Script.Write<string>("{0}.getText()", _ctrl);

        /// <summary>Returns the document word count (0 before mount).</summary>
        public int GetWordCount() => _ctrl == null ? 0 : Script.Write<int>("{0}.getWordCount()", _ctrl);

        // ── Formatting commands ────────────────────────────────────────────────────

        public DocxEditor ToggleBold() { Run("toggleBold"); return this; }
        public DocxEditor ToggleItalic() { Run("toggleItalic"); return this; }
        public DocxEditor ToggleUnderline() { Run("toggleUnderline"); return this; }
        public DocxEditor ToggleStrikethrough() { Run("toggleStrike"); return this; }
        public DocxEditor ToggleBulletList() { Run("toggleBulletList"); return this; }
        public DocxEditor ToggleNumberedList() { Run("toggleNumberedList"); return this; }
        public DocxEditor IncreaseIndent() { Run("increaseIndent"); return this; }
        public DocxEditor DecreaseIndent() { Run("decreaseIndent"); return this; }
        public DocxEditor InsertPageBreak() { Run("insertPageBreak"); return this; }
        public DocxEditor Undo() { Run("undo"); return this; }
        public DocxEditor Redo() { Run("redo"); return this; }

        /// <summary>Sets paragraph alignment: <c>"left"</c>, <c>"center"</c>, <c>"right"</c>, or <c>"justify"</c>.</summary>
        public DocxEditor SetAlignment(string alignment) { if (_ctrl != null) Script.Write("{0}.setAlignment({1})", _ctrl, alignment); return this; }
        /// <summary>Applies a paragraph style by id (e.g. <c>"Heading1"</c>, <c>"Title"</c>, <c>"Normal"</c>).</summary>
        public DocxEditor ApplyStyle(string styleId) { if (_ctrl != null) Script.Write("{0}.applyStyle({1})", _ctrl, styleId); return this; }
        /// <summary>Sets the font size in points.</summary>
        public DocxEditor SetFontSize(double points) { if (_ctrl != null) Script.Write("{0}.setFontSize({1})", _ctrl, points); return this; }
        /// <summary>Sets the font family.</summary>
        public DocxEditor SetFontFamily(string name) { if (_ctrl != null) Script.Write("{0}.setFontFamily({1})", _ctrl, name); return this; }
        /// <summary>Sets the text color (RGB hex without the leading '#', e.g. <c>"FF0000"</c>).</summary>
        public DocxEditor SetTextColor(string rgbHex) { if (_ctrl != null) Script.Write("{0}.setTextColor({1})", _ctrl, rgbHex); return this; }
        /// <summary>Highlights the selection with a color name (e.g. <c>"yellow"</c>).</summary>
        public DocxEditor SetHighlight(string color) { if (_ctrl != null) Script.Write("{0}.setHighlight({1})", _ctrl, color); return this; }
        /// <summary>Inserts a table with the given dimensions at the cursor.</summary>
        public DocxEditor InsertTable(int rows, int cols) { if (_ctrl != null) Script.Write("{0}.insertTable({1}, {2})", _ctrl, rows, cols); return this; }

        /// <summary>Focuses the editor.</summary>
        public DocxEditor Focus() { if (_ctrl != null) Script.Write("{0}.focus()", _ctrl); return this; }

        private void Run(string method)
        {
            if (_ctrl != null) Script.Write("{0}[{1}]()", _ctrl, method);
        }

        // ── Mount / internals ──────────────────────────────────────────────────────

        private void OnMounted()
        {
            try
            {
                var opts = Script.Write<object>("{}");
                if (_readOnly) Script.Write("{0}.readOnly = true", opts);

                Action onState = RefreshToolbar;
                Action onChange = () => { _onChange?.Invoke(); };
                Script.Write("{0}.onStateChange = function(){ {1}(); }", opts, onState);
                Script.Write("{0}.onChange = function(){ {1}(); }", opts, onChange);

                _ctrl = Script.Write<object>("new globalThis.docxeditor.DocxEditorController({0}, {1})", _host, opts);

                if (_initialBuffer != null) Load(_initialBuffer);
                else if (!string.IsNullOrEmpty(_initialText)) NewDocument(_initialText);

                RefreshToolbar();
                _onReady?.Invoke();

                DomObserver.WhenRemoved(_host, () =>
                {
                    if (_ctrl != null) { Script.Write("{0}.destroy()", _ctrl); _ctrl = null; }
                });
            }
            catch (Exception ex)
            {
                console.error("Tesserae.DocxEditor: failed to initialize editor", ex);
            }
        }

        private IComponent BuildToolbar()
        {
            var bold = Button("B").SetTitle("Bold (Ctrl+B)").OnClick(() => ToggleBold());
            var italic = Button("I").SetTitle("Italic (Ctrl+I)").OnClick(() => ToggleItalic());
            var underline = Button("U").SetTitle("Underline (Ctrl+U)").OnClick(() => ToggleUnderline());
            var strike = Button("S").SetTitle("Strikethrough").OnClick(() => ToggleStrikethrough());

            var left = Button("Left").SetTitle("Align left").OnClick(() => SetAlignment("left"));
            var center = Button("Center").SetTitle("Align center").OnClick(() => SetAlignment("center"));
            var right = Button("Right").SetTitle("Align right").OnClick(() => SetAlignment("right"));
            var justify = Button("Justify").SetTitle("Justify").OnClick(() => SetAlignment("justify"));

            _btnBold = bold.Render();
            _btnItalic = italic.Render();
            _btnUnderline = underline.Render();
            _btnStrike = strike.Render();
            _btnLeft = left.Render();
            _btnCenter = center.Render();
            _btnRight = right.Render();
            _btnJustify = justify.Render();

            var style = Dropdown().Width(150.px()).Items(
                DropdownItem("Normal").Selected().OnSelected(_ => ApplyStyle("Normal")),
                DropdownItem("Title").OnSelected(_ => ApplyStyle("Title")),
                DropdownItem("Heading 1").OnSelected(_ => ApplyStyle("Heading1")),
                DropdownItem("Heading 2").OnSelected(_ => ApplyStyle("Heading2")),
                DropdownItem("Heading 3").OnSelected(_ => ApplyStyle("Heading3"))
            );

            _status = TextBlock("0 words").Secondary();

            return VStack().WS().Children(
                HStack().AlignItems(ItemAlign.Center).WS().Wrap().Children(
                    bold, italic, underline, strike,
                    Divider(),
                    left, center, right, justify,
                    Divider(),
                    Button("• List").SetTitle("Bullet list").OnClick(() => ToggleBulletList()),
                    Button("1. List").SetTitle("Numbered list").OnClick(() => ToggleNumberedList()),
                    Button("Indent +").OnClick(() => IncreaseIndent()),
                    Button("Indent -").OnClick(() => DecreaseIndent()),
                    Divider(),
                    style,
                    Button("Table").SetTitle("Insert 3x3 table").OnClick(() => InsertTable(3, 3)),
                    Button("Page break").OnClick(() => InsertPageBreak()),
                    Divider(),
                    Button("Undo").OnClick(() => Undo()),
                    Button("Redo").OnClick(() => Redo()),
                    _status
                )
            ).Padding(8.px());
        }

        private static IComponent Divider()
        {
            var d = Div(_());
            d.style.width = "1px";
            d.style.alignSelf = "stretch";
            d.style.background = "rgba(0,0,0,0.15)";
            d.style.margin = "0 4px";
            return Raw(d);
        }

        private void RefreshToolbar()
        {
            if (_ctrl == null) return;
            SetActive(_btnBold, Script.Write<bool>("{0}.isBold()", _ctrl));
            SetActive(_btnItalic, Script.Write<bool>("{0}.isItalic()", _ctrl));
            SetActive(_btnUnderline, Script.Write<bool>("{0}.isUnderline()", _ctrl));
            SetActive(_btnStrike, Script.Write<bool>("{0}.isStrike()", _ctrl));

            var align = Script.Write<string>("{0}.getAlignment()", _ctrl);
            SetActive(_btnLeft, align == "left" || string.IsNullOrEmpty(align));
            SetActive(_btnCenter, align == "center");
            SetActive(_btnRight, align == "right");
            SetActive(_btnJustify, align == "justify" || align == "both");

            if (_status != null) _status.Text = $"{Script.Write<int>("{0}.getWordCount()", _ctrl)} words";
        }

        private static void SetActive(HTMLElement el, bool active)
        {
            if (el == null) return;
            if (active) el.classList.add("tss-tb-on");
            else el.classList.remove("tss-tb-on");
        }

        /// <summary>Renders the component's root element.</summary>
        public HTMLElement Render()
        {
            if (_rendered != null) return _rendered;
            var children = new List<IComponent>();
            if (_showToolbar) children.Add(BuildToolbar());
            children.Add(Raw(_host).Grow());
            _rendered = Stack().S().Children(children.ToArray()).Render();
            return _rendered;
        }
    }
}

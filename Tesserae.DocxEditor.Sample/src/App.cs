using System;
using System.Collections.Generic;
using H5;
using Tesserae;
using Tesserae.DocxEditor;
using static H5.Core.dom;
using static Tesserae.UI;

namespace Tesserae.DocxEditor.Sample
{
    /// <summary>
    /// Demo of the <see cref="DocxEditor"/> Tesserae.DocxEditor component: a live WYSIWYG DOCX
    /// editor (no JS UI framework) with a Tesserae-built toolbar, plus file open/save/download and
    /// a small showcase of the headless <see cref="DocxDocument"/> API. Mirrors the spirit of the
    /// browser samples to keep feature coverage between the library and the C# wrapper in sync.
    /// </summary>
    internal static class App
    {
        private static TextBlock _status;
        private static DocxEditor _editor;
        private static bool _readOnly;

        private static void Main()
        {
            document.body.style.overflow = "hidden";

            _status = TextBlock("Ready — type in the document, use the toolbar, then Download .docx.").Secondary();

            _editor = DOCX.Editor()
                .WithText("Welcome to the Tesserae DOCX editor.\n\nThis is a real WYSIWYG editor built in pure JavaScript (no React/Vue) and driven from C#. Select some text and try Bold, headings, alignment, lists and tables in the toolbar above — then open or download a .docx file.")
                .OnChange(() => _status.Text = $"Editing — {_editor.GetWordCount()} words");

            var appbar = HStack().AlignItems(ItemAlign.Center).WS().Wrap().Children(
                TextBlock("Tesserae.DocxEditor").Bold(),
                Button("New").SetTitle("Start a new empty document").OnClick(() =>
                {
                    _editor.NewDocument("New document.");
                    _status.Text = "Started a new document.";
                }),
                Button("Open .docx…").SetTitle("Open a .docx file from disk").OnClick(OpenFile),
                Button("Download .docx").SetTitle("Save the current document to disk").OnClick(() =>
                {
                    _editor.Download("document.docx");
                    _status.Text = "Downloaded document.docx";
                }),
                Button("Toggle read-only").OnClick(() =>
                {
                    _readOnly = !_readOnly;
                    _editor.ReadOnly(_readOnly);
                    _status.Text = _readOnly ? "Read-only mode." : "Editing mode.";
                }),
                Button("Headless demo").SetTitle("Build, inspect and download a document with no UI").OnClick(HeadlessDemo),
                _status
            ).Padding(8.px());

            var layout = Stack().S().Children(
                appbar,
                _editor.Grow()
            );

            document.body.appendChild(layout.Render());
        }

        /// <summary>Opens a native file picker, reads the chosen .docx and loads it into the editor.</summary>
        private static void OpenFile()
        {
            Action<object> onBuf = buffer =>
            {
                _editor.Load(buffer, () => _status.Text = $"Loaded .docx — {_editor.GetWordCount()} words");
            };

            // Use a transient <input type=file> + File.arrayBuffer() to obtain an ArrayBuffer,
            // then hand it back to C#. (Done in JS to avoid framework-specific DOM-typing here.)
            Script.Write(@"(function(cb){
                var input = document.createElement('input');
                input.type = 'file';
                input.accept = '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
                input.onchange = function(){
                    var f = input.files && input.files[0];
                    if (!f) return;
                    f.arrayBuffer().then(function(buf){ cb(buf); });
                };
                input.click();
            })({0})", onBuf);
        }

        /// <summary>Demonstrates the headless <see cref="DocxDocument"/> API (no visual editor).</summary>
        private static void HeadlessDemo()
        {
            var doc = DOCX.NewDocument("Invoice for {customer}, dated {date}. Total: {amount}.");
            var variables = doc.GetVariables();
            var words = doc.GetWordCount();

            doc.InsertText(0, 0, "DRAFT — ").ApplyStyle(0, "Heading1");

            _status.Text = $"Headless: {words} words, {variables.Length} variables [{string.Join(", ", variables)}] — downloading headless-demo.docx";
            doc.Download("headless-demo.docx");
        }
    }
}

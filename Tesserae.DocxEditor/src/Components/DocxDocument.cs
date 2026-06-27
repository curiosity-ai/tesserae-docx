using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using H5;
using static H5.Core.dom;

namespace Tesserae.DocxEditor
{
    /// <summary>
    /// A C# wrapper around the docx-editor <b>headless</b> programmatic API (the
    /// JavaScript <c>DocumentAgent</c>) — for reading, editing and saving DOCX documents
    /// without any visual editor. The underlying JS library is bundled with this package as
    /// <c>window.docxeditor</c>.
    ///
    /// The underlying agent is immutable: each editing method swaps in the new agent it returns
    /// and yields <c>this</c>, so calls chain naturally from C#.
    /// </summary>
    [H5.Name("docx.DocxDocument")]
    public sealed class DocxDocument
    {
        // The live JS DocumentAgent instance.
        private object _agent;

        internal DocxDocument(object agent) { _agent = agent; }

        /// <summary>The underlying JS <c>DocumentAgent</c> for advanced use.</summary>
        public object Agent => _agent;

        // ── Construction ───────────────────────────────────────────────────────────

        /// <summary>Creates a new headless document (optionally seeded with a line of text).</summary>
        public static DocxDocument New(string text = null)
        {
            object doc = string.IsNullOrEmpty(text)
                ? Script.Write<object>("globalThis.docxeditor.createEmptyDocument()")
                : Script.Write<object>("globalThis.docxeditor.createDocumentWithText({0})", text);
            var agent = Script.Write<object>("globalThis.docxeditor.DocumentAgent.fromDocument({0})", doc);
            return new DocxDocument(agent);
        }

        /// <summary>Loads a DOCX file (an <c>ArrayBuffer</c>) into a headless document.</summary>
        public static Task<DocxDocument> LoadAsync(object arrayBuffer)
        {
            var tcs = new TaskCompletionSource<DocxDocument>();
            var promise = Script.Write<object>("globalThis.docxeditor.DocumentAgent.fromBuffer({0})", arrayBuffer);
            Action<object> res = agent => tcs.SetResult(new DocxDocument(agent));
            Action<object> rej = err => tcs.SetException(new Exception("Failed to load DOCX"));
            Script.Write("{0}.then(function(a){ {1}(a); }, function(e){ {2}(e); })", promise, res, rej);
            return tcs.Task;
        }

        // ── Reading ────────────────────────────────────────────────────────────────

        /// <summary>The plain-text content of the document.</summary>
        public string GetText() => Script.Write<string>("{0}.getText()", _agent);

        /// <summary>Word count.</summary>
        public int GetWordCount() => Script.Write<int>("{0}.getWordCount()", _agent);

        /// <summary>Character count (including whitespace by default).</summary>
        public int GetCharacterCount(bool includeSpaces = true) => Script.Write<int>("{0}.getCharacterCount({1})", _agent, includeSpaces);

        /// <summary>Number of top-level paragraphs.</summary>
        public int GetParagraphCount() => Script.Write<int>("{0}.getParagraphCount()", _agent);

        /// <summary>Number of top-level tables.</summary>
        public int GetTableCount() => Script.Write<int>("{0}.getTableCount()", _agent);

        /// <summary>Detected template variable names (without the surrounding braces).</summary>
        public string[] GetVariables() => Script.Write<string[]>("{0}.getVariables()", _agent);

        /// <summary>Template tags found in the document (docxtemplater syntax).</summary>
        public string[] GetTemplateTags() => Script.Write<string[]>("globalThis.docxeditor.getTemplateTags({0}.getDocument())", _agent);

        // ── Editing (immutable: swaps the agent, returns this) ──────────────────────

        /// <summary>Inserts <paramref name="text"/> at a paragraph/offset position.</summary>
        public DocxDocument InsertText(int paragraphIndex, int offset, string text)
        {
            var pos = Script.Write<object>("{ paragraphIndex: {0}, offset: {1} }", paragraphIndex, offset);
            _agent = Script.Write<object>("{0}.insertText({1}, {2})", _agent, pos, text);
            return this;
        }

        /// <summary>Applies a named paragraph style (e.g. <c>"Heading1"</c>) to a paragraph.</summary>
        public DocxDocument ApplyStyle(int paragraphIndex, string styleId)
        {
            _agent = Script.Write<object>("{0}.applyStyle({1}, {2})", _agent, paragraphIndex, styleId);
            return this;
        }

        /// <summary>Inserts a table at a paragraph/offset position.</summary>
        public DocxDocument InsertTable(int paragraphIndex, int offset, int rows, int cols)
        {
            var pos = Script.Write<object>("{ paragraphIndex: {0}, offset: {1} }", paragraphIndex, offset);
            _agent = Script.Write<object>("{0}.insertTable({1}, {2}, {3})", _agent, pos, rows, cols);
            return this;
        }

        // ── Template variables ───────────────────────────────────────────────────────

        /// <summary>Queues a single template variable (applied by <see cref="ApplyVariablesAsync"/>).</summary>
        public DocxDocument SetVariable(string name, string value)
        {
            _agent = Script.Write<object>("{0}.setVariable({1}, {2})", _agent, name, value);
            return this;
        }

        /// <summary>
        /// Applies template variables (docxtemplater substitution) and returns the resulting document.
        /// Requires a document loaded from a real .docx file (it needs the original file bytes).
        /// </summary>
        public Task<DocxDocument> ApplyVariablesAsync(Dictionary<string, string> variables = null)
        {
            var tcs = new TaskCompletionSource<DocxDocument>();
            var js = ToJsObject(variables);
            var promise = Script.Write<object>("{0}.applyVariables({1})", _agent, js);
            Action<object> res = agent => tcs.SetResult(new DocxDocument(agent));
            Action<object> rej = err => tcs.SetException(new Exception("Failed to apply variables"));
            Script.Write("{0}.then(function(a){ {1}(a); }, function(e){ {2}(e); })", promise, res, rej);
            return tcs.Task;
        }

        // ── Export ────────────────────────────────────────────────────────────────

        /// <summary>Serializes the document to DOCX bytes (an <c>ArrayBuffer</c>).</summary>
        public Task<object> ToBytesAsync()
        {
            var tcs = new TaskCompletionSource<object>();
            var promise = Script.Write<object>("{0}.toBuffer()", _agent);
            Action<object> res = buf => tcs.SetResult(buf);
            Action<object> rej = err => tcs.SetException(new Exception("Failed to serialize DOCX"));
            Script.Write("{0}.then(function(b){ {1}(b); }, function(e){ {2}(e); })", promise, res, rej);
            return tcs.Task;
        }

        /// <summary>Serializes the document and triggers a browser download as <paramref name="filename"/>.</summary>
        public void Download(string filename = "document.docx")
        {
            Script.Write(@"{0}.toBlob().then(function(blob){
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url; a.download = {1};
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
            }).catch(function(e){ console.error('docx download failed', e); })", _agent, filename);
        }

        // ── Internals ───────────────────────────────────────────────────────────────

        private static object ToJsObject(Dictionary<string, string> map)
        {
            var o = Script.Write<object>("{}");
            if (map != null)
            {
                foreach (var kv in map) Script.Write("{0}[{1}] = {2}", o, kv.Key, kv.Value);
            }
            return o;
        }
    }
}

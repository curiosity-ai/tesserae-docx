using System.Threading.Tasks;

namespace Tesserae.DocxEditor
{
    /// <summary>
    /// Static entry point / factory for Tesserae.DocxEditor, mirroring Tesserae's <c>UI</c> class.
    /// Use <see cref="Editor()"/> for the visual WYSIWYG component, or the document helpers for the
    /// headless (programmatic) API.
    /// </summary>
    public static class DOCX
    {
        /// <summary>Creates a new, empty WYSIWYG <see cref="DocxEditor"/> component.</summary>
        public static DocxEditor Editor() => new DocxEditor();

        /// <summary>Creates a WYSIWYG <see cref="DocxEditor"/> seeded with a line of text.</summary>
        public static DocxEditor Editor(string text) => new DocxEditor().WithText(text);

        /// <summary>Creates a new headless document (optionally seeded with a line of text).</summary>
        public static DocxDocument NewDocument(string text = null) => DocxDocument.New(text);

        /// <summary>Loads a DOCX file (an <c>ArrayBuffer</c>) into a headless document.</summary>
        public static Task<DocxDocument> LoadAsync(object arrayBuffer) => DocxDocument.LoadAsync(arrayBuffer);
    }
}

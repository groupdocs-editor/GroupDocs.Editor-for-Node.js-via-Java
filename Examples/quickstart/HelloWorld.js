const Constants = require('../Constants');
const { Editor, WordProcessingEditOptions, WordProcessingSaveOptions, WordProcessingFormats } = require('@groupdocs/groupdocs.editor'); // Assuming 'groupdocs-editor' is the name of the package

/**
 * This example demonstrates a basic Hello World example using GroupDocs.Editor.
 */
class HelloWorld {
    static async run() {
        const documentPath = Constants.SAMPLE_DOCX;
        console.log(`try to open document ${documentPath}`);
        try {
            const editor = new Editor(documentPath);

            // Obtain editable document from original DOCX document
            const editOptions = new WordProcessingEditOptions();
            const editableDocument = await editor.edit(editOptions); // Open document for editing

            const htmlContent = editableDocument.getEmbeddedHtml();
            // Pass htmlContent to WYSIWYG editor and edit there...
            // ...

            // Save edited EditableDocument object to some WordProcessing format - DOC for example
            const savePath = Constants.getOutputFilePath("HelloWorldOutput", "docx");
            const saveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Docx);
            await editor.save(editableDocument, savePath, saveOptions);

            console.log(`Document edited and saved successfully, path ${savePath}.`);
        } catch (ex) {
            console.log(ex.message);
        }
    }
}

module.exports = HelloWorld;

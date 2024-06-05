// noinspection JSUnresolvedReference

const {
    Editor, EditableDocument, WordProcessingFormats, WordProcessingLoadOptions,
    WordProcessingSaveOptions, StreamBuffer
} = require('@groupdocs/groupdocs.editor'); // Assuming 'groupdocs-editor' is the name of the package
const Constants = require('../Constants');
const {writeFile} = require("node:fs");

class SaveDocument {
    static async run() {
        try {
            const inputFilePath = Constants.SAMPLE_DOCX;
            const editor = new Editor(inputFilePath, new WordProcessingLoadOptions());
            const defaultWordProcessingDoc = await editor.edit();

            // Modify its content somehow. Because there is no attached WYSIWYG-editor, this code simulates document editing
            let allEmbeddedInsideString = defaultWordProcessingDoc.getEmbeddedHtml();
            let allEmbeddedInsideStringEdited = allEmbeddedInsideString.replace("Subtitle", "Edited subtitle");// Modified version

            // Create new EditableDocument instance from modified HTML content
            const editedDoc = EditableDocument.fromMarkup(allEmbeddedInsideStringEdited, null);

            // Save edited document as RTF, specified through file path
            const outputRtfPath = Constants.getOutputFilePath("editedDoc", "rtf");
            const rtfSaveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Rtf);

            await editor.save(editedDoc, outputRtfPath, rtfSaveOptions);

            // Save edited document as DOCM, specified through stream
            const outputDocmPath = Constants.getOutputFilePath("editedDoc", "docm");
            const docmSaveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Docm);
            const buffer = new StreamBuffer();
            await editor.save(editedDoc, buffer, docmSaveOptions);
            writeFile(outputDocmPath, buffer.toByteArray(), (err) => {
                if (err) {
                    console.error('Error writing the file:', err);
                } else {
                    console.log('File written successfully.');
                }
            });

            buffer.close();

            // Dispose EditableDocument and Editor instances
            editedDoc.dispose();
            defaultWordProcessingDoc.dispose();

            console.log("SaveDocument routine has successfully finished");
        } catch (error) {
            console.log("Error:", error);
        }
    }
}

module.exports = SaveDocument;

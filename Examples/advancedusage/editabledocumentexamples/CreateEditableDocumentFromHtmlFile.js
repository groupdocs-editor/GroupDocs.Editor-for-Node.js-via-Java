const {
    EditableDocument,
    Editor,
    WordProcessingFormats,
    WordProcessingSaveOptions, StreamBuffer
} = require('@groupdocs/groupdocs.editor');
const fs = require('fs');
const path = require('path');
const Constants = require('../../Constants');
const {writeFile} = require("node:fs"); // Adjust the path to your Constants module

class CreateEditableDocumentFromHtmlFile {
    static async run() {
        try {
            const htmlFilePath = Constants.SAMPLE_HTML;

            // Remove file extension and get the base name
            const fileName = Constants.removeExtension(path.basename(htmlFilePath));

            // Create EditableDocument from HTML file
            const htmlContent = fs.readFileSync(htmlFilePath, 'utf8');
            const document = EditableDocument.fromMarkup(htmlContent, null);

            // Create an Editor instance
            const editor = new Editor(htmlFilePath);

            // Define save options for WordProcessing format (DOCX)
            const saveOptions = new WordProcessingSaveOptions();
            saveOptions.format = WordProcessingFormats.Docx;

            // Define save path
            const savePath = Constants.getOutputFilePath(fileName, 'docx');

            // Save the document
            const buffer = new StreamBuffer();
            await editor.save(document, buffer, saveOptions);
            writeFile(savePath, buffer.toByteArray(), (err) => {
                if (err) {
                    console.error('Error writing the file:', err);
                } else {
                    console.log('File written successfully.');
                }
            });

            buffer.close();

            console.log('CreateEditableDocumentFromHtmlFile routine has successfully finished');
        } catch (error) {
            console.error('Error:', error);
        }
    }
}

// Run the process
module.exports = CreateEditableDocumentFromHtmlFile;

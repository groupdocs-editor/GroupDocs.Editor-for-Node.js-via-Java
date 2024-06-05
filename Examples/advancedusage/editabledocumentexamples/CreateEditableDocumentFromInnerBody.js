const {
    EditableDocument,
    Editor,
    WordProcessingFormats,
    WordProcessingSaveOptions
} = require('@groupdocs/groupdocs.editor');
const fs = require('fs');
const Constants = require('../../Constants');  // Assuming you have a Constants module

class CreateEditableDocumentFromInnerBody {
    static async run() {
        try {
            // 1. Get path to the file with HTML BODY inner markup
            const pathToHtmlFile = Constants.SAMPLE_HTML_BODY;

            // 2. Read this markup to String
            const content = fs.readFileSync(pathToHtmlFile, 'utf-8');

            // 3. Get path to the resource folder
            const pathToResourceFolder = Constants.SAMPLE_HTML_BODY_RESOURCES;

            // 4. Initialize EditableDocument
            const inputDoc = EditableDocument.fromMarkupAndResourceFolder(content, pathToResourceFolder);

            // 5. Check obtained document
            console.log(`There should be 2 stylesheets, and actually is ${inputDoc.getCss().length}`);
            console.log(`There should be 2 images, and actually is ${inputDoc.getImages().length}`);

            // 6. Save it to the file
            const outputHtmlFilePath = Constants.getOutputDirectoryPath('_output.html');
            await inputDoc.save(outputHtmlFilePath);

            // 7. Save it to the document
            // 7.1. Get saving format
            const saveFormat = WordProcessingFormats.fromExtension('docm');

            // 7.2. Get saving options
            const saveOptions = new WordProcessingSaveOptions(saveFormat);

            // 7.3. Create instance of Editor class, initialize it with anything
            const editor = new Editor(pathToHtmlFile);

            // 7.4. Prepare output path for the DOCM document
            const outputDocmFilePath = Constants.getOutputFilePath('_output', saveFormat.getExtension());

            // 7.5. Save it
            await editor.save(inputDoc, outputDocmFilePath, saveOptions);

            console.log('CreateEditableDocumentFromInnerBody routine has successfully finished');
        } catch (error) {
            console.error('Error:', error);
        }
    }
}

module.exports = CreateEditableDocumentFromInnerBody;

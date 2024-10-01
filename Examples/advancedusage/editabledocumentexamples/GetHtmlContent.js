const {
    EditableDocument,
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions, readDataFromStream
} = require('@groupdocs/groupdocs.editor');
const fs = require('fs');
const Constants = require('../../Constants');

class GetHtmlContent {
    static async run() {
        try {
            const fileStream = fs.createReadStream(Constants.SAMPLE_DOCX);
            const buffer = await readDataFromStream(fileStream);
            const editor = new Editor(buffer, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            const htmlContent = document.getContent();
            fileStream.close();
            buffer.close();
            editor.dispose();
            console.log("HTML content of the input document (first 200 chars): ", htmlContent.substring(0, 200));
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = GetHtmlContent;

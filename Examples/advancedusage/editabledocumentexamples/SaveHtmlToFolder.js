const {
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');
const path = require('path');

class SaveHtmlToFolder {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            const fileName = Constants.removeExtension(path.basename(Constants.SAMPLE_DOCX));
            const outputHtml = Constants.getOutputFilePath(fileName, 'html');
            await document.save(outputHtml);
            console.log(`HTML content saved to: ${outputHtml}`);
        } catch (error) {
            console.error('Error:', error);
        }
    }
}

module.exports = SaveHtmlToFolder;

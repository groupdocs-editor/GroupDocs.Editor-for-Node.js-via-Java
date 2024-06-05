const {
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');

class GetAllEmbeddedHtmlContent {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            const embeddedHtmlContent = document.getEmbeddedHtml();
            console.log("HTML content of the input document, where all resources are embedded in base64 encoding: ", embeddedHtmlContent);
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = GetAllEmbeddedHtmlContent;

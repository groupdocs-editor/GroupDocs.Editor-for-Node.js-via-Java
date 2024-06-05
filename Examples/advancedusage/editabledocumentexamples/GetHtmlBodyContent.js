const {
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');

class GetHtmlBodyContent {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            const bodyContent = document.getBodyContent();
            console.log("Inner content of the HTML->BODY element: ", bodyContent);
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = GetHtmlBodyContent;

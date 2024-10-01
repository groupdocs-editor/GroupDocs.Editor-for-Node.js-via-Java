const {
    EditableDocument,
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');

class GetHtmlBodyContentWithPrefix {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            const externalImagesPrefix = "http://www.mywebsite.com/images/id=";
            const prefixedBodyContent = document.getBodyContent(externalImagesPrefix);
            console.log("Inner content of the HTML->BODY element with external images prefix: ", prefixedBodyContent);
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = GetHtmlBodyContentWithPrefix;

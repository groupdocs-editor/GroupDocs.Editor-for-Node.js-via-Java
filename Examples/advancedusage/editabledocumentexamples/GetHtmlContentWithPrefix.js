const {
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');

class GetHtmlContentWithPrefix {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            const externalImagesPrefix = "http://www.mywebsite.com/images/id=";
            const externalCssPrefix = "http://www.mywebsite.com/css/id=";
            const prefixedHtmlContent = document.getContentString(externalImagesPrefix, externalCssPrefix);
            console.log("HTML content of the input document with custom image and stylesheet prefixes: ", prefixedHtmlContent);
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = GetHtmlContentWithPrefix;

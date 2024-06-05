const {
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');

class GetExternalCssContentWithPrefix {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            const externalImagesPrefix = "http://www.mywebsite.com/images/id=";
            const externalFontsPrefix = "http://www.mywebsite.com/fonts/id=";
            const stylesheets = document.getCssContent(externalImagesPrefix, externalFontsPrefix).toArray();
            console.log(`There are ${stylesheets.length} stylesheets in the input document`);
            for (const css of stylesheets) {
                console.log(css);
            }
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = GetExternalCssContentWithPrefix;

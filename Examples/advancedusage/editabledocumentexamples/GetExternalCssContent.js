const {
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions, IHtmlResource
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');

class GetExternalCssContent {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());
            // const array = [];
            // const stylesheets2 = document.getCssContent().copyTo(array);
            // const stylesheets3 = Array.from(document.getCssContent());.forEach(IHtmlResource.class, (css) => {
            //                 console.log(css)
            //             })
            const stylesheets = document.getCssContent().toArray();
            console.log(`There are ${stylesheets.length} stylesheets in the input document`);
            for (const css of stylesheets) {
                console.log(css);
            }
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = GetExternalCssContent;

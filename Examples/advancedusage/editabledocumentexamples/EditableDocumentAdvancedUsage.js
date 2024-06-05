const {
    EditableDocument,
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const path = require('path');
const Constants = require('../../Constants');  // Assuming you have a Constants module

class EditableDocumentAdvancedUsage {
    static async run() {
        try {
            // 1. Get path to the file with HTML BODY inner markup
            const inputFilePath = Constants.SAMPLE_DOCX;
            const fileName = Constants.removeExtension(path.basename(inputFilePath));

            // 2. Initialize Editor with the input file and options
            const editor = new Editor(inputFilePath, new WordProcessingLoadOptions());

            // 3. Edit the document
            const beforeEdit = await editor.edit(new WordProcessingEditOptions());

            // 4. Get different parts of the document
            const allAsHtmlInsideOneString = beforeEdit.getEmbeddedHtml();

            // 5. Extract all images
            const allImages = beforeEdit.getImages();

            // 6. Extract all fonts
            const allFonts = beforeEdit.getFonts();

            // 7. Extract all stylesheets
            const allStylesheets = beforeEdit.getCss();

            // 8. Gather all resources
            const allResources = beforeEdit.getAllResources();

            // 9. Obtain HTML markup of the document (without embedded resources)
            const htmlMarkup = beforeEdit.getContent();

            // 10. Generate prefixed HTML markup
            const customImagesRequesthandlerUri = "http://example.com/ImagesHandler/id=";
            const customCssRequesthandlerUri = "http://example.com/CssHandler/id=";
            const customFontsRequesthandlerUri = "http://example.com/FontsHandler/id=";

            const prefixedHtmlMarkup = beforeEdit.getContentString(customImagesRequesthandlerUri, customCssRequesthandlerUri);

            // 11. Get only BODY content with prefix
            const onlyBodyContent = beforeEdit.getBodyContent();
            const prefixedBodyContent = beforeEdit.getBodyContent(customImagesRequesthandlerUri);

            // 12. Get all stylesheets content
            const stylesheets = beforeEdit.getCssContent();
            const prefixedStylesheets = beforeEdit.getCssContent(customImagesRequesthandlerUri, customFontsRequesthandlerUri);

            // 13. Save the document as an HTML file with resources
            const htmlFilePath = Constants.getOutputFilePath(fileName, 'html');
            await beforeEdit.save(htmlFilePath);

            // 14. Check if the EditableDocument is disposed
            const res = !beforeEdit.isDisposed() ? 'not' : 'already';
            console.log(`beforeEdit variable of EditableDocument type is ${res} disposed`);

            // 15. Create EditableDocument instances from file and markup
            const afterEditFromFile = EditableDocument.fromFile(htmlFilePath, null);
            const afterEditFromMarkup = EditableDocument.fromMarkup(htmlMarkup, allResources);

            // 16. Dispose all EditableDocument instances
            beforeEdit.dispose();
            afterEditFromFile.dispose();
            afterEditFromMarkup.dispose();
            editor.dispose();

            console.log('EditableDocumentAdvancedUsage routine has successfully finished');
        } catch (error) {
            console.error('Error:', error);
        }
    }
}

module.exports = EditableDocumentAdvancedUsage;

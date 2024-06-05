// noinspection JSUnresolvedReference

const {Editor, WordProcessingFormats, SpreadsheetFormats} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');

class EditDocument {
    constructor() {
        this.editor1 = new Editor(Constants.SAMPLE_DOCX, {format: WordProcessingFormats.Docx});
        this.editor2 = new Editor(Constants.SAMPLE_XLSX, {format: SpreadsheetFormats.Xlsx});
    }

    async run() {
        try {
            const defaultWordProcessingDoc = await this.editor1.edit();
            const wordProcessingEditOptions1 = {
                enablePagination: false,
                enableLanguageInformation: true,
                fontExtraction: 'ExtractAllEmbedded'
            };
            const version1WordProcessingDoc = await this.editor1.edit(wordProcessingEditOptions1);

            const wordProcessingEditOptions2 = {
                enablePagination: true, // Example change for demonstration
                fontExtraction: 'ExtractAll'
            };
            const version2WordProcessingDoc = await this.editor1.edit(wordProcessingEditOptions2);

            const sheetTab1EditOptions = {worksheetIndex: 0};
            const firstTab = await this.editor2.edit(sheetTab1EditOptions);

            const sheetTab2EditOptions = {worksheetIndex: 1};
            const secondTab = await this.editor2.edit(sheetTab2EditOptions);

            const bodyContent = firstTab.getBodyContent();
            const allContent = firstTab.getContent();
            const onlyImages = firstTab.getImages();
            const allResourcesTogether = firstTab.getAllResources();

            defaultWordProcessingDoc.dispose();
            version1WordProcessingDoc.dispose();
            version2WordProcessingDoc.dispose();
            firstTab.dispose();
            secondTab.dispose();
            this.editor1.dispose();
            this.editor2.dispose();

            console.log("EditDocument routine has successfully finished");
        } catch (error) {
            console.log("Error:", error);
        }
    }
}

module.exports = EditDocument;

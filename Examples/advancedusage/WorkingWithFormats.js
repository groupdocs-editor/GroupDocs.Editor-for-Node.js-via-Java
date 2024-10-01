const {
    WordProcessingFormats,
    PresentationFormats,
    SpreadsheetFormats,
    TextualFormats
} = require('@groupdocs/groupdocs.editor');

class WorkingWithFormats {
    static async run() {
        // Parsing from extension
        WordProcessingFormats.getAll().toArray().forEach(element => {
            console.log(`Parsed WordProcessing format is ${element.getName()}`);
        });
        PresentationFormats.getAll().toArray().forEach(element => {
            console.log(`Parsed Presentation format is ${element.getName()}`);
        });
        SpreadsheetFormats.getAll().toArray().forEach(element => {
            console.log(`Parsed Spreadsheet format is ${element.getName()}`);
        });
        const expectedXlsm = SpreadsheetFormats.fromExtension('.xlsm');
        console.log(`Parsed Spreadsheet format is ${expectedXlsm.getName()}`);

        const expectedHtml = TextualFormats.fromExtension('html');
        console.log(`Parsed Textual format is ${expectedHtml.getName()}`);
    }
}

module.exports = WorkingWithFormats;

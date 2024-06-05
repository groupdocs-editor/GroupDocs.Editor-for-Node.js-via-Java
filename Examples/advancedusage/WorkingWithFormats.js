const {
    SpreadsheetFormats,
    TextualFormats
} = require('@groupdocs/groupdocs.editor');

class WorkingWithFormats {
    static async run() {
        // Parsing from extension
        const expectedXlsm = SpreadsheetFormats.fromExtension('.xlsm');
        console.log(`Parsed Spreadsheet format is ${expectedXlsm.getName()}`);

        const expectedHtml = TextualFormats.fromExtension('html');
        console.log(`Parsed Textual format is ${expectedHtml.getName()}`);
    }
}

module.exports = WorkingWithFormats;

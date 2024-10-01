const {
    Editor,
    IncorrectPasswordException,
    PasswordRequiredException,
    WordProcessingDocumentInfo,
    SpreadsheetDocumentInfo,
    TextualDocumentInfo
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');

class ExtractingDocumentInfo {
    static async run() {
        console.log("****************************************");
        console.log("Starting 'ExtractingDocumentInfo' routine");

        const docxInputFilePath = Constants.SAMPLE_DOCX;
        const editorDocx = new Editor(docxInputFilePath);
        const infoDocx = await editorDocx.getDocumentInfo(null);
        const type = "WordProcessing";
        console.log(`Is '${docxInputFilePath}' a Spreadsheet: ${infoDocx.getFormat().getFormatFamily() == "Spreadsheet"? "yes" : "no"}`);
        console.log(`Is '${docxInputFilePath}' a Textual document: ${infoDocx.getFormat().getFormatFamily() ==  "Textual" ? "yes" : "no"}`);
        console.log(`Is '${docxInputFilePath}' a WordProcessing document: ${infoDocx.getFormat().getFormatFamily() == "WordProcessing" ? "yes" : "no"}`);
        console.log(`Format is: ${infoDocx.getFormat().getName()}; extension is: ${infoDocx.getFormat().getExtension()}; Page count: ${infoDocx.getPageCount()}; Size: ${infoDocx.getSize()} bytes; Is encrypted: ${infoDocx.isEncrypted()}`);
        const xlsxInputFilePath = Constants.SAMPLE_XLSX;
        const editorXlsx = new Editor(xlsxInputFilePath);
        const infoXlsx = await editorXlsx.getDocumentInfo(null);

        console.log(`It is:\r\nWordProcessing: ${infoXlsx instanceof WordProcessingDocumentInfo}\r\nSpreadsheet: ${infoXlsx instanceof SpreadsheetDocumentInfo}\r\nTextual: ${infoXlsx instanceof TextualDocumentInfo}`);

        {
            const casted = infoXlsx;
            console.log(`Format is: ${casted.getFormat().getName()}; extension is: ${casted.getFormat().getExtension()}; Tabs count: ${casted.getPageCount()}; Size: ${casted.getSize()} bytes; Is encrypted: ${casted.isEncrypted()}`);
        }

        const xlsInputFilePath = Constants.SAMPLE_XLS_PROTECTED;
        const editorXls = new Editor(xlsInputFilePath);

        try {
            await editorXls.getDocumentInfo(null);
        } catch (error) {
            if (error instanceof PasswordRequiredException) {
                console.log("Oops! We tried to open a password-protected document without password");
            }
        }

        try {
            await editorXls.getDocumentInfo("I don't know the password...");
        } catch (error) {
            if (error instanceof IncorrectPasswordException) {
                console.log("Oops! We specified password at this time, but it is incorrect");
            }
        }

        const infoXls = await editorXls.getDocumentInfo("excel_password");

        console.log(`Password-protected document actually is:\r\nWordProcessing: ${infoXls instanceof WordProcessingDocumentInfo}\r\nSpreadsheet: ${infoXls instanceof SpreadsheetDocumentInfo}\r\nTextual: ${infoXls instanceof TextualDocumentInfo}`);

        {
            const casted = infoXls;
            console.log(`Format is: ${casted.getFormat().getName()}; extension is: ${casted.getFormat().getExtension()}; Tabs count: ${casted.getPageCount()}; Size: ${casted.getSize()} bytes; Is encrypted: ${casted.isEncrypted()}`);
        }

        const xmlInputFilePath = Constants.SAMPLE_XML;
        const editorXml = new Editor(xmlInputFilePath);
        const infoXml = await editorXml.getDocumentInfo(null);
        console.log(`XML document is:\r\nWordProcessing: ${infoXml instanceof WordProcessingDocumentInfo}\r\nSpreadsheet: ${infoXml instanceof SpreadsheetDocumentInfo}\r\nTextual: ${infoXml instanceof TextualDocumentInfo}`);

        {
            const casted = infoXml;
            console.log(`Format is: ${casted.getFormat().getName()}; extension is: ${casted.getFormat().getExtension()}; Encoding: ${casted.getEncoding()}; Size: ${casted.getSize()} bytes`);
        }

        const txtInputFilePath = Constants.SAMPLE_TXT;
        const editorTxt = new Editor(txtInputFilePath);
        const infoTxt = await editorTxt.getDocumentInfo(null);
        console.log(`Text document is:\r\nWordProcessing: ${infoTxt instanceof WordProcessingDocumentInfo}\r\nSpreadsheet: ${infoTxt instanceof SpreadsheetDocumentInfo}\r\nTextual: ${infoTxt instanceof TextualDocumentInfo}`);

        {
            const casted = infoTxt;
            console.log(`Format is: ${casted.getFormat().getName()}; extension is: ${casted.getFormat().getExtension()}; Encoding: ${casted.getEncoding()}; Size: ${casted.getSize()} bytes`);
        }

        editorDocx.dispose();
        editorXlsx.dispose();
        editorXls.dispose();
        editorXml.dispose();
        editorTxt.dispose();

        console.log("ExtractingDocumentInfo routine has successfully finished");
    }
}

module.exports = ExtractingDocumentInfo;

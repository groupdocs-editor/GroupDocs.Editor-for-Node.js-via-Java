const {
    Editor,
    SpreadsheetFormats,
    SpreadsheetEditOptions,
    SpreadsheetLoadOptions,
    SpreadsheetSaveOptions,
    StreamBuffer,
    PasswordRequiredException, IncorrectPasswordException
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const path = require('path');
const {writeFile} = require("node:fs");

class WorkingWithSpreadsheetPasswordProtected {
    static async run() {
        const inputFilePath = Constants.SAMPLE_XLS_PROTECTED;

        let editor = new Editor(inputFilePath);
        try {
            await editor.edit();
        } catch (error) {
            if (error instanceof PasswordRequiredException) {
                console.log("Cannot edit a document, because it is password-protected, so the password is required");
            }
        }
        editor.dispose();

        const loadOptions = new SpreadsheetLoadOptions();
        loadOptions.setPassword("incorrect_password");
        editor = new Editor(inputFilePath, loadOptions);
        try {
            await editor.edit();
        } catch (error) {
            if (error instanceof IncorrectPasswordException) {
                console.log("Cannot edit a document, because it is password-protected, and password was specified, but it is incorrect");
            }
        }
        editor.dispose();

        loadOptions.setPassword("excel_password");
        loadOptions.setOptimizeMemoryUsage(true);
        editor = new Editor(inputFilePath, loadOptions);

        const editOptions = new SpreadsheetEditOptions();
        const beforeEdit = await editor.edit(editOptions);

        const xlsmFormat = SpreadsheetFormats.Xlsm;
        const saveOptions = new SpreadsheetSaveOptions(xlsmFormat);
        saveOptions.setPassword("new password");

        const outputPath = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)), xlsmFormat.getExtension());
        const bufferOutput = new StreamBuffer();
        await editor.save(beforeEdit, bufferOutput, saveOptions);
        writeFile(outputPath, bufferOutput.toByteArray(), (err) => {
            if (err) {
                console.error('Error writing the file:', err);
            } else {
                console.log('File written successfully.');
            }
        });


        beforeEdit.dispose();
        editor.dispose();

        console.log(`WorkingWithSpreadsheetPasswordProtected routine has successfully finished. Editor instance was manually disposed`);
    }
}

module.exports = WorkingWithSpreadsheetPasswordProtected;

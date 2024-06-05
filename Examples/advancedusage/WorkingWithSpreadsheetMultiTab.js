const {
    Editor,
    EditableDocument,
    SpreadsheetFormats,
    SpreadsheetEditOptions,
    SpreadsheetLoadOptions,
    SpreadsheetSaveOptions, readDataFromStream, StreamBuffer
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const fs = require('fs-extra');
const path = require('path');
const {writeFile} = require("node:fs");

class WorkingWithSpreadsheetMultiTab {
    static async run() {
        const inputFilePath = Constants.SAMPLE_XLSX;
        const inputStream = fs.createReadStream(inputFilePath);
        const stream = await readDataFromStream(inputStream);
        const editor = new Editor(stream, new SpreadsheetLoadOptions());

        const editOptions1 = new SpreadsheetEditOptions();
        editOptions1.setWorksheetIndex(0);
        const firstTabBeforeEdit = await editor.edit(editOptions1);

        const editOptions2 = new SpreadsheetEditOptions();
        editOptions2.setWorksheetIndex(1);
        const secondTabBeforeEdit = await editor.edit(editOptions2);

        const saveOptions1 = new SpreadsheetSaveOptions(SpreadsheetFormats.Xlsm);
        const outputPath1 = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)) + "_tab1", "xlsm");
        const bufferOutput = new StreamBuffer();
        await editor.save(firstTabBeforeEdit, outputPath1, saveOptions1);
        writeFile(outputPath1, bufferOutput.toByteArray(), (err) => {
            if (err) {
                console.error('Error writing the file:', err);
            } else {
                console.log('File written successfully.');
            }
        });

        const saveOptions2 = new SpreadsheetSaveOptions(SpreadsheetFormats.Xlsb);
        const outputPath2 = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)) + "_tab2", "xlsb");

        await editor.save(secondTabBeforeEdit, outputPath2, saveOptions2);

        firstTabBeforeEdit.dispose();
        secondTabBeforeEdit.dispose();
        editor.dispose();

        console.log("WorkingWithSpreadsheetMultiTab routine has successfully finished");
    }
}

module.exports = WorkingWithSpreadsheetMultiTab;

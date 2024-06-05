const {
    EditableDocument,
    Editor,
    DelimitedTextEditOptions,
    DelimitedTextSaveOptions,
    SpreadsheetSaveOptions,
    SpreadsheetFormats, StandardCharsets
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const path = require('path');
const fs = require('fs-extra');

class WorkingWithDsv {
    static async run() {
        try {
            const inputFilePath = Constants.SAMPLE_CSV;
            const fileName = Constants.removeExtension(path.basename(inputFilePath));

            const editor = new Editor(inputFilePath);

            const editOptions = new DelimitedTextEditOptions(",");
            editOptions.setConvertDateTimeData(false);
            editOptions.setConvertNumericData(true);
            editOptions.setTreatConsecutiveDelimitersAsOne(true);

            const beforeEdit = await editor.edit(editOptions);

            const originalTextContent = beforeEdit.getContent();
            const updatedTextContent = originalTextContent.replace("SsangYong", "Chevrolet").replace("Kyron", "Camaro");
            const allResources = beforeEdit.getAllResources();

            const afterEdit = EditableDocument.fromMarkup(updatedTextContent, allResources);

            const csvSaveOptions = new DelimitedTextSaveOptions(",");
            csvSaveOptions.setEncoding(StandardCharsets.UTF_8);

            const tsvSaveOptions = new DelimitedTextSaveOptions("\t");
            tsvSaveOptions.setTrimLeadingBlankRowAndColumn(false);
            tsvSaveOptions.setEncoding(StandardCharsets.UTF_8);

            const cellsSaveOptions = new SpreadsheetSaveOptions(SpreadsheetFormats.Xlsm);

            const outputCsvPath = Constants.getOutputFilePath(fileName, "csv");
            const outputTsvPath = Constants.getOutputFilePath(fileName, "tsv");
            const outputCellsPath = Constants.getOutputFilePath(fileName, "xlsm");

            await editor.save(afterEdit, outputCsvPath, csvSaveOptions);
            await editor.save(afterEdit, outputTsvPath, tsvSaveOptions);
            await editor.save(afterEdit, outputCellsPath, cellsSaveOptions);

            beforeEdit.dispose();
            afterEdit.dispose();

            console.log("WorkingWithDsv routine has successfully finished");
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = WorkingWithDsv;

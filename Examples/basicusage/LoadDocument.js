const { Editor, WordProcessingLoadOptions, SpreadsheetLoadOptions, readDataFromStream} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const fs = require("fs");

class LoadDocument {
    constructor() {}

    async run() {
        try {
            const inputPath = Constants.SAMPLE_DOCX;
            console.log(`LoadDocument: try to open document ${inputPath}`);

            // Load document as file via path and without load options
            const editor1 = new Editor(inputPath);
            editor1.dispose();

            // Load document as file via path and with load options
            const wordLoadOptions = new WordProcessingLoadOptions();
            wordLoadOptions.setPassword("some password");
            const editor2 = new Editor(inputPath, wordLoadOptions);
            editor2.dispose();

            // Load document as content from byte stream and without load options
            const readStream = fs.createReadStream(Constants.SAMPLE_XLSX)
            const stream = await readDataFromStream(readStream);
            const editor3 = new Editor(stream);
            editor3.dispose();
            stream.close();

            // Load document as content from byte stream and with load options
            const sheetLoadOptions = new SpreadsheetLoadOptions();
            sheetLoadOptions.setOptimizeMemoryUsage(true);
            const inputStream2 = fs.createReadStream(Constants.SAMPLE_XLSX);
            const stream2 = await readDataFromStream(inputStream2);
            const editor4 = new Editor(stream2, sheetLoadOptions);
            editor4.dispose();
            inputStream2.close();

            console.log("LoadDocument routine has successfully finished");
        } catch (error) {
            console.log("Error:", error);
        }
    }
}

module.exports = LoadDocument;

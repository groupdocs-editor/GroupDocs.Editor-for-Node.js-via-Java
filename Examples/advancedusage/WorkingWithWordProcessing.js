const {
    Editor,
    EditableDocument,
    WordProcessingFormats,
    WordProcessingEditOptions,
    WordProcessingLoadOptions,
    WordProcessingSaveOptions, readDataFromStream, ByteArrayOutputStream
} = require('@groupdocs/groupdocs.editor');
const fs = require("fs");
const Constants = require('../Constants');
const path = require('path');
const {writeFile} = require("node:fs");

class WorkingWithWordProcessing {
    static async run() {
        const inputFilePath = Constants.SAMPLE_DOCX;
        const inputStream = fs.createReadStream(inputFilePath);
        const buffer = await readDataFromStream(inputStream);
        const loadOptions = new WordProcessingLoadOptions();
        loadOptions.setPassword('some_password_to_open_a_document');

        const editor = new Editor(buffer, loadOptions);

        const editOptions = new WordProcessingEditOptions();
        editOptions.setEnableLanguageInformation(true);
        editOptions.setEnablePagination(true);

        const beforeEdit = await editor.edit(editOptions);
        const originalContent = beforeEdit.getContent();
        const allResources = beforeEdit.getAllResources();

        const editedContent = originalContent.replace('document', 'edited document');
        const afterEdit = EditableDocument.fromMarkup(editedContent, allResources);

        const docmFormat = WordProcessingFormats.Docm;
        const saveOptions = new WordProcessingSaveOptions(docmFormat);
        saveOptions.setPassword('password');
        saveOptions.setEnablePagination(true);
        saveOptions.setOptimizeMemoryUsage(true);

        const outputPath = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)), docmFormat.getExtension());
        const outputStream = new ByteArrayOutputStream();
            
        await editor.save(afterEdit, outputStream, saveOptions);
        const byteArray = outputStream.toByteArray();
            const bufferOutput = Buffer.from(byteArray);
            fs.writeFileSync(outputPath, bufferOutput);


        beforeEdit.dispose();
        afterEdit.dispose();
        editor.dispose();

        console.log('WorkingWithWordProcessing routine has successfully finished');
    }
}

module.exports = WorkingWithWordProcessing;

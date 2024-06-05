const {
    Editor,
    EditableDocument,
    WordProcessingFormats,
    TextEditOptions,
    TextLeadingSpacesOptions,
    TextTrailingSpacesOptions,
    TextSaveOptions,
    WordProcessingSaveOptions, StandardCharsets, Locale
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const path = require('path');
const {readFileSync} = require('fs');

class WorkingWithPlainTextDocuments {
    static async run() {
        const inputFilePath = Constants.SAMPLE_TXT;

        const editor = new Editor(inputFilePath);

        const editOptions = new TextEditOptions();
        editOptions.setEncoding(StandardCharsets.UTF_8);
        editOptions.setRecognizeLists(true);
        editOptions.setLeadingSpaces(TextLeadingSpacesOptions.ConvertToIndent);
        editOptions.setTrailingSpaces(TextTrailingSpacesOptions.Trim);

        const beforeEdit = await editor.edit(editOptions);

        const originalTextContent = beforeEdit.getContent();
        const updatedTextContent = originalTextContent.replace("text", "EDITED text");

        const allResources = beforeEdit.getAllResources();
        const afterEdit = EditableDocument.fromMarkup(updatedTextContent, allResources);

        const wordSaveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Docm);
        wordSaveOptions.setLocale(Locale.US);

        const txtSaveOptions = new TextSaveOptions();
        txtSaveOptions.setEncoding(StandardCharsets.UTF_8);
        txtSaveOptions.setPreserveTableLayout(true);

        const outputWordPath = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)), "docm");
        const outputTxtPath = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)), "txt");

        await editor.save(afterEdit, outputWordPath, wordSaveOptions);
        await editor.save(afterEdit, outputTxtPath, txtSaveOptions);

        beforeEdit.dispose();
        afterEdit.dispose();
        editor.dispose();

        console.log("WorkingWithPlainTextDocuments routine has successfully finished");
    }
}

module.exports = WorkingWithPlainTextDocuments;

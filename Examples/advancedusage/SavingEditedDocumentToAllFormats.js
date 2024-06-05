const {
    Editor,
    EditableDocument,
    WordProcessingFormats,
    WordProcessingLoadOptions,
    WordProcessingEditOptions,
    WordProcessingSaveOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const path = require('path');

class SavingEditedDocumentToAllFormats {
    static async run() {
        try {
            const inputFilePath = Constants.SAMPLE_DOCX;
            const loadOptions = new WordProcessingLoadOptions();
            const editor = new Editor(inputFilePath, loadOptions);
            const editOptions = new WordProcessingEditOptions();

            const beforeEdit = await editor.edit(editOptions);
            let allEmbeddedInsideString = beforeEdit.getEmbeddedHtml();
            let allEmbeddedInsideStringEdited = allEmbeddedInsideString.replace("Subtitle", "Edited subtitle");

            const afterEdit = EditableDocument.fromMarkup(allEmbeddedInsideStringEdited, null);

            let oneFormat = WordProcessingFormats.Doc;
            let saveOptions = new WordProcessingSaveOptions(oneFormat);
            let savePath = Constants.getOutputFilePath("MultipleOutFormats", saveOptions.getOutputFormat().getExtension());
            await editor.save(afterEdit, savePath, saveOptions);
            oneFormat = WordProcessingFormats.Docx;
            saveOptions = new WordProcessingSaveOptions(oneFormat);
            savePath = Constants.getOutputFilePath("MultipleOutFormats", saveOptions.getOutputFormat().getExtension());
            await editor.save(afterEdit, savePath, saveOptions);
            oneFormat = WordProcessingFormats.Rtf;
            saveOptions = new WordProcessingSaveOptions(oneFormat);
            savePath = Constants.getOutputFilePath("MultipleOutFormats", saveOptions.getOutputFormat().getExtension());
            await editor.save(afterEdit, savePath, saveOptions);

            console.log("SavingEditedDocumentToAllFormats routine has successfully finished");
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = SavingEditedDocumentToAllFormats;

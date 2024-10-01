const {
    Editor,
    EditableDocument,
    WordProcessingFormats,
    WordProcessingSaveOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');

class WorkingWithMarkdown {
    static async run() {
        const mdInputPath = Constants.SAMPLE_MD;
        const mdEditor = new Editor(mdInputPath);

        const info = await mdEditor.getDocumentInfo(null);

        const doc = await mdEditor.edit(null);

        const saveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Docx);

        const outputPath = Constants.getOutputFilePath("OutputMd", saveOptions.getOutputFormat().getExtension());
        await mdEditor.save(doc, outputPath, saveOptions);

        mdEditor.dispose();
        console.log("WorkingWithMarkdown routine has successfully finished");
    }
}

module.exports = WorkingWithMarkdown;

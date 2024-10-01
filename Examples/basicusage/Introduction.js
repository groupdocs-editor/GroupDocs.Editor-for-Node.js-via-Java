const { Editor, WordProcessingFormats, EditableDocument, WordProcessingSaveOptions, Action} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const fs = require('fs');

class Introduction {
    constructor() {
        this.inputFilePath = Constants.SAMPLE_DOCX;
        console.log(`Introduction: try to open document ${this.inputFilePath}`);
        this.outputPath = Constants.getOutputFilePath("IntroductionOutput", "rtf");
        this.outputPathForStream = Constants.getOutputFilePath("IntroductionOutputFromStream", "rtf");
        this.saveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Rtf);
        this.editor = new Editor(this.inputFilePath);
    }

    async run() {
        try {
            const beforeEdit = await this.editor.edit();
            const allEmbeddedInsideString = beforeEdit.getEmbeddedHtml();
            const allEmbeddedInsideStringEdited = allEmbeddedInsideString.replace("Subtitle", "Edited subtitle");
            const afterEdit = EditableDocument.fromMarkup(allEmbeddedInsideStringEdited, null);
            await this.editor.save(afterEdit, this.outputPath, this.saveOptions);

            beforeEdit.dispose();
            afterEdit.dispose();
            this.editor.dispose();

            console.log("Introduction routine has successfully finished");
        } catch (error) {
            console.log("Error:", error);
        }
    }
}

module.exports = Introduction;

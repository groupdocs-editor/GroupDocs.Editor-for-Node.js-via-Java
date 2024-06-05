const {Editor} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const path = require('path');

class WorkingWithPresentationPreview {
    static async run() {
        const inputPath = Constants.getSampleFilePath("FormatingExample.pptx");
        const outputFolder = Constants.getOutputDirectoryPath("");

        const editor = new Editor(inputPath);

        const infoUncasted = await editor.getDocumentInfo(null);

        const slidesCount = infoUncasted.getPageCount();

        for (let i = 0; i < slidesCount; i++) {
            const oneSvgPreview = await infoUncasted.generatePreview(i);
            const outputFilePath = path.join(outputFolder, oneSvgPreview.getFilenameWithExtension());
            oneSvgPreview.save(outputFilePath)
            console.log(`Saved slide ${i + 1} preview to ${outputFilePath}`);
        }

        editor.dispose();
        console.log("WorkingWithPresentationPreview routine has successfully finished");
    }
}

module.exports = WorkingWithPresentationPreview;

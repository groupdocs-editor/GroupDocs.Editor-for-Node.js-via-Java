const {
    Editor,
    EditableDocument,
    MarkdownEditOptions,
    MarkdownSaveOptions,
    MarkdownTableContentAlignment
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');

class MarkdownRoundtrip {
    static async run() {
        const inputFolderPath = Constants.getSampleFilePath("Markdown");
        const outputFolder = Constants.getOutputDirectoryPath("");
        const outputMdPath = Constants.getOutputFilePath("MarkdownRoundtripOutput", "md");

        const inputPath = Constants.SAMPLE_MD;

        const editOptions = new MarkdownEditOptions();
        editOptions.setImageLoadCallback(new MdImageLoader(inputFolderPath));

        const saveOptions = new MarkdownSaveOptions();
        saveOptions.setTableContentAlignment(MarkdownTableContentAlignment.Center);
        saveOptions.setImagesFolder(outputFolder);

        const editor = new Editor(inputPath);
        try {
            const doc = await editor.edit(editOptions);
            await editor.save(doc, outputMdPath, saveOptions);
        } finally {
            editor.dispose();
        }

        console.log("MarkdownRoundtrip routine has successfully finished");
    }
}

class MdImageLoader {
    constructor(imagesFolder) {
        this._imagesFolder = imagesFolder;
    }

    processImage(args) {
        const filePath = path.join(this._imagesFolder, path.basename(args.getImageFileName()));
        try {
            const data = fs.readFileSync(filePath);
            args.setData(data);
        } catch (error) {
            throw new Error(error.message);
        }
        return MarkdownImageLoadingAction.UserProvided;
    }
}

module.exports = MarkdownRoundtrip;

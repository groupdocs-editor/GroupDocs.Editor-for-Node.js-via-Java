const {
    Editor,
    EditableDocument,
    WordProcessingFormats,
    MarkdownEditOptions,
    MarkdownImageLoadArgs,
    WordProcessingSaveOptions,
    MarkdownImageLoadingAction
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const path = require('path');
const fs = require('fs-extra');
const assert = require('assert');

class EditingMarkdown {
    static async run() {
        try {
            const inputMdPath = Constants.SAMPLE_MD_EDIT;
            const imagesFolder = Constants.SAMPLE_MD_FOLDER;

            const editOptions = new MarkdownEditOptions();
            editOptions.setImageLoadCallback(new MdImageLoader(imagesFolder));

            const editor = new Editor(inputMdPath);

            const beforeEdit = await editor.edit(editOptions);
            const originalHtmlContent = beforeEdit.getEmbeddedHtml();
            const afterEdit = EditableDocument.fromMarkup(originalHtmlContent, null);
            const saveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Docx);
            const outputDocxPath = Constants.getOutputFilePath("OutputEditingMarkdown", saveOptions.getOutputFormat().getExtension());
            await editor.save(afterEdit, outputDocxPath, saveOptions);
            editor.dispose();
            console.log("EditingMarkdown routine has successfully finished");
        } catch (error) {
            console.error("Error:", error);
        }
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

module.exports = EditingMarkdown;

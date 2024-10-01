const {
    EditableDocument,
    Editor,
    WordProcessingLoadOptions,
    WordProcessingEditOptions
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../../Constants');
const path = require('path');
const fs = require('fs-extra');

class SaveHtmlResourcesToFolder {
    static async run() {
        try {
            const editor = new Editor(Constants.SAMPLE_DOCX, new WordProcessingLoadOptions());
            const document = await editor.edit(new WordProcessingEditOptions());

            const images = document.getImages().toArray();
            const fonts = document.getFonts().toArray();
            const stylesheets = document.getCss().toArray();

            const outputFolder = Constants.getOutputDirectoryPath('');

            for (const image of images) {
                console.log(`Saving ${image.getFilenameWithExtension()} of ${image.getType().getFormalName()} type and ${image.getLinearDimensions()} dimensions`);
                image.save(path.join(outputFolder, image.getFilenameWithExtension()));
            }

            for (const font of fonts) {
                console.log(`Saving ${font.getFilenameWithExtension()} of ${font.getType().getFormalName()} type`);
                font.save(path.join(outputFolder, font.getFilenameWithExtension()));
            }

            for (const stylesheet of stylesheets) {
                console.log(`Saving ${stylesheet.getFilenameWithExtension()} of ${stylesheet.getType().getFormalName()} type and ${stylesheet.getEncoding()} encoding`);
                stylesheet.save(path.join(outputFolder, stylesheet.getFilenameWithExtension()))
            }

            console.log('Resources have been saved successfully');
        } catch (error) {
            console.error('Error:', error);
        }
    }
}

module.exports = SaveHtmlResourcesToFolder;

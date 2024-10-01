const fs = require('fs');
const path = require('path');
const Constants = require('../../Constants');

const { Editor, WordProcessingLoadOptions, WordProcessingEditOptions, FontExtractionOptions, readDataFromStream } = require('@groupdocs/groupdocs.editor'); // Assuming similar groupdocs-editor module exists

class WorkingWithResources {
    static async run() {
        try {
            /*
             * Almost every document contains or may contain some attachments of binary or textual nature; first of all they are images and/or fonts.
             * Also, when input document is converted to the intermediate EditableDocument instance,
             * all style-related data may be represented as stylesheets (CSS).
             * All of these are grouped in the GroupDocs.Editor.HtmlCss.Resources namespace, where every resource has its own class and methods.
             */
            console.log("****************************************");
            console.log("Investigating resources of DOCX document");

            // 1. Load and edit the input document in a supported format
            const outputFolderPath = Constants.getOutputDirectoryPath("WorkingWithResources");
            const inputPath = Constants.SAMPLE_DOCX2;
            // Ensure output folder exists
            if (!fs.existsSync(outputFolderPath)) {
                fs.mkdirSync(outputFolderPath);
            }

            // Create Editor instance
            const editor = new Editor(inputPath, new WordProcessingLoadOptions());

            // Set up WordProcessing edit options and enable font extraction
            const editOptions = new WordProcessingEditOptions();
            editOptions.fontExtraction = FontExtractionOptions.ExtractAll;

            // Edit document
            const beforeEdit = await editor.edit(editOptions);

            // 2. Gather resources
            const images = beforeEdit.getImages().toArray();
            const fonts = beforeEdit.getFonts().toArray();
            const stylesheets = beforeEdit.getCss().toArray();

            // 3. Print detailed info about every resource and save them to a file
            console.log(`${images.length} images are used:`);
            images.forEach((oneImage, i) => {
                console.log(
                    `${i}. ${oneImage.getFilenameWithExtension()}. Type: ${oneImage.getType().getFormalName()}. Aspect ratio: ${oneImage.getAspectRatio().toString()}. Dimensions: ${oneImage.getLinearDimensions().toString()}px`
                );
                oneImage.save(path.join(outputFolderPath, oneImage.getFilenameWithExtension()));
            });

            console.log(`${fonts.length} fonts are used:`);
            fonts.forEach((oneFont, i) => {
                console.log(
                    `${i}. ${oneFont.getFilenameWithExtension()}. Type: ${oneFont.getType().getFormalName()}.`
                );
                oneFont.save(path.join(outputFolderPath, oneFont.getFilenameWithExtension()));

            });

            console.log(`${stylesheets.length} stylesheets are used:`);
            stylesheets.forEach((oneStylesheet, i) => {
                console.log(
                    `${i}. ${oneStylesheet.getFilenameWithExtension()}. Type: ${oneStylesheet.getType().getFormalName()}.`
                );
                oneStylesheet.save(path.join(outputFolderPath, oneStylesheet.getFilenameWithExtension()));
            });

            // 4. Get resource content as a byte stream or base64 string
            const firstImageStream = images[0].getByteContent();  // Example: getting the byte content of the first image

            // 4.2 Base64 encoded string of the first image
            const base64EncodedResource = images[0].getTextContent(); 
            console.log(`Base64 encoded content of the first image: ${base64EncodedResource}`);

            // Dispose of all resources
            beforeEdit.dispose();
            editor.dispose();
            console.log("WorkingWithResources routine has successfully finished");

        } catch (error) {
            console.error('Error:', error);
        }
    }
}
module.exports = WorkingWithResources;

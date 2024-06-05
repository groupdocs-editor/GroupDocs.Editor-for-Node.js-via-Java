const {
    Editor,
    EditableDocument,
    PresentationFormats,
    PresentationEditOptions,
    PresentationLoadOptions,
    PresentationSaveOptions, readDataFromStream, StreamBuffer
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const fs = require('fs-extra');
const path = require('path');
const {writeFile} = require("node:fs");

class WorkingWithPresentations {
    static async run() {
        const inputFilePath = Constants.SAMPLE_PPTX;
        const inputFileStream = fs.createReadStream(inputFilePath);
        const stream = await readDataFromStream(inputFileStream);
        const loadOptions = new PresentationLoadOptions();
        loadOptions.setPassword('some_password_to_open_a_document');

        const editor = new Editor(stream, loadOptions);

        const editOptions = new PresentationEditOptions();
        editOptions.setSlideNumber(0);
        editOptions.setShowHiddenSlides(true);

        const beforeEdit = await editor.edit(editOptions);
        const originalContent = beforeEdit.getContent();
        const allResources = beforeEdit.getAllResources();

        const editedContent = originalContent.replace('New text', 'edited text');
        const afterEdit = EditableDocument.fromMarkup(editedContent, allResources);

        const saveOptions = new PresentationSaveOptions(PresentationFormats.Pptm);
        saveOptions.setPassword('password');

        const outputPath = Constants.getOutputFilePath('sample_out', saveOptions.getOutputFormat().getExtension());
        const buffer = new StreamBuffer();
        await editor.save(afterEdit, buffer, saveOptions);
        writeFile(outputPath, buffer.toByteArray(), (err) => {
            if (err) {
                console.error('Error writing the file:', err);
            } else {
                console.log('File written successfully.');
            }
        });


        beforeEdit.dispose();
        afterEdit.dispose();
        editor.dispose();

        console.log('WorkingWithPresentations routine has successfully finished');
    }
}

module.exports = WorkingWithPresentations;

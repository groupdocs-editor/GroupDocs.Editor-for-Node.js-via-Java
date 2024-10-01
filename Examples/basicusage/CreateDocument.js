const { Editor, WordProcessingFormats,WordProcessingEditOptions,WordProcessingSaveOptions,SpreadsheetFormats,
    SpreadsheetEditOptions,PresentationEditOptions,PresentationFormats,EmailFormats, EmailEditOptions, ByteArrayOutputStream} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const fs = require("fs");

class CreateDocument {
    constructor() {}

    async run() {
        try {
            // Create a new WordProcessing document and save it using a callback.
            console.log(`CreateDocument: Editing WordProcessing document ${WordProcessingFormats.Docx}`);

            // Edit the WordProcessing document with default options.
            const editorWord = new Editor(WordProcessingFormats.Docx);
            
            const savePath = Constants.getOutputFilePath("NewDocument", "docx");
            
            const outputStream = await editorWord.save(new ByteArrayOutputStream());
            const byteArray = outputStream.toByteArray();
            const buffer = Buffer.from(byteArray);
            fs.writeFileSync(savePath, buffer);

            console.log(`CreateDocument: saved document ${savePath}`);

            // Create a new Spreadsheet document and save it via callback.
            console.log(`CreateDocument: Editing Spreadsheet document ${SpreadsheetFormats.Xlsx}`);

            // Edit the Spreadsheet document with default options.
            const editorSpreadsheet = new Editor(SpreadsheetFormats.Xlsx);
            const defaultEditableSpreadsheetDocument = editorSpreadsheet.edit();
            editorSpreadsheet.dispose();

            // Create a new Presentation document and save it via callback.
            console.log(`CreateDocument: Editing Presentation document ${PresentationFormats.Pptx}`);

            // Edit the Presentation document with default options.
            const editorPresentation = new Editor(PresentationFormats.Pptx);

            // Edit the Presentation document with specified options and some defined settings.
            const presentationEditOptions = new PresentationEditOptions();
            presentationEditOptions.setShowHiddenSlides(false);
            presentationEditOptions.setSlideNumber(0);

            const editablePresentationDocument = editorPresentation.edit(presentationEditOptions);
            editorPresentation.dispose();

            // Create a new Email document and save it via callback.
            console.log(`CreateDocument: Editing Email document ${EmailFormats.Eml}`);

            // Edit the Email document with default options.
            const editorEmail = new Editor(EmailFormats.Eml);
            const defaultEditableEmailDocument = editorEmail.edit();
            editorEmail.dispose();

            console.log("CreateDocument routine has successfully finished");
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = CreateDocument;

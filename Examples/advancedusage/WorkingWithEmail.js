const {
    Editor,
    EditableDocument,
    EmailEditOptions,
    EmailSaveOptions,
    MailMessageOutput, StreamBuffer
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const fs = require('fs-extra');
const {writeFile} = require("node:fs");

class WorkingWithEmail {
    static async run() {
        try {
            const msgInputPath = Constants.SAMPLE_MSG;
            const msgEditor = new Editor(msgInputPath);

            const editOptions = new EmailEditOptions(MailMessageOutput.All);
            const originalDoc = await msgEditor.edit(editOptions);

            const savedHtmlContent = originalDoc.getEmbeddedHtml();
            const editedDoc = EditableDocument.fromMarkup(savedHtmlContent, null);

            const saveOptions1 = new EmailSaveOptions(MailMessageOutput.Common);
            const saveOptions2 = new EmailSaveOptions(MailMessageOutput.Body | MailMessageOutput.Attachments);

            const outputMsgPath = Constants.getOutputFilePath("OutputFile", "msg");
            await msgEditor.save(editedDoc, outputMsgPath, saveOptions1);

            const buffer = new StreamBuffer();
            await msgEditor.save(editedDoc, buffer, saveOptions2);
            writeFile(outputMsgPath.replace(".msg", "_stream.msg"), buffer.toByteArray(), (err) => {
                if (err) {
                    console.error('Error writing the file:', err);
                } else {
                    console.log('File written successfully.');
                }
            });

            editedDoc.dispose();
            originalDoc.dispose();
            msgEditor.dispose();

            console.log("WorkingWithEmail routine has successfully finished");
        } catch (error) {
            console.error("Error:", error);
        }
    }
}

module.exports = WorkingWithEmail;

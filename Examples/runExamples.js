const Constants = require('./Constants');
const SetLicenseFromFile = require('./quickstart/SetLicenseFromFile');
const SetLicenseFromStream = require('./quickstart/SetLicenseFromStream');
const HelloWorld = require('./quickstart/HelloWorld');

// Basic Usage
// const CreateDocument = require('./basicusage/CreateDocument'); will be fixed in the next release.
const Introduction = require('./basicusage/Introduction');
const LoadDocument = require('./basicusage/LoadDocument');
const EditDocument = require('./basicusage/EditDocument');
const SaveDocument = require('./basicusage/SaveDocument');

// Advanced Usage
const WorkingWithWordProcessing = require('./advancedusage/WorkingWithWordProcessing');
const WorkingWithSpreadsheetPasswordProtected = require('./advancedusage/WorkingWithSpreadsheetPasswordProtected');
const WorkingWithSpreadsheetMultiTab = require('./advancedusage/WorkingWithSpreadsheetMultiTab');
const WorkingWithDsv = require('./advancedusage/WorkingWithDsv');
const WorkingWithPresentations = require('./advancedusage/WorkingWithPresentations');
const WorkingWithPlainTextDocuments = require('./advancedusage/WorkingWithPlainTextDocuments');
const WorkingWithXml = require('./advancedusage/WorkingWithXml');
const ExtractingDocumentInfo = require('./advancedusage/ExtractingDocumentInfo');
const SavingEditedDocumentToAllFormats = require('./advancedusage/SavingEditedDocumentToAllFormats');
const WorkingWithFormats = require('./advancedusage/WorkingWithFormats');
const MarkdownRoundtrip = require('./advancedusage/MarkdownRoundtrip');
const WorkingWithEmail = require('./advancedusage/WorkingWithEmail');
const EditingMarkdown = require("./advancedusage/EditingMarkdown");
const WorkingWithMarkdown = require("./advancedusage/WorkingWithMarkdown");
const WorkingWithPresentationPreview = require("./advancedusage/WorkingWithPresentationPreview");

// EditableDocument Examples
const CreateEditableDocumentFromHtmlFile = require('./advancedusage/editabledocumentexamples/CreateEditableDocumentFromHtmlFile');
const CreateEditableDocumentFromInnerBody = require('./advancedusage/editabledocumentexamples/CreateEditableDocumentFromInnerBody');
const GetHtmlContent = require('./advancedusage/editabledocumentexamples/GetHtmlContent');
const GetHtmlContentWithPrefix = require('./advancedusage/editabledocumentexamples/GetHtmlContentWithPrefix');
const GetHtmlBodyContent = require('./advancedusage/editabledocumentexamples/GetHtmlBodyContent');
const GetHtmlBodyContentWithPrefix = require('./advancedusage/editabledocumentexamples/GetHtmlBodyContentWithPrefix');
const GetAllEmbeddedHtmlContent = require('./advancedusage/editabledocumentexamples/GetAllEmbeddedHtmlContent');
const GetExternalCssContent = require('./advancedusage/editabledocumentexamples/GetExternalCssContent');
const GetExternalCssContentWithPrefix = require('./advancedusage/editabledocumentexamples/GetExternalCssContentWithPrefix');
const SaveHtmlToFolder = require('./advancedusage/editabledocumentexamples/SaveHtmlToFolder');
const SaveHtmlResourcesToFolder = require('./advancedusage/editabledocumentexamples/SaveHtmlResourcesToFolder');
const EditableDocumentAdvancedUsage = require('./advancedusage/editabledocumentexamples/EditableDocumentAdvancedUsage');


async function main() {
    console.log(`Using GroupDocs.Editor for Node.js version ${require('@groupdocs/groupdocs.editor').version}`);
    console.log('Open main.js. In main() function uncomment the example that you want to run.');
    console.log(`Output folder is '${Constants.OutputPath}'`);
    console.log('=====================================================');
    // Quick Start
    await SetLicenseFromFile.run();
    // await SetLicenseFromStream.run();
    await HelloWorld.run();

    ////// *** Documents Editor Examples (Un-Comment to run each example demo methods) ***

    // Basic Usage
    await new Introduction().run();
    await new LoadDocument().run();
    await new EditDocument().run();
    await SaveDocument.run();


    // Advanced Usage
    await WorkingWithWordProcessing.run();
    await WorkingWithSpreadsheetPasswordProtected.run();
    await WorkingWithSpreadsheetMultiTab.run();
    await WorkingWithDsv.run();
    await WorkingWithPresentations.run();
    await WorkingWithPlainTextDocuments.run();
    await WorkingWithXml.run();
    await ExtractingDocumentInfo.run();
    await SavingEditedDocumentToAllFormats.run();
    await WorkingWithFormats.run();
    await EditingMarkdown.run();
    await MarkdownRoundtrip.run();
    await WorkingWithEmail.run();
    await WorkingWithMarkdown.run();
    await WorkingWithPresentationPreview.run();

    // EditableDocument Examples
    await CreateEditableDocumentFromHtmlFile.run();
    await CreateEditableDocumentFromInnerBody.run();
    await GetHtmlContent.run();
    await GetHtmlContentWithPrefix.run();
    await GetHtmlBodyContent.run();
    await GetHtmlBodyContentWithPrefix.run();
    await GetAllEmbeddedHtmlContent.run();
    await GetExternalCssContent.run();
    await GetExternalCssContentWithPrefix.run();
    await SaveHtmlToFolder.run();
    await SaveHtmlResourcesToFolder.run();
    await EditableDocumentAdvancedUsage.run();

    console.log('\r\n\r\n__________________________\r\nAll done.');
    process.exit(0);
}

main().catch(err => console.error(err));

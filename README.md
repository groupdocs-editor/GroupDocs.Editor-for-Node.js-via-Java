# Document Editor API for Node.js

GroupDocs.Editor for Node.js via Java is a high-performance and cross-platform library that enhances your apps to perform document, spreadsheets, DSV & XML files editing operations for a wide range of file formats. It supports over 170 document types from popular categories such as Microsoft Office, OpenOffice, metafiles, messages, PDF & more.

Edit DOCX, PPTX, XLSX, EML, MD among [many other documents](https://docs.groupdocs.com/editor/java/supported-document-formats/) in HTML5, PDF modes with fast and high-quality rendering. The API also provides the ability to customize document appearance via additional editing options, extract document text, and much more!

Edit Word, Excel and PowerPoint documents using GroupDocs.Editor for .NET powerful document editing API. It can be used with any external, open source or paid HTML editor.

GroupDocs.Editor API will load documents, convert it into HTML that allows to edit document with any external UI, after making changes to document content edited HTML will be saved to original document format.

GroupDocs.Editor can also be used to generate different Microsoft Word documents, Excel workbooks, PowerPoint presentations, XML, OpenDocument Formats and TXT files.

| Directory                                                                                                  | Description                                                           |
|------------------------------------------------------------------------------------------------------------|-----------------------------------------------------------------------|
| [Examples](https://github.com/groupdocs-editor/GroupDocs.Editor-for-Node.js-via-Java/tree/master/Examples) | Node.js examples and sample documents for you to get started quickly. |

## Node.js Editor API Features

- Edit document content in any web-browser.
- Edit document pages separately.
- Customizable resource management options for CSS, fonts & images.
- Render documents to HTML for editing.
- Boost document loading speed with configurable caching.
- Extract document text along with words' coordinates.
- Extract basic information about source documents such as file type, pages count, and so on.
- Auto-detect document type.
- Replace missing font or use custom fonts for rendering.
- Edit word processing documents in a flow or paged mode.
- Fetch language information for multi-lingual document editing.
- Extract font information to provide consistent editing and appearance behavior.
- Edit multi-tabbed spreadsheets.
- Supports DSV (Delimiter-Separated Values) documents.
- Specify separator, flexible numeric, and data conversion for CSV & TSV files.
- Optimize memory usage for large CSV & TSV files.
- Fix incorrect XML document structure.
- Recognize URIs and email addresses in XML files.
- Set character encoding of the input text document.
- Grab document metadata information.
- Fetch whole HTML document or BODY content.
- Get an HTML document along with all its resources (stylesheets, images).
- Open any supported format file in HTML format and save it to disk.
- Fetch HTML markup from DB or remote storage.

## Supported File Formats for Editing

GroupDocs.Editor supports editing a wide range of document formats across various categories. Below are the supported formats:

### WordProcessing family formats

- DOC, DOCX, DOCM, DOT, DOTX, DOTM, FlatOPC, ODT, OTT, RTF, WordML

### Spreadsheet family formats

- XLS, XLT, XLSX, XLSM, XLTX, XLTM, XLSB, XLAM, SpreadsheetML, ODS, FODS, SXC, DIF (Import, Export, Auto Detection), DSV (Import only), CSV (Import, Export, Auto Detection), TSV (Import, Export, Auto Detection)

### Presentation family formats

- PPT, PPTX, PPTM, PPS, PPSX, PPSM, POT, POTX, POTM, ODP, OTP, FODP (Create, Import, Auto Detection)

### Fixed-layout family formats

- PDF

### Email family formats

- EML, EMLX, MSG, MBOX, TNEF, MHT, PST, OFT, OST, VCF (Auto Detection only), ICS (Auto Detection only)

### eBook family formats

- ePub, MOBI, AZW3

### Markup formats

- HTML (Import, Auto Detection only), MHTML, CHM (Import, Auto Detection only), XML (Import, Auto Detection only), JSON (Import, Auto Detection only), MD

### Other formats

- TXT (Import, Export, Auto Detection)

## Platform Independence

GroupDocs.Editor for Node.js via Java does not require any external software or third-party tool to be installed. GroupDocs.Editor for Node.js via Java supports any 32-bit or 64-bit operating system where Node.js is installed. The other details are as follows:

### Supported Operating Systems

- **Microsoft Windows**: Microsoft Windows Desktop (x86, x64) (XP & up), Microsoft Windows Server (x86, x64) (2000 & up), Windows Azure
- **Mac OS**: Mac OS X
- **Linux**: Linux (Ubuntu, OpenSUSE, CentOS, and others)

### Development Environments

- **IDEs**: Microsoft Visual Studio (2010 & up), Xamarin.Android, Xamarin.IOS, Xamarin.Mac, MonoDevelop 2.4 and later
- **Frameworks**: GroupDocs.Editor for Node.js via Java supports Node.js

## Getting Started with GroupDocs.Editor for Node.js via Java

GroupDocs.Editor for Node.js via Java requires J2SE 7.0 (1.7), J2SE 8.0 (1.8), or above. Please install Java first if you do not have it already.

GroupDocs hosts all Java APIs on [GroupDocs Artifact Repository](https://artifact.groupdocs.com/webapp/#/artifacts/browse/tree/General/repo/com/groupdocs/groupdocs-editor), so simply [configure](https://docs.groupdocs.com/editor/java/installation/) your Maven project to fetch the dependencies automatically.

### Installation

From the command line:

```bash
npm install @groupdocs/groupdocs.editor
```

### Run Examples

Change directory to Examples:

```bash
cd Examples
```

Run runExamples.js:

```bash
node runExamples.js
```

## Basic Usage Example

```javascript
const fs = require('fs');
const { Editor, WordProcessingEditOptions, WordProcessingSaveOptions, WordProcessingFormats } = require('@groupdocs/groupdocs.editor');

static async run() {
    const documentPath = Constants.SAMPLE_DOCX;
    console.log(`try to open document ${documentPath}`);
    try {
        const editor = new Editor(documentPath);

        // Obtain editable document from original DOCX document
        const editOptions = new WordProcessingEditOptions();
        const editableDocument = await editor.edit(editOptions); // Open document for editing

        const htmlContent = editableDocument.getEmbeddedHtml();
        // Pass htmlContent to WYSIWYG editor and edit there...
        // ...

        // Save edited EditableDocument object to some WordProcessing format - DOCX for example
        const savePath = Constants.getOutputFilePath("HelloWorldOutput", "docx");
        const saveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Docx);
        await editor.save(editableDocument, savePath, saveOptions);

        console.log(`Document edited and saved successfully, path ${savePath}.`);
    } catch (ex) {
        console.log(ex.message);
    }
}
```

## Set License Example

```javascript
const fs = require('fs');
const { License } = require('@groupdocs/groupdocs.editor');

static run() {
    console.log(`try to set LicenseFromFile, path: ${Constants.LicensePath}`);
    try {
        // Check if the license file exists
        if (fs.existsSync(Constants.LicensePath)) {
            const license = new License();
            // Set license
            license.setLicense(Constants.LicensePath);
            console.log("License set successfully.");
        } else {
            console.log("License file does not exist at the specified path: " + Constants.LicensePath);
        }
    } catch (ex) {
        console.log('Exception: ' + ex.message);
    }
}
```

## Load Document Example

```javascript
const fs = require('fs');
const { Editor, WordProcessingLoadOptions, SpreadsheetLoadOptions } = require('@groupdocs/groupdocs.editor');

const inputPath = Constants.SAMPLE_DOCX;

// Load document as file via path and without load options
const editor1 = new Editor(inputPath);
editor1.dispose();

// Load document as file via path and with load options
const wordLoadOptions = new WordProcessingLoadOptions();
wordLoadOptions.setPassword("some password");
const editor2 = new Editor(inputPath, wordLoadOptions);
editor2.dispose();

// Load document as content from byte stream and without load options
const readStream = fs.createReadStream(Constants.SAMPLE_XLSX);
const stream = await readDataFromStream(readStream);
const editor3 = new Editor(stream);
editor3.dispose();
stream.close();

// Load document as content from byte stream and with load options
const sheetLoadOptions = new SpreadsheetLoadOptions();
sheetLoadOptions.setOptimizeMemoryUsage(true);
const inputStream2 = fs.createReadStream(Constants.SAMPLE_XLSX);
const stream2 = await readDataFromStream(inputStream2);
const editor4 = new Editor(stream2, sheetLoadOptions);
editor4.dispose();
inputStream2.close();
```

[Home](https://www.groupdocs.com/) | [Product Page](https://products.groupdocs.com/editor/java) | [Documentation](https://docs.groupdocs.com/editor/java/) | [Demo](https://products.groupdocs.app/editor/family) | [API Reference](https://reference.groupdocs.com/editor/nodejs-java/) | [Examples](https://github.com/groupdocs-editor/GroupDocs.Editor-for-Java/tree/master/Examples) | [Blog](https://blog.groupdocs.com/category/editor/) | [Search](https://search.groupdocs.com/) | [Free Support](https://forum.groupdocs.com/c/editor) | [Temporary License](https://purchase.groupdocs.com/temporary-license)

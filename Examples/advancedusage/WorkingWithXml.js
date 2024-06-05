const {
    Editor,
    EditableDocument,
    WordProcessingFormats,
    TextualFormats,
    XmlEditOptions,
    XmlHighlightOptions,
    XmlFormatOptions,
    TextSaveOptions,
    WordProcessingSaveOptions,
    ArgbColor,
    FontSize,
    FontStyle,
    FontWeight,
    TextDecorationLineType,
    QuoteType,
    Length, readDataFromStream, StandardCharsets
} = require('@groupdocs/groupdocs.editor');
const Constants = require('../Constants');
const fs = require('fs-extra');
const path = require('path');

class WorkingWithXml {
    static async run() {
        const inputFilePath = Constants.SAMPLE_XML;

        const editor = new Editor(inputFilePath);

        const editOptions = new XmlEditOptions();
        editOptions.setAttributeValuesQuoteType(QuoteType.DoubleQuote);
        editOptions.setRecognizeEmails(true);
        editOptions.setRecognizeUris(true);
        editOptions.setTrimTrailingWhitespaces(true);

        const beforeEdit = await editor.edit(editOptions);

        const originalTextContent = beforeEdit.getContent();
        const updatedTextContent = originalTextContent.replace("John", "Samuel");

        const allResources = beforeEdit.getAllResources();
        const afterEdit = EditableDocument.fromMarkup(updatedTextContent, allResources);

        const wordSaveOptions = new WordProcessingSaveOptions(WordProcessingFormats.Docx);
        const txtSaveOptions = new TextSaveOptions();
        txtSaveOptions.setEncoding(StandardCharsets.UTF_8);

        const outputWordPath = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)), "docx");
        const outputTxtPath = Constants.getOutputFilePath(path.basename(inputFilePath, path.extname(inputFilePath)), "txt");

        await editor.save(afterEdit, outputWordPath, wordSaveOptions);
        await editor.save(afterEdit, outputTxtPath, txtSaveOptions);

        beforeEdit.dispose();
        afterEdit.dispose();
        editor.dispose();

        console.log("WorkingWithXml routine has successfully finished");
    }

    static async loadXml() {
        const xmlInputPath = Constants.SAMPLE_XML;
        const xmlStream = fs.createReadStream(xmlInputPath);
        const stream = await readDataFromStream(xmlStream);
        const editorFromPath = new Editor(xmlInputPath);
        const editorFromStream = new Editor(stream);

        try {
            // Here two Editor instances can separately work with one file
        } finally {
            editorFromStream.dispose();
            editorFromPath.dispose();
        }
    }

    static async editXmlShort() {
        const xmlInputPath = Constants.SAMPLE_XML;
        const outputPath = Constants.getOutputFilePath(path.basename(xmlInputPath, path.extname(xmlInputPath)), "html");

        const editor = new Editor(xmlInputPath);

        try {
            const edited = await editor.edit();
            await edited.save(outputPath);
            edited.dispose();
        } finally {
            editor.dispose();
        }
    }

    static async highlightOptionsDemo() {
        const editOptions = new XmlEditOptions();
        console.assert(editOptions.getHighlightOptions().isDefault());
        const highlightOptions = editOptions.getHighlightOptions();

        highlightOptions.getXmlTagsFontSettings().setSize(FontSize.Large);
        highlightOptions.getXmlTagsFontSettings().setColor(ArgbColor.KnownColors.CssLevel1.Olive);
        highlightOptions.getAttributeNamesFontSettings().setName("Arial");
        highlightOptions.getAttributeNamesFontSettings().setLine(TextDecorationLineType.Underline);
        highlightOptions.getAttributeNamesFontSettings().setWeight(FontWeight.Lighter);
        highlightOptions.getAttributeValuesFontSettings().setLine(TextDecorationLineType.op_Addition(TextDecorationLineType.Underline, TextDecorationLineType.Overline));
        highlightOptions.getAttributeValuesFontSettings().setStyle(FontStyle.Italic);
        highlightOptions.getCDataFontSettings().setLine(TextDecorationLineType.LineThrough);
        highlightOptions.getCDataFontSettings().setSize(FontSize.Smaller);
        highlightOptions.getHtmlCommentsFontSettings().setColor(ArgbColor.KnownColors.CssLevel3.Lightgreen);
        highlightOptions.getHtmlCommentsFontSettings().setName("Courier New");
        highlightOptions.getInnerTextFontSettings().setWeight(FontWeight.fromNumber(300));
        highlightOptions.getInnerTextFontSettings().setSize(FontSize.XSmall);

        console.assert(!editOptions.getHighlightOptions().isDefault());
        highlightOptions.resetToDefault();
        console.assert(editOptions.getHighlightOptions().isDefault());
    }

    static async formatOptionsDemo() {
        const editOptions = new XmlEditOptions();
        console.assert(editOptions.getFormatOptions().isDefault());
        const formatOptions = editOptions.getFormatOptions();

        formatOptions.setEachAttributeFromNewline(true);
        formatOptions.setLeafTextNodesOnNewline(true);
        formatOptions.setLeftIndent(Length.fromValueWithUnit(20, Length.Unit.Px));

        console.assert(!editOptions.getFormatOptions().isDefault());
        formatOptions.setLeftIndent(Length.UnitlessZero);
    }

    static async complexEditDemo() {
        const xmlInputPath = Constants.SAMPLE_XML;

        const outputPath1 = Constants.getOutputFilePath("1--" + path.basename(xmlInputPath, path.extname(xmlInputPath)), "html");
        const outputPath2 = Constants.getOutputFilePath("2--" + path.basename(xmlInputPath, path.extname(xmlInputPath)), "html");

        const editOptions1 = new XmlEditOptions();
        editOptions1.setRecognizeEmails(true);
        editOptions1.setRecognizeUris(true);
        editOptions1.setFixIncorrectStructure(true);
        editOptions1.setAttributeValuesQuoteType(QuoteType.SingleQuote);
        editOptions1.getFormatOptions().setLeftIndent(Length.parse("20px"));
        editOptions1.getHighlightOptions().getXmlTagsFontSettings().setLine(TextDecorationLineType.op_Addition(TextDecorationLineType.Underline, TextDecorationLineType.Overline));
        editOptions1.getHighlightOptions().getXmlTagsFontSettings().setWeight(FontWeight.Bold);

        const editOptions2 = new XmlEditOptions();
        editOptions2.setTrimTrailingWhitespaces(true);
        editOptions2.setAttributeValuesQuoteType(QuoteType.DoubleQuote);
        editOptions2.getFormatOptions().setLeafTextNodesOnNewline(true);
        editOptions2.getHighlightOptions().getXmlTagsFontSettings().setSize(FontSize.XLarge);
        editOptions2.getHighlightOptions().getHtmlCommentsFontSettings().setLine(TextDecorationLineType.LineThrough);

        const editor = new Editor(xmlInputPath);

        try {
            const edited1 = await editor.edit(editOptions1);
            const edited2 = await editor.edit(editOptions2);

            await edited1.save(outputPath1);
            await edited2.save(outputPath2);

            edited1.dispose();
            edited2.dispose();
        } finally {
            editor.dispose();
        }
    }

    static async getXmlMetainfo() {
        const xmlInputPath = Constants.SAMPLE_XML;

        const editor = new Editor(xmlInputPath);
        try {
            const info = await editor.getDocumentInfo(null);
            const xmlInfo = info instanceof TextualDocumentInfo ? info : null;

            if (xmlInfo) {
                console.assert(xmlInfo.getEncoding() === 'UTF-8');
                console.assert(!xmlInfo.isEncrypted());
                console.assert(xmlInfo.getPageCount() === 1);
                console.assert(xmlInfo.getFormat() === TextualFormats.Xml);
            }
        } finally {
            editor.dispose();
        }
    }
}

module.exports = WorkingWithXml;

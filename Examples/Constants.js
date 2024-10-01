const path = require('path');
const fs = require('fs');

class Constants {
    static PROJECT_PATH = process.cwd();
    static SamplesPath = path.join('Resources');
    static OutputPath = path.join(__dirname, '../', 'testResults');
    static LicensePath = (process.env.PATH_TO_LICS + process.env.PRODUCT_LIC) || 'C:\\GroupDocs\\Licenses\\GroupDocs.Editor.NodeJs.lic';
    static SAMPLE_DOCX = Constants.getSampleFilePath('sample.docx');
    static SAMPLE_DOCX2 = Constants.getSampleFilePath('SampleDoc1.docx');
    static SAMPLE_HTML = Constants.getSampleFilePath('SampleDoc1.html');
    static SAMPLE_HTML_BODY = Constants.getSampleFilePath('PureContentSample.html');
    static SAMPLE_HTML_BODY_RESOURCES = Constants.getSampleFilePath('PureContentSample_resources');
    static SAMPLE_CSV = Constants.getSampleFilePath('CarsComma.csv');
    static SAMPLE_XLSX = Constants.getSampleFilePath('sample.xlsx');
    static SAMPLE_XLS_PROTECTED = Constants.getSampleFilePath('Sample_2SpreadSheet_pwd.xlsx');
    static SAMPLE_XML = Constants.getSampleFilePath('SampleXmlCorrect.xml');
    static SAMPLE_TXT = Constants.getSampleFilePath('SamplePlainText1.txt');
    static SAMPLE_PPTX = Constants.getSampleFilePath('sample.pptx');
    static SAMPLE_MSG = Constants.getSampleFilePath('ComplexExample.msg');
    static SAMPLE_MD = Constants.getSampleFilePath(path.join('Markdown', 'ComplexImage.md'));
    static SAMPLE_MD_EDIT = Constants.getSampleFilePath(path.join('Markdown', 'input.md'));
    static SAMPLE_MD_FOLDER = Constants.getSampleFilePath(path.join('Markdown', 'images'));

    static getSampleFilePath(fileName) {
        return path.join(Constants.PROJECT_PATH, Constants.SamplesPath, fileName);
    }

    static getOutputDirectoryPath(callerFilePath) {
        const outputDir = path.join(Constants.OutputPath, callerFilePath);
        Constants.ensureDirectoryExists(outputDir);
        return outputDir;
    }

    static getOutputFilePath(fileName, fileExtension) {
        const outputDir = Constants.OutputPath;
        Constants.ensureDirectoryExists(outputDir);
        return path.join(outputDir, `${fileName}.${fileExtension}`);
    }

    static ensureDirectoryExists(directoryPath) {
        if (!fs.existsSync(directoryPath)) {
            fs.mkdirSync(directoryPath, { recursive: true });
        }
    }

    static removeExtension(filename) {
        return path.parse(filename).name;
    }
}

module.exports = Constants;

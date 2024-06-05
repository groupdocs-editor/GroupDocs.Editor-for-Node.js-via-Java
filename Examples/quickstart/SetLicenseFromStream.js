const fs = require('fs');
const Constants = require('../Constants');
const {License, readDataFromStream} = require('@groupdocs/groupdocs.editor'); // Assuming 'groupdocs-editor' is the name of the package


/**
 * This example demonstrates how to set a license from a stream.
 */
class SetLicenseFromStream {
    static async run() {
        try {
            // Check if the license file exists
            if (fs.existsSync(Constants.LicensePath)) {
                // Read the license file into a stream
                const fileStream = fs.createReadStream(Constants.LicensePath);
                const stream = await readDataFromStream(fileStream);
                const license = new License();
                // Set license from the stream
                license.setLicense(stream);
                fileStream.close();
                console.log("License set successfully from stream.");
            } else {
                console.log("License file does not exist at the specified path: " + Constants.LicensePath);
            }
        } catch (ex) {
            console.log('Exception: ' + ex.message);
        }
    }
}

module.exports = SetLicenseFromStream;

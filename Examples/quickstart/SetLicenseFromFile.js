const fs = require('fs');
const Constants = require('../Constants');
const { License } = require('@groupdocs/groupdocs.editor'); // Assuming 'groupdocs-editor' is the name of the package

/**
 * This example demonstrates how to set a license from a file.
 *
 * The setLicense method attempts to set a license from several locations
 * relative to the executable and GroupDocs.Editor.dll. You can also use the
 * additional overload to load a license from a stream, this is useful for
 * instance when the License is stored as an embedded resource.
 */
class SetLicenseFromFile {
    static run() {
        console.log(`try to set LicenseFromFile, path: ${Constants.LicensePath}`);
        try {
            // Check if the license file exists
            if (fs.existsSync(Constants.LicensePath)) {
                // ExStart:applyLicense
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
}

module.exports = SetLicenseFromFile;

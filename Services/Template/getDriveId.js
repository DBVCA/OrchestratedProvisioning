var request = require('request');

module.exports = function getDriveId(context, token, siteId,
    libraryName) {

    return new Promise((resolve, reject) => {

        // Note: $filter does not work so get all the libraries
        const url = `https://graph.microsoft.com/v1.0/sites/` +
                    `${siteId}/drives`;
        try {
		resolve("b!hBTZLDvpDU-zHKgrl20pV08OSAbEnDFAmbSr3NgwNMqT5gOX5dT0RbA9I7QjAryZ");
			context.log('Library: ' + libraryName);
			context.log('URL: ' + url);
            request.get(url, {
                'auth': {
                    'bearer': token
                }
            }, (error, response, body) => {

                if (!error && response && response.statusCode == 200) {

                    const result = JSON.parse(response.body);

                    if (result.value && result.value[0]) {
                        const library = result.value.find((l) =>
                            l.name.toLowerCase() === libraryName.toLowerCase());
                        if (library) {
                            resolve(library.id);
                        } else {
                            reject("Drive ${libraryName} not found in getDriveId");
                        }
                    } else {
                        reject("No drives found in getDriveId");
                    }

                } else {

                    if (error) {
                        reject(`Error in getDriveId: ${error}`);
                    } else {
                        let b = JSON.parse(response.body);
                        reject(`Error ${b.error.code} in getDriveId: ${b.error.message}`);
                    }
                    
                }
            });
            
        } catch (ex) {
            reject(`Error in getDriveId: ${ex}`);
        }

    });
}

var getSiteId = require('./getSiteId');
var getDriveId = require('./getDriveId');
var getDownloadUrl = require('./getDownloadUrl');
var getUserId = require('./getUserId');
var downloadDriveItem = require('./downloadDriveItem');
var settings = require('../Settings/settings');

module.exports = function getTemplate(context, token, jsonTemplate,
    displayName, description, owner) {

    context.log('Getting template ' + jsonTemplate);
	if( jsonTemplate == "GreenTemplate")
	{
		
	    return new Promise((resolve, reject) => {

		var template;   // The template as a JavaScript object
	
		context.log(settings().TENANT_NAME + ' ' + settings().TEMPLATE_SITE_URL + ' ' + settings().TEMPLATE_LIB_NAME);
		// 1. Get ID of the SharePoint site where template files are stored
		getSiteId(context, token, settings().TENANT_NAME, 
		    settings().TEMPLATE_SITE_URL)
		.then((siteId) => {
		// 2. Get the Graph API drive ID for the doc library where template files are stored
		     return getDriveId(context, token, siteId, settings().TEMPLATE_LIB_NAME);
		})
		.then((driveId) => {
		// 3. Get the download URL for the template file
		    return getDownloadUrl(context, token, driveId,
			`${jsonTemplate}${settings().TEMPLATE_FILE_EXTENSION}`);
		})
		.then((downloadUrl) => {
		// 4. Get the contents of the template file
		    return downloadDriveItem(context, token, downloadUrl);
		})
		.then((templateString) => {

		// 5. Parse the template; get owner's user ID
		    template = JSON.parse(templateString.trimLeft());
		    return getUserId (context, token, owner);
		})
	
		
		.then((ownerId) => {
		// 6. Add the per-team properties to the template  '${ownerId}'
			nikiID= getUserId (context,token,'nmorejon@vca-green.com');
			dbvmid = getUserId (context, token, 'dbvm@vcastructural.com');
			dbvm2id = getUserId (context, token, 'DBVM.2@vcastructural.com');
			 context.log(dbvmid);
			context.log(dbvm2id);
			context.log(${nikiID});
			context.log(ownerId);
			context.log('${ownerId}');
			//context.log(${ownerId});
		    template['displayName'] = displayName;
		    template['description'] = description;
		    template['owners@odata.bind'] = [
			`https://graph.microsoft.com/beta/users('`+dbvm2id+`')`
		    ];
		    
		template['channels'][0]['displayName'] = 'Design Phase-' + displayName;https://github.com/DBVCA/OrchestratedProvisioning/blob/dev/Services/Template/getTemplate.js
		template['channels'][1]['displayName'] = 'Energy Modeling-' + displayName;
		template['channels'][2]['displayName'] = 'Field Services-' + displayName;
			template['membersodata.bind'] = [`https://graph.microsoft.com/beta/users('${ownerId}')`];
//		    template['channels'] = [
//				{
//				    'displayName': 'Design Phase-${displayname}',
//				    'isFavoriteByDefault': true,
//				    'description': ''
//				}
//			    ];
		// 7. Return the finished template as a string
		    resolve(JSON.stringify(template));
		})
		.catch((ex) => {
		    reject(`Error in getTemplate(): ${ex}`);
		});
	});	
	}
	else
	{
		return new Promise((resolve, reject) => {

		var template;   // The template as a JavaScript object
	
		context.log(settings().TENANT_NAME + ' ' + settings().TEMPLATE_SITE_URL + ' ' + settings().TEMPLATE_LIB_NAME);
		// 1. Get ID of the SharePoint site where template files are stored
		getSiteId(context, token, settings().TENANT_NAME, 
		    settings().TEMPLATE_SITE_URL)
		.then((siteId) => {
		// 2. Get the Graph API drive ID for the doc library where template files are stored
		     return getDriveId(context, token, siteId, settings().TEMPLATE_LIB_NAME);
		})
		.then((driveId) => {
		// 3. Get the download URL for the template file
		    return getDownloadUrl(context, token, driveId,
			`${jsonTemplate}${settings().TEMPLATE_FILE_EXTENSION}`);
		})
		.then((downloadUrl) => {
		// 4. Get the contents of the template file
		    return downloadDriveItem(context, token, downloadUrl);
		})
		.then((templateString) => {

		// 5. Parse the template; get owner's user ID
		    template = JSON.parse(templateString.trimLeft());
		    return getUserId (context, token, owner);
		})
	
		
		.then((ownerId) => {
		// 6. Add the per-team properties to the template

		    template['displayName'] = displayName;
		    template['description'] = description;
		    template['owners@odata.bind'] = [
			`https://graph.microsoft.com/beta/users('${ownerId}')`
		    ];
		// 7. Return the finished template as a string
		    resolve(JSON.stringify(template));
		})
		.catch((ex) => {
		    reject(`Error in getTemplate(): ${ex}`);
		});
		});
	}
        


    
}

//TEST UPDATE
const settings = require('../Services/Settings/settings');
const getToken = require('../Services/Token/getToken');
const getTemplate = require('../Services/Template/getTemplate');
const createTeam = require('../Services/Team/createTeam');
const getChannelId = require('../Services/Team/getChannelId');
const url = `https://graph.microsoft.com/beta/teams`;

module.exports = async function (context, myQueueItem) {

    var token = "";

    return new Promise((resolve, reject) => {

        context.log('JavaScript queue trigger function processed work item', myQueueItem);

        if (myQueueItem && 
            myQueueItem.displayName &&
            myQueueItem.owner && 
            myQueueItem.jsonTemplate) {

            try {

                const displayName = myQueueItem.displayName || "New team";
                const description = myQueueItem.description || "";
                const owner = myQueueItem.owner;
                const requestId = myQueueItem.requestId || "";
                const jsonTemplate = myQueueItem.jsonTemplate;
                var newTeamId;
                var CadTeamChannelID;

                context.log(`Creating Team ${displayName} using ${jsonTemplate} json template`);
                context.log(settings().TENANT_NAME);
                context.log(settings().TEMPLATE_SITE_URL);

                getToken(context)
                .then ((accessToken) => {
                    context.log(`Got access token of ${accessToken.length} characters`);
                    token = accessToken;
                    return getTemplate(context, token, jsonTemplate,
                        displayName, description, owner);
                })
                .then((templateString) => {
                    return createTeam(context, token, templateString);
                })
                .then((teamId) => {
                    newTeamId = teamId;
                    context.log(`createTeam created team ${newTeamId}`);
                    return getChannelId(context, token, newTeamId, "General");
                })
                .then((channelId) => {
                    //newChannelID = channelId;
                      
                    /*cadChannelId = getChannelId(context, token, newTeamId, "CAD Team");
                    
                                const Channelurl = `https://graph.microsoft.com/beta/teams/${newTeamID}/channels/${cadChannelId}/filesFolder`;
                                context.log("Channel URL:");
                                context.log(Channelurl);
                                request.post(Channelurl, {
                                  'name': 'Setup',
                                  'folder': { },
                                  '@microsoft.graph.conflictBehavior': 'fail'
                                }, (error, response, body) => {

                              context.log(`Received a response with status code ${response.statusCode} error=${error}`);

                              if (response && response.statusCode == 202) {

                                // If here we successfully issued the request
                                const opUrl = `https://graph.microsoft.com/beta${response.headers.location}`;
                                context.log(`operation url is ${opUrl}`);

                                pollUntilDone(resolve, reject, opUrl, token, NUMBER_OF_RETRIES);

                              } else {

                                context.log(`Exception path response ${response.statusCode}`);
                                // If here something went wrong, reject with an error
                                // message
                                if (error) {
                                  reject(error);
                                } else {
                                  let b = JSON.parse(response.body);
                                  reject(`${b.error.code} - ${b.error.message}`);
                                }

                              }
                            });*/

                    context.bindings.myOutputQueueItem = {
                        success: true,
                        requestId: requestId,
                        teamId: newTeamId,
                        teamUrl: `https://teams.microsoft.com/l/team/${channelId}/conversations?groupId=${newTeamId}&tenantId=${settings().TENANT}`,
                        teamName: displayName,
                        teamDescription: description,
                        owner: owner,
                        error: ''
                    };
                    
                    resolve();
                })
                
                .catch((error) => {
                    context.log(`ERROR: ${error}`);
                    context.bindings.myOutputQueueItem = {
                        success: false,
                        requestId: requestId,
                        teamId: '',
                        teamUrl: '',
                        teamName: displayName,
                        teamDescription: description,
                        owner: owner,
                        error: error
                    };
                    resolve();
                });

            } catch (ex) {
                context.log(`Error: ${ex}`);
                context.bindings.myOutputQueueItem = {
                    success: false,
                    requestId: requestId,
                    teamId: '',
                    teamUrl: '',
                    teamName: displayName,
                    teamDescription: description,
                    owner: owner,
                    error: ex
                };
                resolve();
            }
        } else {
            context.log('Skipping empty or invalid queue entry');
        }

    });

};

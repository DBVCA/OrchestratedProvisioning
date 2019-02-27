var getTeamId = require('./getTeamId');
var postClone = require('./postClone');
var getToken = require('./getToken');

module.exports = async function (context, myQueueItem) {

    var token = "";

    return new Promise((resolve, reject) => {

        context.log('JavaScript queue trigger function processed work item', myQueueItem);

        if (myQueueItem && myQueueItem.oldTeam && myQueueItem.newTeam) {

            try {

                var oldTeam = myQueueItem.oldTeam;
                var newTeam = myQueueItem.newTeam;
                context.log(`Cloning ${oldTeam} to ${newTeam}`);

                let token = context.bindings.graphToken;

                getToken(context)
                    .then((accessToken) => {
                        context.log(`Got access token of ${accessToken.length} characters`);
                        token = accessToken;
                        return getTeamId(context, token, oldTeam);
                    })
                    .then((teamId) => {
                        context.log(`Got team ID ${teamId}`);
                        return postClone(context, token, teamId, newTeam);
                    })
                    .then((newTeamId) => {
                        context.log(`postCreate created team ${newTeamId}`);
                        context.bindings.myOutputQueueItem = [newTeamId];
                        resolve();
                    })
                    .catch((error) => {
                        context.log(`ERROR: ${error}`);
                        reject(error);
                    })

            } catch (ex) {
                context.log(`Error: ${ex}`);
                reject(ex);
            }
        } else {
            context.log('Skipping empty queue entry');
        }

    });

};
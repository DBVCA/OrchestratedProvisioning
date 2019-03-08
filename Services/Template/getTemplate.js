var request = require('request');

// getTemplate() - Return a template used to create a Team
module.exports = function getTemplate(context, token, jsonTemplate,
    displayName, description, owner) {

    context.log('Getting template ' + jsonTemplate);

    return new Promise((resolve, reject) => {

        const template = `
        {
            "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
            "displayName": "${displayName}",
            "description": "${description}",
            "owners@odata.bind": [
                "https://graph.microsoft.com/beta/users('${owner}')"
            ],
            "visibility": "Private",
            "channels": [
                {
                    "displayName": "Announcements 📢",
                    "isFavoriteByDefault": true,
                    "description": "This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements."
                },
                {
                    "displayName": "Training 🏋️",
                    "isFavoriteByDefault": true,
                    "description": "This is a sample training channel, that is favorited by default, and contains an example of pinned website and YouTube tabs.",
                    "tabs": [
                        {
                            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.web')",
                            "name": "A Pinned Website",
                            "configuration": {
                                "contentUrl": "https://docs.microsoft.com/en-us/microsoftteams/microsoft-teams"
                            }
                        },
                        {
                            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.youtube')",
                            "name": "A Pinned YouTube Video",
                            "configuration": {
                                "contentUrl": "https://tabs.teams.microsoft.com/Youtube/Home/YoutubeTab?videoId=X8krAMdGvCQ",
                                "websiteUrl": "https://www.youtube.com/watch?v=X8krAMdGvCQ"
                            }
                        }
                    ]
                },
                {
                    "displayName": "Planning 📅 ",
                    "description": "This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.",
                    "isFavoriteByDefault": false
                },
                {
                    "displayName": "Issues and Feedback 🐞",
                    "description": "This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu."
                }
            ],
            "memberSettings": {
                "allowCreateUpdateChannels": true,
                "allowDeleteChannels": true,
                "allowAddRemoveApps": true,
                "allowCreateUpdateRemoveTabs": true,
                "allowCreateUpdateRemoveConnectors": true
            },
            "guestSettings": {
                "allowCreateUpdateChannels": false,
                "allowDeleteChannels": false
            },
            "funSettings": {
                "allowGiphy": true,
                "giphyContentRating": "Moderate",
                "allowStickersAndMemes": true,
                "allowCustomMemes": true
            },
            "messagingSettings": {
                "allowUserEditMessages": true,
                "allowUserDeleteMessages": true,
                "allowOwnerDeleteMessages": true,
                "allowTeamMentions": true,
                "allowChannelMentions": true
            },
            "installedApps": [
                {
                    "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.vsts')"
                },
                {
                    "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('1542629c-01b3-4a6d-8f76-1938b779e48d')"
                }
            ]
        }
        `
        resolve(template);

    });
}
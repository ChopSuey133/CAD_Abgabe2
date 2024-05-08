//module.exports = {};

require('isomorphic-fetch');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new azure.DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
  });

  const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    _deviceCodeCredential, {
      scopes: settings.graphUserScopes
    });

  _userClient = graph.Client.initWithMiddleware({
    authProvider: authProvider
  });
}
module.exports.initializeGraphForUserAuth = initializeGraphForUserAuth;

async function getUserTokenAsync() {
    // Ensure credential isn't undefined
    if (!_deviceCodeCredential) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    // Ensure scopes isn't undefined
    if (!_settings?.graphUserScopes) {
      throw new Error('Setting "scopes" cannot be undefined');
    }
  
    // Request token with given scopes
    const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
    return response.token;
  }
  module.exports.getUserTokenAsync = getUserTokenAsync;
  
  async function getUserAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    return _userClient.api('/me')
      // Only request specific properties
      .select(['displayName', 'mail', 'userPrincipalName'])
      .get();
  }
  module.exports.getUserAsync = getUserAsync;



  async function getInboxAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    return _userClient.api('/me/mailFolders/inbox/messages')
      .select(['from', 'isRead', 'receivedDateTime', 'subject'])
      .top(25)
      .orderby('receivedDateTime DESC')
      .get();
  }
  module.exports.getInboxAsync = getInboxAsync;

  async function sendMailAsync(subject, body, recipient) {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    // Create a new message
    const message = {
      subject: subject,
      body: {
        content: body,
        contentType: 'text'
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipient
          }
        }
      ]
    };
  
    // Send the message
    return _userClient.api('me/sendMail')
      .post({
        message: message
      });
  }
  module.exports.sendMailAsync = sendMailAsync;

// This function serves as a playground for testing Graph snippets
// or other code
async function makeGraphCallAsync() {
    // Ensure the Graph client is initialized
    if (!_userClient) {
        throw new Error('Graph has not been initialized for user auth');
    }

    // Construct a new date object for the current date/time
    const today = new Date();
    // Set the end date to 7 days from today
    const nextWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7);

    // Format dates as ISO 8601 strings (YYYY-MM-DD) for the query
    const startDate = today.toISOString().split('T')[0];
    const endDate = nextWeek.toISOString().split('T')[0];

    try {
        // Fetch calendar view within a specific date range
        const response = await _userClient.api(`/me/calendarView?startDateTime=${startDate}&endDateTime=${endDate}`)
            .select('subject,organizer,start,end')
            .orderby('start/dateTime')
            .get();

        // Check and display each event detail
        if (response.value.length > 0) {
            console.log('Upcoming events:');
            response.value.forEach(event => {
                console.log(`Subject: ${event.subject}`);
                console.log(`Organizer: ${event.organizer.emailAddress.name}`);
                console.log(`Start: ${event.start.dateTime}`);
                console.log(`End: ${event.end.dateTime}`);
            });
        } else {
            console.log('No upcoming events found.');
        }
    } catch (err) {
        console.log(`Error accessing calendar events: ${err}`);
    }
}
 module.exports.makeGraphCallAsync = makeGraphCallAsync;
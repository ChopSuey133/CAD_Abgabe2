const readline = require('readline-sync');

const settings = require('./appSettings');
const graphHelper = require('./graphHelper');

async function main() {
  console.log('JavaScript Graph Tutorial');

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  // Greet the user by name
  await greetUserAsync();

  const choices = [
    'Display access token (displayAccessTokenAsync)',
    'List my inbox (listInboxAsync)',
    'Send mail (sendMailAsync)',
    'Display 7-day calender entries (makeGraphCallAsync, Custom Function)'
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        // Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List emails from user's inbox
        await listInboxAsync();
        break;
      case 2:
        // Send an email message
        await sendMailAsync();
        break;
      case 3:
        // Run any Graph code
        await makeGraphCallAsync();
        break;
      default:
        console.log('Invalid choice! Please try again.');
    }
  }
}
main();


function initializeGraph(settings) {
    graphHelper.initializeGraphForUserAuth(settings, (info) => {
      // Display the device code message to
      // the user. This tells them
      // where to go to sign in and provides the
      // code to use.
      console.log(info.message);
    });
  }
  
  async function greetUserAsync() {
    try {
      const user = await graphHelper.getUserAsync();
      console.log(`Hello, ${user?.displayName}!`);
      // For Work/school accounts, email is in mail property
      // Personal accounts, email is in userPrincipalName
      console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ''}`);
    } catch (err) {
      console.log(`Error getting user: ${err}`);
    }
  }
  
  async function displayAccessTokenAsync() {
    try {
      console.log(`In function displayAccessTokenAsync()`);
      const userToken = await graphHelper.getUserTokenAsync();
      console.log(`User token: ${userToken}`);
    } catch (err) {
      console.log(`Error getting user access token: ${err}`);
    }
  }
  
  async function listInboxAsync() {
    try {
      const messagePage = await graphHelper.getInboxAsync();
      const messages = messagePage.value;
  
      // Output each message's details
      for (const message of messages) {
        console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
        console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
        console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
        console.log(`  Received: ${message.receivedDateTime}`);
      }
  
      // If @odata.nextLink is not undefined, there are more messages
      // available on the server
      const moreAvailable = messagePage['@odata.nextLink'] != undefined;
      console.log(`\nMore messages available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting user's inbox: ${err}`);
    }
  }
  
  async function sendMailAsync() {
    try {
      // Send mail to the signed-in user
      // Get the user for their email address
      const user = await graphHelper.getUserAsync();
      const userEmail = user?.mail ?? user?.userPrincipalName;
  
      if (!userEmail) {
        console.log('Couldn\'t get your email address, canceling...');
        return;
      }
  
      await graphHelper.sendMailAsync('Testing Microsoft Graph',
        'Hello world!', userEmail);
      console.log('Mail sent.');
    } catch (err) {
      console.log(`Error sending mail: ${err}`);
    }
  }
  
// This function serves as a playground for testing Graph snippets
// or other code
async function makeGraphCallAsync() {
    try {
      await graphHelper.makeGraphCallAsync();
    } catch (err) {
      console.log(`Error making Graph call: ${err}`);
    }
}
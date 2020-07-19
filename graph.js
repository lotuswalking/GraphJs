// Create an options object with the same scopes from the login
const options =
  new MicrosoftGraph.MSALAuthenticationProviderOptions([
    'user.read',
    'calendars.read',
    'Presence.Read',
    'Presence.Read.All',
    'Chat.Read'
  ]);
// Create an authentication provider for the implicit flow
const authProvider =
  new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalClient, options);
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});

// 异步操作,显示状态信息
async function getPresence() {
    try {
        let presence = await graphClient
        .api('/me/presence')
        .version('beta')
        // .select('userPrincipalName, id')
        // .select('availability')
        .get();
        updatePage(msalClient.getAccount(), Views.presence, presence)        
    }
    catch(error)
    {
        console.log(error);
    }
}
// 显示邮件
async function getEmails() {
    try {
        // console.log("###########")
        let mails = await graphClient
        .api('/me/messages')
        // .select('subject,from,receivedDateTime,id')
        .orderby('createdDateTime DESC')
        .get();
        // console.log("77777777")
        updatePage(msalClient.getAccount(), Views.mail, mails);
    }catch (error) {
        updatePage(msalClient.getAccount(), Views.error, {
          message: 'Error getting events',
          debug: error
        });
}
}
async function getEmailDetail(id)
{
    try {
        // console.log("###########")
        let mailRead = await graphClient
        .api('/me/messages/'+id)
        // .select('subject,from,receivedDateTime,id')
        // .orderby('createdDateTime DESC')
        .get();
        // console.log("77777777")
        updatePage(msalClient.getAccount(), Views.mailRead, mailRead);
    }catch (error) {
        updatePage(msalClient.getAccount(), Views.error, {
          message: 'Error getting events',
          debug: error
        });
}
}
//获取其他人的状态信息
async function getPresenceByEmail(emailer)
{
    console.log()
    try {
        // console.log("###########")
        var emailer = document.getElementById('mailAddr').value
        let userProfile = await graphClient
        .api('/users/'+emailer)
        // .select('subject,from,receivedDateTime,id')
        // .orderby('createdDateTime DESC')
        .get();
        var userId = userProfile.id;
        // console.log(userId);
        let userPresence = await graphClient
        .api('/users/'+userId+'/presence')
        .version('beta')
        .get();
        // console.log(userPresence);
        showOtherPresence(emailer, userPresence);
        //post a message to user
        let chatMessage = {
            "body": {
                "content": "Hello, this is a message from Junyan's Bot!"
            }
        }
        let res = await graphClient
        .api('/chats/'+userId+'/messages')
        .version('beta')
        .post(chatMessage);
        console.log(res);

        
    }catch (error) {
        updatePage(msalClient.getAccount(), Views.error, {
          message: 'Error getting events',
          debug: error
        });
}
}
// 显示日历项
async function getEvents() {
    try {
      let events = await graphClient
          .api('/me/events')
          .select('subject,organizer,start,end')
          .orderby('createdDateTime DESC')
          .get();
  
      updatePage(msalClient.getAccount(), Views.calendar, events);
    } catch (error) {
      updatePage(msalClient.getAccount(), Views.error, {
        message: 'Error getting events',
        debug: error
      });
    }
  }
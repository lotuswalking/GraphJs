/**
 * The code below demonstrates how you can use MSAL as a custom authentication provider for the Microsoft Graph JavaScript SDK.
 * You do NOT need to implement a custom provider. Microsoft Graph JavaScript SDK v3.0 (preview) offers AuthCodeMSALBrowserAuthenticationProvider
 * which handles token acquisition and renewal for you automatically. For more information on how to use it, visit:
 * https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/dev/docs/AuthCodeMSALBrowserAuthenticationProvider.md
 */

/**
 * Returns a graph client object with the provided token acquisition options
 * @param {Object} providerOptions: object containing user account, required scopes and interaction type
 */
const getGraphClient = (providerOptions) => {
  /**
   * Pass the instance as authProvider in ClientOptions to instantiate the Client which will create and set the default middleware chain.
   * For more information, visit: https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/dev/docs/CreatingClientInstance.md
   */
  let clientOptions = {
    authProvider: new MsalAuthenticationProvider(providerOptions),
  };

  const graphClient = MicrosoftGraph.Client.initWithMiddleware(clientOptions);

  return graphClient;
};
function getEmails() {
  getGraphClient({
    account: localAccount,
    scopes: graphConfig.graphMailEndpoint.scopes,
    interactionType: msal.InteractionType.Popup,
  })
    .api("/me/messages")
    .get()
    .then((response) => {
      return updatePage(localAccount, Views.mail, response);
    })
    .catch((error) => {
      console.log(error);
    });
}
function getEvents() {
  getGraphClient({
    account: localAccount,
    scopes: graphConfig.graphEventEndpoint.scopes,
    interactionType: msal.InteractionType.Popup,
  })
    .api("/me/events")
    .get()
    .then((response) => {
      return updatePage(localAccount, Views.calendar, response);
    })
    .catch((error) => {
      console.log(error);
    });
}

async function getPresence() {
  getGraphClient({
    account: localAccount,
    scopes: graphConfig.graphPresenceEndpoint.scopes,
    interactionType: msal.InteractionType.Popup,
  })
    .api("/me/presence")
    .version("beta")
    .get()
    .then((presence) => {
      updatePage(localAccount, Views.presence, presence);
    })
    .catch((error) => {
      console.log(error);
    });
}
function getPresenceByEmail(emailer) {
  //get user Profile  graphUsersEndpoint
  getGraphClient({
    account: localAccount,
    scopes: graphConfig.graphUsersEndpoint.scopes,
    interactionType: msal.InteractionType.Popup,
  })
    .api("/users/" + emailer)
    .get()
    .then((user) => {
      getPresenceById(emailer, user.id);
    })
    .catch((error) => {
      console.log(error);
    });
}
function getPresenceById(emailer, userId) {
  getGraphClient({
    account: localAccount,
    scopes: graphConfig.graphUsersEndpoint.scopes,
    interactionType: msal.InteractionType.Popup,
  })
    .api("/users/" + userId + "/presence")
    .version("beta")
    .get()
    .then((userPresence) => {
      showOtherPresence(emailer, userPresence);
      //   sendMessage(userId);
      // updatePage(localAccount, Views.presence, presence)
    })
    .catch((error) => {
      console.log(error);
    });
}
function sendMessage(userId) {
  let chatMessage = {
    body: {
      content: "Hello, this is a message from Junyan's Bot!",
    },
  };
  getGraphClient({
    account: localAccount,
    scopes: graphConfig.graphChatEndpoint.scopes,
    interactionType: msal.InteractionType.Popup,
  })
    .api("/chats/" + userId + "/messages")
    .version("beta")
    .post(chatMessage)
    .then((data) => {
      console.log(data);
      // updatePage(localAccount, Views.presence, presence)
    })
    .catch((error) => {
      console.log(error);
    });
}

function updatePage(account, view, data) {
  //    如果没有账号登录或者View的入口参数为空
  if (!view || !account) {
    view = Views.home;
  }
  //   显示右上角登录小下拉条
  showAccountNav(account);

  // 显示Email/Calendar/Presence等按钮
  showAuthenticatedNav(account, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(account);
      break;
    case Views.calendar:
      showCalendar(data);
      break;
    case Views.mail:
      showEmail(data);
      break;
    case Views.presence:
      showPresence(data);
      break;
    case Views.mailRead:
      ShowEmailDetail(data);
      break;
  }
}

// 显示邮件
function showEmail(emails) {
  try {
    // console.log(emails);
    var div = document.createElement("div");
    div.appendChild(createElement("h1", null, "Email"));
    var table = createElement("table", "table");
    div.appendChild(table);

    var thead = document.createElement("thead");
    table.appendChild(thead);

    var headerRow = document.createElement("tr");
    thead.appendChild(headerRow);

    var subject = createElement("th", null, "From");
    subject.setAttribute("scope", "col");
    headerRow.appendChild(subject);

    var subject = createElement("th", null, "Subject");
    subject.setAttribute("scope", "col");
    headerRow.appendChild(subject);

    var receivedDateTime = createElement("th", null, "receivedDateTime");
    receivedDateTime.setAttribute("scope", "col");
    headerRow.appendChild(receivedDateTime);
    var tbody = document.createElement("tbody");
    table.appendChild(tbody);

    for (const mail of emails.value) {
      var mailRow = document.createElement("tr");
      var mailRow = document.createElement("tr");
      mailRow.setAttribute("key", mail.id);
      mailRow.setAttribute("onclick", 'getEmailDetail("' + mail.id + '")');
      tbody.appendChild(mailRow);

      var fromCell = createElement("td", null, mail.from.emailAddress.address);
      mailRow.appendChild(fromCell);
      var subjectCell = createElement("td", null, mail.subject);
      //   subjectCell.setAttribute('onclick','ShowEmailDetail(this)');

      mailRow.appendChild(subjectCell);
      var startCell = createElement("td", null, mail.receivedDateTime);
      mailRow.appendChild(startCell);
    }

    mainContainer.innerHTML = "";
    mainContainer.appendChild(div);
  } catch (ex) {
    console.log(ex);
  }
}

/**
 * This class implements the IAuthenticationProvider interface, which allows a custom authentication provider to be
 * used with the Graph client. See: https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/dev/src/IAuthenticationProvider.ts
 */
class MsalAuthenticationProvider {
  account; // user account object to be used when attempting silent token acquisition
  scopes; // array of scopes required for this resource endpoint
  interactionType; // type of interaction to fallback to when silent token acquisition fails

  constructor(providerOptions) {
    this.account = providerOptions.account;
    this.scopes = providerOptions.scopes;
    this.interactionType = providerOptions.interactionType;
  }

  /**
   * This method will get called before every request to the ms graph server
   * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
   * Basically this method will contain the implementation for getting and refreshing accessTokens
   */
  getAccessToken() {
    return new Promise(async (resolve, reject) => {
      let response;

      try {
        response = await myMSALObj.acquireTokenSilent({
          account: this.account,
          scopes: this.scopes,
        });

        if (response.accessToken) {
          resolve(response.accessToken);
        } else {
          reject(Error("Failed to acquire an access token"));
        }
      } catch (error) {
        // in case if silent token acquisition fails, fallback to an interactive method
        if (error instanceof msal.InteractionRequiredAuthError) {
          switch (this.interactionType) {
            case msal.InteractionType.Popup:
              response = await myMSALObj.acquireTokenPopup({
                scopes: this.scopes,
              });

              if (response.accessToken) {
                resolve(response.accessToken);
              } else {
                reject(Error("Failed to acquire an access token"));
              }
              break;

            case msal.InteractionType.Redirect:
              /**
               * This will cause the app to leave the current page and redirect to the consent screen.
               * Once consent is provided, the app will return back to the current page and then the
               * silent token acquisition will succeed.
               */
              myMSALObj.acquireTokenRedirect({
                scopes: this.scopes,
              });
              break;

            default:
              break;
          }
        }
      }
    });
  }
}

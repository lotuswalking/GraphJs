const msalConfig = {
    auth: {
      clientId: '80c1687b-440a-4dd6-811a-8c1d76be8129',
      authority: "https://login.microsoftonline.com/5c7d0b28-bdf8-410c-aa93-4df372b16203",
      redirectUri: 'https://lotuswalking.github.io/'
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false
    }
  };
  // Add here the endpoints for MS Graph API services you would like to use.
const graphConfig = {
  graphMeEndpoint: {
      uri: "https://graph.microsoft.com/v1.0/me",
      scopes: ["User.Read"]
  },
  graphMailEndpoint: {
      uri: "https://graph.microsoft.com/v1.0/me/messages",
      scopes: ["Mail.Read"]
  }
};

/**
* Scopes you add here will be prompted for user consent during sign-in.
* By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
* For more information about OIDC scopes, visit: 
* https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
*/
const loginRequest = {
  scopes: ["User.Read"]
};

// exporting config object for jest
if (typeof exports !== 'undefined') {
  module.exports = {
      msalConfig: msalConfig,
  };
}
const msalConfig = {
    auth: {
      clientId: '80c1687b-440a-4dd6-811a-8c1d76be8129',
      redirectUri: 'https://lotuswalking.github.io/'
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
      forceRefresh: false
    }
  };
  
  const loginRequest = {
    scopes: [
      'openid',
      'profile',
      'user.read',
      'mailboxsettings.read', 
      'calendars.read',
      'Presence.Read',
      'Presence.Read.All',
      'Chat.Read'
    ]
  }

const msalConfig = {
    auth: {
      clientId: '80c1687b-440a-4dd6-811a-8c1d76be8129',
      redirectUri: 'http://localhost'
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
      'calendars.read',
      'Presence.Read',
      'Presence.Read.All',
      'Chat.Read'
    ]
  }
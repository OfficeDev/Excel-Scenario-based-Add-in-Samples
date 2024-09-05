window.addEventListener('load', function () {
  let params = new URLSearchParams(location.search);
  let userClientId = params.get('data');
  if (userClientId != null) {
    localStorage.setItem('userClientId', userClientId);
  }
  else {
    userClientId = localStorage.getItem('userClientId');
  }

  const msalUrl = location.href.substring(0, location.href.lastIndexOf('/')) + '/consent.html';
  const msalClient = new msal.PublicClientApplication({
    auth: {
      clientId: userClientId,
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: msalUrl // Must be registered as "spa" type
    },
    cache: {
      cacheLocation: 'localStorage', // needed to avoid "login required" error
      storeAuthStateInCookie: true   // recommended to avoid certain IE/Edge issues
    }
  });

  // handleRedirectPromise should be invoked on every page load
  msalClient.handleRedirectPromise()
    .then(response => {
      // If response is non-null, it means page is returning from AAD with a successful response
      if (response) {
        // Check if the origin of the request is same as the registered redirect URI
        if (new URL(msalUrl).origin === window.location.origin) {
          window.opener.postMessage({ status: 'success', result: response.accessToken }, window.location.origin);
        }
      } else {
        // Otherwise, invoke login
        msalClient.loginRedirect({
          scopes: ['mail.send']
        });
      }
    })
    .catch(error => {
      // Check if the origin of the request is same as the registered redirect URI
      if (new URL(msalUrl).origin === window.location.origin) {
        window.opener.postMessage({
          status: 'failure', result: {
            errorCode: error.errorCode,
            errorMessage: error.errorMessage,
            stack: error.stack,
          }
        }, window.location.origin);
    }
    });
}); 

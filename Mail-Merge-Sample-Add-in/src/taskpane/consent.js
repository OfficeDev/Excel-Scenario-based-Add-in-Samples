(() => {
    // The initialize function must be run each time a new page is loaded
    Office.initialize = () => {
      const msalUrl = location.href.substring(0, location.href.lastIndexOf('/')) + '/consent.html';
      const msalClient = new msal.PublicClientApplication({
          auth: {
            clientId: '8d733658-f5be-4ba9-bb31-0da453645c49', // YOUR_APP_ID_HERE
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
                  Office.context.ui.messageParent( JSON.stringify({ status: 'success', result : response.accessToken }) );
              } else {
                  // Otherwise, invoke login
                  msalClient.loginRedirect({
                      scopes: ['mail.send']
                  });
              }
          })
          .catch(error => {
              Office.context.ui.messageParent( JSON.stringify({ status: 'failure', result: {
                errorCode: error.errorCode,
                errorMessage: error.errorMessage,
                stack: error.stack,
              } }));
          });
    };
  })();
  
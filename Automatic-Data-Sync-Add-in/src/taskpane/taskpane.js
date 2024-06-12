/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */
var Buffer = require('buffer/').Buffer;
let fileName = "https://graph.microsoft.com/v1.0/me/drive/root:/Demo.xlsx";
let filePath = "https://graph.microsoft.com/v1.0/me/drive/items";
let transferTokenUrl = "http://localhost:3001/graphApi";
let configDataSyncUrl = "http://localhost:3001/configDataSync";

Office.onReady((info) => {
  // Check that we loaded into Excel
  if (info.host === Office.HostType.Excel) {
    let userClientId = 'YOUR_APP_ID_HERE'; //Register your app at https://aad.portal.azure.com/
    localStorage.setItem('client-id', userClientId);

    document.getElementById("createSampleData").onclick = createSampleData;

    // Add event listener for configDataSync button
    document.getElementById('configDataSync').addEventListener('click', function () {
      var inputElement = document.querySelector('input[type="number"]');
      var inputValue = inputElement.value;

      // Send the inputValue to server.js
      fetch(configDataSyncUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ value: inputValue }),
      })
        .then(response => response.json())
        .then(showStatus("Transfer the task to Server side", false))
        .catch((error) => showStatus(`Exception transfer the task to Server: ${JSON.stringify(error)}`, true));
    });
  }
});

class DialogAPIAuthProvider {
  async getAccessToken() {
    if (this._accessToken) {
      return this._accessToken;
    } else {
      return this.login();
    }
  }

  async login() {
    return new Promise((resolve, reject) => {
      let data = encodeURIComponent(localStorage.getItem('client-id'));
      const dialogLoginUrl = location.href.substring(0, location.href.lastIndexOf('/')) + `/consent.html?data=${data}`;
      Office.context.ui.displayDialogAsync(
        dialogLoginUrl,
        { height: 60, width: 60 },
        result => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(result.error);
          }
          else {
            const loginDialog = result.value;

            loginDialog.addEventHandler(Office.EventType.DialogEventReceived, args => {
              reject(args.error);
            });

            loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, args => {
              const messageFromDialog = JSON.parse(args.message);

              loginDialog.close();

              if (messageFromDialog.status === 'success') {
                // We now have a valid access token.
                this._accessToken = messageFromDialog.result;

                // Send the access token to the server
                $.ajax({
                  url: transferTokenUrl,
                  type: 'POST',
                  data: JSON.stringify({ accessToken: this._accessToken }),
                  contentType: 'application/json',
                  success: function (data) {
                    if (data.status === 'success') {
                      console.log('Server successfully received the access token.');
                    } else {
                      console.error('Failed to send the access token to the server.');
                    }
                  },
                  error: function () {
                    console.error('Failed to send the request to the server.');
                  }
                });

                resolve(this._accessToken);

              }
              else {
                // Something went wrong with authentication or the authorization of the web application.
                reject(messageFromDialog.result);
              }
            });
          }
        }
      );
    });
  }
}

const dialogAPIAuthProvider = new DialogAPIAuthProvider();

// Display a status
/**
 * @param {unknown} message
 * @param {boolean} isError
 */
function showStatus(message, isError) {
  $('.status').empty();
  $('<div/>', {
    class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`
  }).append($('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: isError ? 'An error occurred' : 'Success'
  })).append($('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: message
  })).appendTo('.status');
}

// Create Sample Data
async function createSampleData() {
  const client = MicrosoftGraph.Client.initWithMiddleware({ authProvider: dialogAPIAuthProvider });
  await Excel.run(async (context) => {
    const emptyExcelFile = Buffer.alloc(0);
    client
      .api(fileName + ":/content")
      .put(emptyExcelFile)
      .then((res) => {
        console.log('New Excel file created', res);

        return client.api(fileName).get();
      })
      .then((file) => {
        console.log('File:', file);

        const workbookTable = {
          address: 'Sheet1!A1:B1',
          hasHeaders: true,
          name: "Table1"
        };

        return client
          .api(filePath + `/${file.id}/workbook/tables/add`)
          .post(workbookTable);
      })
      .then((updateResult) => {
        showStatus("Successfully create Demo.xlsx and add Table into it.", false);
      })
      .catch((error) => {
        console.error('Error creating new Excel file', error);
        showStatus("Error creating new Excel file", true);
      });
  });
}
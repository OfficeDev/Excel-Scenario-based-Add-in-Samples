const WebSocket = require('ws');
const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const cors = require('cors');
require('isomorphic-fetch');
const express = require('express');

const app = express();
app.use(cors()); // Enable CORS for all requests
app.use(express.json()); // for parsing application/json

// Start the server
app.listen(3001, () => console.log('Server listening on port 3001'));

app.post('/configDataSync', async (req, res) => {
    // Check if the request contains a value
    if (req.body && req.body.value) {
        // Print the value
        console.log('Received value:', req.body.value);

        // Convert value to milliseconds
        const interval = req.body.value * 1000;

        // Create an auth provider that returns the stored access token
        const dialogAPIAuthProvider = {
            getAccessToken: () => Promise.resolve(accessToken),
        };

        // Initialize the Graph client with the auth provider
        const client = MicrosoftGraph.Client.initWithMiddleware({ authProvider: dialogAPIAuthProvider });

        // Use the Graph client to get the file
        const file = await client.api('https://graph.microsoft.com/v1.0/me/drive/root:/Demo.xlsx').get();

        // Set an interval to add a row to the table every req.body.value seconds
        setInterval(async () => {
            // Define the row to add
            const row = {
                values: [[new Date().toISOString(), 'New row']]
            };
            console.log('Adding row:', row);

            // Use the Graph client to add a row to the table
            await client
                .api(`https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/workbook/tables/Table1/rows`)
                .post(row);
        }, interval);
    }

    // Send a response back to the client
    res.json({ message: 'Received value' });
});

let accessToken;
app.post('/graphApi', async (req, res) => {
    // Check if the request contains an access token
    if (req.body.accessToken) {
        accessToken = req.body.accessToken;
        res.json({ status: 'success' });
    } else {
        console.error('No access token in the request.');
        res.json({ status: 'error' });
    }
});
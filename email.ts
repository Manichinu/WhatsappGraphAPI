const express = require("express");
const session = require("express-session");
const cors = require("cors");
const app = express();
const {
    PublicClientApplication,
    ConfidentialClientApplication,
} = require("@azure/msal-node");
const msal = require('@azure/msal-node');
import axios from "axios";
import { Client } from '@microsoft/microsoft-graph-client';


app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));


console.log("test")
let port = 80;




// async function checkSubscriptionStatus(subscriptionId:any) {
//   try {
//     // Set the request URL to fetch the subscription status
//     const url = `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`;

//     // Send a GET request to the Microsoft Graph API to get subscription details
//     const response = await axios.get(url, {
//       headers: {
//         'Authorization': `Bearer ${accessToken}`,
//         'Content-Type': 'application/json'
//       }
//     });

//     // Log the subscription details returned from the Graph API
//     console.log('Subscription Status:', response.data);

//     // Optionally, you can access specific data from the response, like:
//     // - response.data.id (Subscription ID)
//     // - response.data.status (Subscription status)
//     // - response.data.expirationDateTime (Expiration date of the subscription)
//     // - response.data.notificationUrl (Notification URL for the subscription)

//     return response.data;
//   } catch (error:any) {
//     // Handle error
//     console.error('Error fetching subscription status:', error.response ? error.response.data : error.message);
//     throw error;
//   }
// }

// // Call the function to check the subscription status
// checkSubscriptionStatus(subscriptionId);

const msalConfig = {
    auth: {
        clientId: 'a866d7cd-4504-45e0-87d0-3d1039df49bf',
        // authority: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        authority: 'https://login.microsoftonline.com/common/',
        clientSecret: 'wB08Q~S71hpzEYxJoPUt4dEmFvQR6VhvIITdMbuI',
    },
};
const REDIRECT_URI = 'http://localhost:80/auth/callback';

const msalClient = new msal.ConfidentialClientApplication(msalConfig);
const GRAPH_API_URL = 'https://graph.microsoft.com/v1.0';
const WEBHOOK_URL = 'https://webhook.remodigital.in/notifications';


async function createSubscription(accessToken: any) {
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        },
    });
    try {
        // Verify the mailbox is accessible
        // const mailbox = await client.api('/me').get();
        // console.log('Mailbox details:', mailbox);
        const subscription = await client.api('/subscriptions').post({
            changeType: 'created', // Events to track
            notificationUrl: WEBHOOK_URL, // Your endpoint to receive notifications
            resource: '/me/messages', // Resource to track
            expirationDateTime: '2025-03-01T23:59:59.0000000Z', // Expiry time (max 1 hour for messages)
            // clientState: 'secretClientValue', // Optional: Ensures the notification is from Microsoft
        });

        console.log("Subscription created successfully:", subscription)
    } catch (error: any) {
        console.log("Error:", error)
    }
}
async function refreshToken(token: any) {
    const newTokenResponse = await msalClient.acquireTokenByRefreshToken({
        refreshToken: token,
        scopes: ['Mail.Read', 'Mail.ReadWrite', 'MailboxSettings.ReadWrite'],
    });

    console.log('New Access Token:', newTokenResponse.accessToken);
}

// Step 1: Redirect to the authorization URL
app.get('/auth', async (req: any, res: any) => {
    try {
        const authUrl = await msalClient.getAuthCodeUrl({
            scopes: ['api://a866d7cd-4504-45e0-87d0-3d1039df49bf/.default'],
            // scopes: ['User.Read', 'Mail.ReadWrite', 'MailboxSettings.ReadWrite'],
            redirectUri: 'http://localhost:80/auth/callback',
            prompt: 'consent'
        });

        res.redirect(authUrl); // Redirect to the resolved URL
    } catch (error) {
        console.error('Error generating auth URL:', error);
        res.status(500).send('Failed to generate authorization URL');
    }
});


// Step 2: Handle the callback and exchange for tokens
app.get('/auth/callback', async (req: any, res: any) => {
    try {
        const tokenResponse = await msalClient.acquireTokenByCode({
            code: req.query.code,
            scopes: ['api://a866d7cd-4504-45e0-87d0-3d1039df49bf/.default'],
            // scopes: ["openid", "profile", "offline_access", 'api://a866d7cd-4504-45e0-87d0-3d1039df49bf/.default'],
            // scopes: ['User.Read', 'Mail.ReadWrite', 'MailboxSettings.ReadWrite'],
            redirectUri: REDIRECT_URI,
        });

        console.log('Access Token:', tokenResponse.accessToken);
        // console.log('Refresh Token:', tokenResponse);
        createSubscription(tokenResponse.accessToken)
        res.send('Authentication successful!');
    } catch (error) {
        console.error('Error during token exchange:', error);
        res.status(500).send('Authentication failed');
    }
});




app.listen(port, () => {
    console.log(`app listening on port ${port}`);
});
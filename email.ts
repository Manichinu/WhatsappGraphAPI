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

// // Replace with your subscription ID
// const subscriptionId = '9d977e16-8d97-42e6-b399-0e3747179416';

// // Replace with your access token
// const accessToken = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6IjlUaXAwSWFhSlp1UGlQZ1R3WXkwdlhmeHVLRGFUSWg2S2R0ZEhYdmRybm8iLCJhbGciOiJSUzI1NiIsIng1dCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyIsImtpZCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81MzQ2M2E0Zi0wYzFkLTQ2YmMtODkwZS1kNmVlZWY3ODU3MzAvIiwiaWF0IjoxNzM1ODg3NDEyLCJuYmYiOjE3MzU4ODc0MTIsImV4cCI6MTczNTg5MTU3MywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhZQUFBQWNBTlI5ek5taEozT0ZPUUhFTFZWRStUcjBVNW1mRk4rUUIycUg3WGVvc3ZlU0k4c1ZQUkt3RWtOODVPMVB5OXVxS1QxeTlkTTZWcWJlQ2QwdWpadFAvMlpyU2NOd3ZKRmhlWjVJY2grcThFPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiV2hhdHNhcHAgYXBpIGNoZWNrIiwiYXBwaWQiOiIwMjFmNTNmOS01ZjU4LTRkNWUtYjYxNS00ZjVkODk5NDE0MTMiLCJhcHBpZGFjciI6IjEiLCJmYW1pbHlfbmFtZSI6IjIiLCJnaXZlbl9uYW1lIjoibWFuaSIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjI0MDk6NDBmNDozMDU5OmIzOTg6ZDg4Mjo3Yjg0OmM1NTc6YTVhNyIsIm5hbWUiOiJtYW5pIDIiLCJvaWQiOiJmYmE1ZTk1Ny05Y2ZkLTQ4OTQtODg2Yi02ZDY3YzkxMmViZjIiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDFFRjAyRkJBQyIsInJoIjoiMS5BWFlBVHpwR1V4ME12RWFKRHRidTczaFhNQU1BQUFBQUFBQUF3QUFBQUFBQUFBQzBBSWQyQUEuIiwic2NwIjoiTWFpbC5SZWFkV3JpdGUgTWFpbGJveFNldHRpbmdzLlJlYWRXcml0ZSBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgZW1haWwiLCJzaWduaW5fc3RhdGUiOlsia21zaSJdLCJzdWIiOiJFV3VFYWhiaHpIbmdHQ1RUUFNuSW5VSjB1c2lBVEM5Nkd6NmJmbjRYWlhrIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6Ik5BIiwidGlkIjoiNTM0NjNhNGYtMGMxZC00NmJjLTg5MGUtZDZlZWVmNzg1NzMwIiwidW5pcXVlX25hbWUiOiJtYW5pMkA2ejBsN3Yub25taWNyb3NvZnQuY29tIiwidXBuIjoibWFuaTJANnowbDd2Lm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6Ii1wY04wV3dzUTBTQXRzNEhjNS1GQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfZnRkIjoiUE9YU0xZSVBpVndKWmpzSDBJNDJNTGRtTXZtNFhWcGJLRURveEd1bHg0cyIsInhtc19pZHJlbCI6IjEgMTQiLCJ4bXNfc3QiOnsic3ViIjoiZ28wa3JDUHM1azgyYzJaNmhRNzJPcE94N3pQcmphVS1NV181Z1p6cG5ZVSJ9LCJ4bXNfdGNkdCI6MTY0Njg3OTI5MX0.oqAzQowomgSkIjM_whUJ2BPVxtf1lp7uRS03TrIQ9D1ISDbo31KUnTF7hRnHnZvTbmIUEIAKWI0v7yZ0iXYqo-ts0iH7TXHEXWTBTZebtYX0POtE2teuwKyGBbUQ-Pskkkwn9ja19GblfSej8BvmhPmY-V51kmbkFuvgx-cWvQYam13U-YHrUS4AnXEB4Fa5FZqx_PI2eae4X2nLbJ0zOp60t4DEnIXcnuY7BcxBaJFh81cWuC0OwC-xBAasWV1lkBXkdIq7XNpz1-fVBg-vxI0-qJrLFvGX0VYKY6Fp5lesNgLMCwgYIzXphmeEmO3poOEJ_Nv2-qbw-uzs0lvHbA';

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
        const subscription = await client.api('/subscriptions').post({
            changeType: 'created', // Events to track
            notificationUrl: WEBHOOK_URL, // Your endpoint to receive notifications
            resource: '/me/messages', // Resource to track
            expirationDateTime: '2025-01-10T23:59:59.0000000Z', // Expiry time (max 1 hour for messages)
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
            scopes: ['User.Read'],
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
            scopes: ['User.Read','Mail.Read'],
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
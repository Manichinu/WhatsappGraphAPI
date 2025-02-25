import express from "express";
import axios from "axios";
import cors from "cors";
import dotenv from "dotenv";
import { Document, Packer, Table, TableRow, TableCell, Paragraph, TextRun, BorderStyle, WidthType, AlignmentType } from "docx";
import fs from "fs";
import bodyParser from "body-parser";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import mammoth from 'mammoth';
const ImageModule = require('open-docxtemplater-image-module');
import 'isomorphic-fetch';
import { Client } from '@microsoft/microsoft-graph-client';

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

interface EnvVariables {
  WEBHOOK_VERIFY_TOKEN: string;
  GRAPH_API_TOKEN: string;
  PORT: string;
  SITE_URL: string;
  CLIENT_ID: string;
  TENANT_ID: string;
  LIST_NAME: string;
  UserNameVal: string;
  PasswordVal: any;
  CLIENT_SECRET: string;
  LIBRARY_NAME: string;
}

const {
  WEBHOOK_VERIFY_TOKEN,
  GRAPH_API_TOKEN,
  PORT,
  SITE_URL,
  CLIENT_ID,
  TENANT_ID,
  LIST_NAME,
  UserNameVal,
  PasswordVal,
  CLIENT_SECRET,
  LIBRARY_NAME
} = process.env as unknown as EnvVariables;


// Step 1: Set your required parameters
const GRAPH_API_URL = 'https://graph.microsoft.com/v1.0';
const WEBHOOK_URL = 'https://webhook.remodigital.in/notifications';  // Your webhook URL
let ACCESS_TOKEN_For_Email = '';  // Your OAuth access token

// Step 2: Define the subscription body
const subscriptionBody = {
  changeType: 'created',  // Trigger when a new email is created
  notificationUrl: WEBHOOK_URL,  // Your webhook endpoint
  resource: '/users/siva@782yjz.onmicrosoft.com/messages',  // Subscription to new emails (use /users/{id}/messages for specific users)
  expirationDateTime: '2025-03-01T23:59:59.0000000Z',  // Set an expiration time for the subscription
  // clientState: 'your-client-state',  // Optional, custom state to validate the webhook response
};
async function createSubscription() {
  try {
    ACCESS_TOKEN_For_Email = await getAccessToken();
    console.log(ACCESS_TOKEN_For_Email)
    // fetchEmails(ACCESS_TOKEN_For_Email)
    // const userDetails = await axios.get(`${GRAPH_API_URL}/users/ea40397d-6d24-4de0-9acd-07f60abe667d`, {
    //   headers: {
    //     Authorization: `Bearer ${ACCESS_TOKEN_For_Email}`,
    //   },
    // });
    // console.log(userDetails.data);
    const response = await axios.post(
      `${GRAPH_API_URL}/subscriptions`,
      subscriptionBody,
      {
        headers: {
          Authorization: `Bearer ${ACCESS_TOKEN_For_Email}`,
          'Content-Type': 'application/json',
        },
      }
    );

    console.log('Subscription created successfully:', response.data);
    // const responses = await axios.get(`${GRAPH_API_URL}/subscriptions`, {
    //   headers: {
    //     Authorization: `Bearer ${ACCESS_TOKEN_For_Email}`,
    //   },
    // });

    // console.log('Subscriptions:', responses.data);
    // // Iterate through each subscription to get the user's email
    // for (const subscription of responses.data.value) {
    //   const userId = subscription.resource.split('/')[1];  // Extract userId from the resource path

    //   // Fetch the user details using userId
    //   const userResponse = await axios.get(`${GRAPH_API_URL}/users/${userId}`, {
    //     headers: {
    //       Authorization: `Bearer ${ACCESS_TOKEN_For_Email}`,
    //     },
    //   });
    //   console.log("Response", userResponse.data)
    //   // Extract and log the email
    //   const userEmail = userResponse.data.mail || userResponse.data.userPrincipalName;
    //   console.log(`User ID: ${userId}, Email: ${userEmail}`);
    // }
  } catch (error: any) {
    console.error('Error creating subscription:', error.response ? error.response.data : error.message);
  }
}
createSubscription();
async function fetchEmails(accessToken: any) {
  try {
    const response = await axios.get(`${GRAPH_API_URL}/users/e03ef8d5-a78b-4f3f-946d-1191dafbd3c0/messages`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    console.log("Emails:", response.data);
  } catch (error: any) {
    console.error("Error fetching emails:", error.response?.data || error.message);
  }
}

async function getAccessToken() {
  const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append("client_id", CLIENT_ID);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("grant_type", "client_credentials");
  params.append("client_secret", CLIENT_SECRET);
  // params.append("client_id", CLIENT_ID);
  // params.append("scope", "user.read openid profile offline_access");
  // params.append("username", UserNameVal);
  // params.append("password", PasswordVal);
  // params.append("grant_type", "password");
  // params.append("client_secret", CLIENT_SECRET)

  try {
    const response = await axios.post(tokenEndpoint, params, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });
    const { access_token } = response.data;

    if (!access_token) {
      throw new Error("Failed to obtain access token");
    }
    return access_token;
  } catch (error: any) {
    if (axios.isAxiosError(error)) {
      console.error("Error acquiring access token:", error.response?.data || error.message);
    } else {
      console.error("Error acquiring access token:", error.message);
    }
    throw error;
  }
}
async function getSiteId(accessToken: string) {
  const siteEndpoint = `https://graph.microsoft.com/v1.0/sites/${SITE_URL}`;
  try {
    const response = await axios.get(siteEndpoint, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    return response.data.id;
  } catch (error: any) {
    console.error("Error acquiring site ID:", error.response?.data || error.message);
    throw error;
  }
}
async function getListId(accessToken: string, siteId: string) {
  const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${LIST_NAME}`;
  try {
    const response = await axios.get(listEndpoint, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    return response.data.id;
  } catch (error: any) {
    console.error("Error acquiring list ID:", error.response?.data || error.message);
    throw error;
  }
}
async function getLibraryId(accessToken: string, siteId: string) {
  const listEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${LIBRARY_NAME}`;
  try {
    const response = await axios.get(listEndpoint, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    return response.data.id;
  } catch (error: any) {
    console.error("Error acquiring list ID:", error.response?.data || error.message);
    throw error;
  }
}
async function getAllListItems(accessToken: string, siteId: string, listId: string) {
  let listItemsEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  let items: any[] = [];
  const fields = ['Title', 'MaximumAllowedNotificationCount', 'ConsumedNotificationCount', 'CustomerWhatsAppNumber', 'IsServiceLive', 'ValidUpto', 'ID'];

  // Construct the expand and select query parameters for fields
  const expandQuery = fields.length > 0 ? `?$expand=fields($select=${fields.join(',')})` : '';

  try {
    while (listItemsEndpoint) {
      const response = await axios.get(listItemsEndpoint + expandQuery, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      items = items.concat(response.data.value);

      // Check if there is a next link for pagination
      listItemsEndpoint = response.data['@odata.nextLink'] || null;
    }
    return items;
  } catch (error: any) {
    console.error("Error acquiring list items:", error.response?.data || error.message);
    throw error;
  }
}
async function updateListItem(accessToken: any, siteId: any, listId: any, itemId: any, fieldsToUpdate: any) {
  const updateItemEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}`;

  try {
    const response = await axios.patch(updateItemEndpoint, {
      fields: fieldsToUpdate
    }, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    // console.log('Updated Item:', response.data);
    return response.data;
  } catch (error: any) {
    console.error("Error updating list item:", error.response?.data || error.message);
    throw error;
  }
}
async function getDriveId(accessToken: any, siteId: any) {
  const driveEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;

  try {
    const response = await axios.get(driveEndpoint, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    const drives = response.data.value;
    if (!drives || drives.length === 0) {
      throw new Error("No drives found in the specified site.");
    }

    return drives[0].id; // Return the ID of the first drive; adjust if necessary
  } catch (error: any) {
    console.error("Error fetching drive ID:", error.response?.data || error.message);
    throw error;
  }
}
app.post("/test", async (req, res) => {
  var body = req.body

  // Process the actual change notification (new email)
  const changeNotification = body.value[0]; // Assuming it's an array, so we access the first element
  const resourceData = changeNotification.resourceData;

  // Manually extracting the userId and messageId from the payload
  const resourcePath = changeNotification.resource; // "Users/48346217-e4c8-4c4f-b783-3e3fce527d3f/Messages/AAMkADg4OTQxZGMwLTM1M2YtNDg5Ny05YjczLThhNmQ0MmYxNTQ5MwBGAAAAAAC3t5WqMQWDSpbtKLt6yo4bBwC7F3pBwatPTq99X0VNUbmHAAAAAAEMAAC7F3pBwatPTq99X0VNUbmHAALaw9ywAAA="
  const userId = resourcePath.split('/')[1]; // "48346217-e4c8-4c4f-b783-3e3fce527d3f"
  const messageId = resourceData.id; // "AAMkADg4OTQxZGMwLTM1M2YtNDg5Ny05YjczLThhNmQ0MmYxNTQ5MwBGAAAAAAC3t5WqMQWDSpbtKLt6yo4bBwC7F3pBwatPTq99X0VNUbmHAAAAAAEMAAC7F3pBwatPTq99X0VNUbmHAALaw9ywAAA="

  console.log('Extracted userId:', userId);
  console.log('Extracted messageId:', messageId);

  // Fetch the full email content using Microsoft Graph API
  try {
    const accessToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkNqeFZHakVWM3I4c0pXbWN2bnJ1bFlVWVBSeDlKQlFoS3pYRXNVc3JoY0EiLCJhbGciOiJSUzI1NiIsIng1dCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyIsImtpZCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81MzQ2M2E0Zi0wYzFkLTQ2YmMtODkwZS1kNmVlZWY3ODU3MzAvIiwiaWF0IjoxNzM2MTM5MDQ5LCJuYmYiOjE3MzYxMzkwNDksImV4cCI6MTczNjE0NDM5OCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhZQUFBQUdKcFJlRTc5UW1FNTBDTUszQVhSb0UvMzBYVGE5RmQ0d1pINUZiRlA1YXdzQ0p2Z202bTdFc1RZZC9GV1B0bVJZZGNnckp5a1Yvc1dkZVIyYTF6MnloM0xUTTEvcytkaXFUL0oxaGI1YVJJPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiV2hhdHNhcHAgYXBpIGNoZWNrIiwiYXBwaWQiOiIwMjFmNTNmOS01ZjU4LTRkNWUtYjYxNS00ZjVkODk5NDE0MTMiLCJhcHBpZGFjciI6IjEiLCJmYW1pbHlfbmFtZSI6IjIiLCJnaXZlbl9uYW1lIjoibWFuaSIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjI0MDk6NDBmNDozMDEyOjdkMzU6MjE5MTo4YTMzOjc4NGQ6MTliZCIsIm5hbWUiOiJtYW5pIDIiLCJvaWQiOiJmYmE1ZTk1Ny05Y2ZkLTQ4OTQtODg2Yi02ZDY3YzkxMmViZjIiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDFFRjAyRkJBQyIsInJoIjoiMS5BWFlBVHpwR1V4ME12RWFKRHRidTczaFhNQU1BQUFBQUFBQUF3QUFBQUFBQUFBQzBBSWQyQUEuIiwic2NwIjoiTWFpbC5SZWFkV3JpdGUgTWFpbGJveFNldHRpbmdzLlJlYWRXcml0ZSBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgZW1haWwiLCJzaWduaW5fc3RhdGUiOlsia21zaSJdLCJzdWIiOiJFV3VFYWhiaHpIbmdHQ1RUUFNuSW5VSjB1c2lBVEM5Nkd6NmJmbjRYWlhrIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6Ik5BIiwidGlkIjoiNTM0NjNhNGYtMGMxZC00NmJjLTg5MGUtZDZlZWVmNzg1NzMwIiwidW5pcXVlX25hbWUiOiJtYW5pMkA2ejBsN3Yub25taWNyb3NvZnQuY29tIiwidXBuIjoibWFuaTJANnowbDd2Lm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImN0MkNIbkQ4OTBLbHpTQVItbWNWQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfZnRkIjoiYUpUeGdsTDV5cUhDY2E5REhiazlTWFNFS1ZjVjVINXlMZXF2dU5ySEY0QSIsInhtc19pZHJlbCI6IjEgMTIiLCJ4bXNfc3QiOnsic3ViIjoiZ28wa3JDUHM1azgyYzJaNmhRNzJPcE94N3pQcmphVS1NV181Z1p6cG5ZVSJ9LCJ4bXNfdGNkdCI6MTY0Njg3OTI5MX0.Bzt76wAyB3ZPQUwwLrLLPZ2XYPGDasRmxPpgrojVr9mA88IayShc-B_HxIrEGxuml5kg5XYztudS7PYXTWo1VcZRB8UMeahHjjhqqPBqHY81iGXTGjUTOltkw55RYydRghXTObCuL_Wf6KP9Q5Z78AeRjnHnfoA1ZoGWQU1X4auYT5F4nU2sxL6t4tjOH03LoxrbJAHC9DXLakTLOi4No0_hLwQhmQ6ClLUUon51Z-kmOiux2cULOHyZNVrLfATsN0y36X-nyesxdgg4oLubU5sRS0JiE868N92djBBL9yoFkOpyR7g38B_10OpMig1pF9sGccZUjkuZ5QE7S4CbqA"; // Ensure this function returns a valid token

    // Construct the API URL to fetch the email message
    const emailResponse = await axios.get(`https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      }
    });

    // Extract relevant email information from the response
    const emailContent = {
      messageId: emailResponse.data.id,
      subject: emailResponse.data.subject,
      sender: emailResponse.data.sender,
      receivedDateTime: emailResponse.data.receivedDateTime,
      bodyPreview: emailResponse.data.bodyPreview,
      body: emailResponse.data.body // Depending on how you want the body (text or HTML)
    };

    console.log('Full email content:', emailContent);
    addCategoryToMessage(accessToken, userId, messageId, "Original");

  } catch (error) {
    console.error('Error retrieving email content or sending to Logic App:', error);
  }

  // Respond with status 200 OK to acknowledge receipt of the notification
})

// Function to add category to an email message
async function addCategoryToMessage(accessToken: any, userId: any, messageId: any, category: any) {
  const url = `https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}`;

  try {
    const response = await axios.patch(url, {
      categories: [category], // Assign the category to the email
    }, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    return response.data;
  } catch (error: any) {
    console.error('Error adding category to message:', error.response?.data || error.message);
    throw error;
  }
}

let accessToken: any;
let siteId;
let listId;
let ListItems;
let libraryId;

// async function createWordDocument() {
//   const table = new Table({
//     rows: [
//       new TableRow({
//         children: [
//           new TableCell({
//             children: [new Paragraph("S.No")],
//           }),
//           new TableCell({
//             children: [new Paragraph("Name")],
//           }),
//           new TableCell({
//             children: [new Paragraph("Age")],
//           }),
//           new TableCell({
//             children: [new Paragraph("District")],
//           }),
//         ],
//       }),
//       ...data.map(item =>
//         new TableRow({
//           children: [
//             new TableCell({
//               children: [new Paragraph(item.sNo.toString())],
//             }),
//             new TableCell({
//               children: [new Paragraph(item.name)],
//             }),
//             new TableCell({
//               children: [new Paragraph(item.age.toString())],
//             }),
//             new TableCell({
//               children: [new Paragraph(item.district)],
//             }),
//           ],
//         })
//       ),
//     ],
//   });

//   const doc = new Document({
//     sections: [
//       {
//         children: [table],
//       },
//     ],
//   });

//   const buffer = await Packer.toBuffer(doc);
//   fs.writeFileSync("DataTable.docx", buffer);
// }

// createWordDocument().catch(console.error);

app.post("/webhook", async (req, res) => {
  console.log("Incoming webhook message:", JSON.stringify(req.body, null, 2));

  const message = req.body.entry?.[0]?.changes?.[0]?.value?.messages?.[0];
  const businessPhoneNumberId = req.body.entry?.[0]?.changes?.[0]?.value?.metadata?.phone_number_id;

  if (message?.type === "text") {
    await axios({
      method: "POST",
      url: `https://graph.facebook.com/v18.0/${businessPhoneNumberId}/messages`,
      headers: {
        Authorization: `Bearer ${GRAPH_API_TOKEN}`,
      },
      data: {
        messaging_product: "whatsapp",
        status: "read",
        message_id: message.id,
      },
    });
  }

  res.sendStatus(200);
});

app.get("/webhook", (req, res) => {
  const mode = req.query["hub.mode"];
  const token = req.query["hub.verify_token"];
  const challenge = req.query["hub.challenge"];

  if (mode === "subscribe" && token === WEBHOOK_VERIFY_TOKEN) {
    res.status(200).send(challenge);
    console.log("Webhook verified successfully!");
  } else {
    res.sendStatus(403);
  }
});

app.post("/whatsapp", async (req, res) => {
  // console.log("Details: ", req.body);
  // console.log("Response: ", res);

  const { PhoneNumberID, from, Accesstoken, to } = req.body


  accessToken = await getAccessToken();
  siteId = await getSiteId(accessToken);
  listId = await getListId(accessToken, siteId);
  ListItems = await getAllListItems(accessToken, siteId, listId)
  var MatchedItem = ListItems.filter((item) => {
    return item.fields.CustomerWhatsAppNumber == from;
  });
  let TotalCounts = MatchedItem[0].fields.MaximumAllowedNotificationCount;
  let ConsumedCounts = MatchedItem[0].fields.ConsumedNotificationCount;
  let ID = MatchedItem[0].fields.id;
  let Status = MatchedItem[0].fields.IsServiceLive
  // console.log(TotalCounts, ConsumedCounts, ID)
  if ((ConsumedCounts < TotalCounts) && Status == true) {
    var settings = {
      "url": `https://graph.facebook.com/v19.0/${PhoneNumberID}/messages`,
      "method": "POST",
      "timeout": 0,
      "headers": {
        "Authorization": `Bearer ${Accesstoken}`,
        "Content-Type": "application/json"
      },
      "data": req.body,
    };
    try {
      const response = await axios(settings);
      console.log("Response:", response.data); // Log the response data     
    } catch (error: any) {
      console.error(error.response ? error.response.data : error.message);
    }
    const fieldsToUpdate = {
      ConsumedNotificationCount: ConsumedCounts + 1
    };
    updateListItem(accessToken, siteId, listId, ID, fieldsToUpdate)
      .then(updatedItem => {
        // console.log('Updated Item:', updatedItem);
      })
      .catch(error => {
        console.error('Error:', error);
      });
    // res.send("Message sent");

  } else {
    const fieldsToUpdate = {
      IsServiceLive: false
    };
    updateListItem(accessToken, siteId, listId, ID, fieldsToUpdate)
      .then(updatedItem => {
        // console.log('Updated Item:', updatedItem);
      })
      .catch(error => {
        console.error('Error:', error);
      });
    res.send("You’ve reached the maximum limit for WhatsApp notification. Please contact your service provider for further assistance.");
    console.log("Total Count exceeded")
  }
});
app.post("/outlook", async (req, res) => {
  const fetch = require("node-fetch");
  accessToken = await getAccessToken();
  console.log(accessToken)
  var messageId = "CAHHnKiijqGP4n9HTDvm6wREwRsNO1n+5BytfX0eZCSWctyEGsQ@mail.gmail.com"
  var category = "Original"
  const url = `https://graph.microsoft.com/v1.0/me/messages/${messageId}`;
  const response = await fetch(url, {
    method: "PATCH",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      "categories": [category] // Example: "Phishing", "Legitimate"
    })
  });
  if (response.ok) {
    console.log("Email categorized successfully");
  } else {
    console.error("Error categorizing email", await response.text());
  }

})
app.post('/notifications', (req, res) => {
  console.log('Received notification:', req.body);

  // Check if it's a validation request (required for subscription creation)
  if (req.query.validationToken) {
    return res.status(200).send(req.query.validationToken);  // Respond to validation request
  }

  // Process the actual change notification (new email)
  const changeNotification = req.body.value;
  console.log('Change notification received:', changeNotification);

  // Respond with status 200 OK to acknowledge receipt
  res.sendStatus(200);
});
app.post("/generate-documents", async (req, res) => {
  const data = req.body;
  // Define the paths
  const templatePath = path.join(__dirname, 'Assets', 'Templates.docx');
  const outputPath = path.join(__dirname, 'Assets', 'output.docx');

  // Load the docx file as binary content
  const content = fs.readFileSync(templatePath, 'binary');

  // Create a new PizZip instance to read the binary content
  const zip = new PizZip(content);



  // Create a new Docxtemplater instance with the image module
  // const doc = new Docxtemplater(zip, {
  //   paragraphLoop: true,
  //   linebreaks: true,
  // });


  // Replace placeholders with actual values
  // doc.render({
  //   User: data.title,
  //   price: data.price,
  //   details: data.details
  // });
  const ImageModule = require("docxtemplater-image-module");

  const imageOptions = {
    getImage(tagValue: fs.PathOrFileDescriptor, tagName: any, meta: any) {
      console.log({ tagValue, tagName, meta });
      return fs.readFileSync(tagValue);
    },
    getSize(img: any) {
      // it also is possible to return a size in centimeters, like this : return [ "2cm", "3cm" ];
      return [150, 150];
    },
  };

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [new ImageModule(imageOptions)],
  });
  doc.render({ image: "./Assets/bmw.jpg" });


  // doc.render(data)
  // Generate the modified document
  const buf = doc.getZip().generate({ type: 'nodebuffer' });

  // Save the modified document to a new file
  fs.writeFileSync(outputPath, buf);
  // Set headers for file download
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", "attachment; filename=GeneratedTemplate.docx");
  res.send(buf);

  console.log('Document created successfully!');


  // Step 1: Fetch the template file from SharePoint  
  // async function getTemplateFile(accessToken: any, siteId: any, libraryId: any, fileName: string) {
  //   const listItemsEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${libraryId}/items`;

  //   try {
  //     // Step 1: List items in the drive
  //     const listResponse = await axios.get(listItemsEndpoint, {
  //       headers: {
  //         Authorization: `Bearer ${accessToken}`,
  //       },
  //     });

  //     // console.log("List response:", listResponse.data);

  //     const items = listResponse.data.value;
  //     if (!items || items.length === 0) {
  //       throw new Error("No items found in the specified library.");
  //     }

  //     const fileId = items[0].id;
  //     const fileEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${libraryId}/items/${fileId}/driveItem/content`;

  //     // Step 3: Fetch the item content
  //     const fileResponse = await axios.get(fileEndpoint, {
  //       headers: {
  //         Authorization: `Bearer ${accessToken}`,
  //       },
  //       responseType: 'arraybuffer', // Important to get the file as binary data
  //     });

  //     // console.log("File response:", fileResponse.data);

  //     return Buffer.from(fileResponse.data);
  //   } catch (error: any) {
  //     console.error("Error fetching template file:", error.response?.data || error.message);
  //     throw error;
  //   }
  // }

  // // Step 2: Upload the generated document back to SharePoint
  // async function uploadFileToSharePoint(accessToken: any, siteId: any, fileName: any, fileContent: string) {
  //   const libraryId = await getDriveId(accessToken, siteId);

  //   const uploadEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${libraryId}/root:/${fileName}:/content`;

  //   try {
  //     const response = await axios.put(uploadEndpoint, fileContent, {
  //       headers: {
  //         Authorization: `Bearer ${accessToken}`,
  //         'Content-Type': 'application/octet-stream',
  //       },
  //     });

  //     console.log(`File '${fileName}' uploaded successfully to SharePoint.`);
  //     return response.data;
  //   } catch (error: any) {
  //     console.error("Error uploading file to SharePoint:", error.response?.data || error.message);
  //     throw error;
  //   }
  // }



  // try {
  //   accessToken = await getAccessToken();
  //   siteId = await getSiteId(accessToken);
  //   libraryId = await getLibraryId(accessToken, siteId);
  //   // console.log("accessToken", accessToken)
  //   // console.log("siteId", siteId)
  //   // console.log("libraryId", libraryId)

  //   // Get the template file from SharePoint
  //   const templateFileContent = await getTemplateFile(accessToken, siteId, libraryId, "Templates.docx");

  //   // Create a new PizZip instance to read the binary content
  //   const zip = new PizZip(templateFileContent);

  //   // Create a new Docxtemplater instance
  //   const doc = new Docxtemplater(zip, {
  //     paragraphLoop: true,
  //     linebreaks: true,
  //   });

  //   // Replace placeholders with actual values
  //   doc.render(data);

  //   // Generate the modified document
  //   const buf: any = doc.getZip().generate({ type: 'nodebuffer' });
  //   // Define the name for the generated document
  //   const generatedFileName = `GeneratedDocument_${Date.now()}.docx`;

  //   // Upload the generated document back to SharePoint
  //   const uploadedFile = await uploadFileToSharePoint(accessToken, siteId, generatedFileName, buf);

  //   res.send({ message: 'Document created and uploaded successfully!', file: uploadedFile });
  // } catch (error: any) {
  //   res.status(500).send({ error: error.message });
  // }
})
// Endpoint to handle authentication response
// app.post('/auth/callback', (req, res) => {
//   const { accessToken, userEmail } = req.body;

//   if (!accessToken || !userEmail) {
//     return res.status(400).json({ error: "Access token or user email is missing." });
//   }

//   console.log(`Access token received: ${accessToken}`);
//   console.log(`User email: ${userEmail}`);

//   // Optionally fetch user details or emails using Microsoft Graph
//   getUserEmails(accessToken)
//     .then((emails) => {
//       console.log('User emails:', emails);
//       res.status(200).json({ message: "Data fetched successfully", emails });
//     })
//     .catch((error) => {
//       console.error("Error fetching emails:", error);
//       res.status(500).json({ error: "Failed to fetch emails." });
//     });
// });

// // Function to fetch emails using Microsoft Graph
// async function getUserEmails(accessToken: any) {
//   const client = Client.init({
//     authProvider: (done) => {
//       done(null, accessToken);
//     },
//   });

//   const emails = await client.api('/me/messages').get();
//   return emails.value;
// }
app.get("/auth/callback", async (req: any, res: any) => {
  const { code } = req.query;
  console.log(req)

  if (!code) {
    return res.status(400).send("Authorization code is missing now.");
  }

  try {
    // Exchange authorization code for access token
    const tokenEndpoint = `https://login.microsoftonline.com/common/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append("client_id", CLIENT_ID);
    params.append("scope", "Mail.ReadWrite");
    params.append("code", code);
    params.append("redirect_uri", "https://webhook.remodigital.in/auth/callback");
    params.append("grant_type", "authorization_code");
    params.append("client_secret", CLIENT_SECRET);

    const tokenResponse = await axios.post(tokenEndpoint, params, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });

    const { access_token } = tokenResponse.data;

    if (!access_token) {
      return res.status(500).send("Failed to obtain access token.");
    }

    console.log(`Access token received: ${access_token}`);

    // Optional: Use the access token to get user details (e.g., email)
    const userResponse = await axios.get("https://graph.microsoft.com/v1.0/me", {
      headers: {
        Authorization: `Bearer ${access_token}`,
      },
    });

    const userEmail = userResponse.data.mail || userResponse.data.userPrincipalName;
    console.log(`User email: ${userEmail}`);

    // Create subscription
    const subscription = await createEmailSubscription(access_token);

    res.status(200).json({
      message: "Subscription created successfully",
      subscription,
    });
    // Send the email content to the Azure Logic App
    const postItem = {
      url: "https://prod-37.westus.logic.azure.com:443/workflows/7263399fc246463e9f7cadf1209f40b8/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aR2w_zRDafMqw5ucMCrK4jREfRkHDJKhG5aoBz_VssU",
      method: "POST",
      timeout: 0,
      headers: {
        'Access-Control-Allow-Origin': '*',
        "Accept": "application/json; odata=nometadata",
        "Content-Type": "application/json; odata=nometadata"
      },
      data: {
        Message: subscription,

      }
    };

    const response = await axios(postItem);
    console.log('Response from Azure Logic App:', response.data);
  } catch (error: any) {
    console.error("Error handling callback:", error.response?.data || error.message);
    res.status(500).send("Failed to handle callback.");
  }
});

// Function to create an email subscription
async function createEmailSubscription(accessToken: any) {
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  const subscription = await client.api('/subscriptions').post({
    changeType: 'created', // Events to track
    notificationUrl: 'https://webhook.remodigital.in/notifications', // Your endpoint to receive notifications
    resource: '/me/messages', // Resource to track
    expirationDateTime: '2024-12-30T23:59:59.0000000Z', // Expiry time (max 1 hour for messages)
    // clientState: 'secretClientValue', // Optional: Ensures the notification is from Microsoft
  });

  return subscription;
}
app.post("/usercreation", async (req: any, res: any) => {
  ACCESS_TOKEN_For_Email = await getAccessToken();
  const EXTERNAL_USER_EMAIL = 'eservices@tmax.in';
  try {
    const body = {
      accountEnabled: true, // Enables the account immediately
      displayName: 'Eservice', // Display name of the user
      mailNickname: 'Eservice', // Unique nickname for the user
      userPrincipalName: EXTERNAL_USER_EMAIL.replace('@', '_') + `#EXT#@ilgtech.onmicrosoft.com`, // Required for guest users
      mail: "eservices@tmax.in", // Email address
      userType: 'Guest', // Specifies this is a guest user
      passwordProfile: {
        password: 'Temp@12345', // Temporary password
        forceChangePasswordNextSignIn: false, // No password change required
      },
    };

    const response = await axios.post(`${GRAPH_API_URL}/users`, body, {
      headers: {
        Authorization: `Bearer ${ACCESS_TOKEN_For_Email}`,
        'Content-Type': 'application/json',
      },
    });

    console.log('External user created successfully:', response.data);


  } catch (error: any) {
    console.error('Error creating external user:', error.response?.data || error.message);
  }
})

app.get("*", (req, res) => {
  res.send("API is hosted for Graph API");
});

app.listen(PORT, () => {
  console.log(`Server is listening on port: ${PORT}`);
});

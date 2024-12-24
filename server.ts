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
  expirationDateTime: '2024-12-30T23:59:59.0000000Z',  // Set an expiration time for the subscription
  // clientState: 'your-client-state',  // Optional, custom state to validate the webhook response
};
async function createSubscription() {
  try {
    ACCESS_TOKEN_For_Email = await getAccessToken();
    console.log(ACCESS_TOKEN_For_Email)
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
  } catch (error: any) {
    console.error('Error creating subscription:', error.response ? error.response.data : error.message);
  }
}
createSubscription();

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
    const accessToken = await getAccessToken(); // Ensure this function returns a valid token

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

app.get("*", (req, res) => {
  res.send("API is hosted for Graph API");
});

app.listen(PORT, () => {
  console.log(`Server is listening on port: ${PORT}`);
});

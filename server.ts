import express from "express";
import axios from "axios";
import cors from "cors";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

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
} = process.env as unknown as EnvVariables;

async function getAccessToken() {
  const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  // params.append("client_id", CLIENT_ID);
  // params.append("scope", "https://graph.microsoft.com/.default");
  // params.append("grant_type", "client_credentials");
  // params.append("client_secret", CLIENT_SECRET);
  params.append("client_id", CLIENT_ID);
  params.append("scope", "user.read openid profile offline_access");
  params.append("username", UserNameVal);
  params.append("password", PasswordVal);
  params.append("grant_type", "password");
  params.append("client_secret", CLIENT_SECRET)

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
async function getAllListItems(accessToken: string, siteId: string, listId: string) {
  let listItemsEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  let items: any[] = [];
  const fields = ['Title', 'ConsumedCounts', 'PhoneNumber', 'TotalCounts', 'ID'];

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

let accessToken: any;
let siteId;
let listId;
let ListItems;

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

app.post("/data", async (req, res) => {
  // console.log("Details: ", req.body);
  const { PhoneNumberID, from, to, token, MessageTemplate } = req.body

  accessToken = await getAccessToken();
  siteId = await getSiteId(accessToken);
  listId = await getListId(accessToken, siteId);
  ListItems = await getAllListItems(accessToken, siteId, listId)
  var MatchedItem = ListItems.filter((item) => {
    return item.fields.PhoneNumber == from;
  });

  let TotalCounts = MatchedItem[0].fields.TotalCounts;
  let ConsumedCounts = MatchedItem[0].fields.ConsumedCounts;
  let ID = MatchedItem[0].fields.id;

  if (ConsumedCounts < TotalCounts) {
    var settings = {
      "url": `https://graph.facebook.com/v19.0/${PhoneNumberID}/messages`,
      "method": "POST",
      "timeout": 0,
      "headers": {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      "data": JSON.stringify({
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": `${to}`,
        "type": "text",
        "text": {
          "body": `${MessageTemplate}`
        }
      }),
    };
    const response = await axios(settings);
    const fieldsToUpdate = {
      ConsumedCounts: ConsumedCounts + 1
    };
    updateListItem(accessToken, siteId, listId, ID, fieldsToUpdate)
      .then(updatedItem => {
        // console.log('Updated Item:', updatedItem);
      })
      .catch(error => {
        console.error('Error:', error);
      });

  } else {
    console.log("Total Count exceeded")
  }
  res.send("Data received");
});

app.get("*", (req, res) => {
  res.send("API is hosted for Graph API");
});

app.listen(PORT, () => {
  console.log(`Server is listening on port: ${PORT}`);
});

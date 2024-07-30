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


dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

app.use(bodyParser.json());


// const data = [
//   { sNo: 1, name: "John Doe", age: 30, district: "New York" },
//   { sNo: 2, name: "Jane Smith", age: 25, district: "Los Angeles" }
// ];

// module.exports = data;

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
    res.send("Message sent");

  } else {
    res.send("Total Count exceeded");
    console.log("Total Count exceeded")
  }
});

app.post("/generate-document", async (req, res) => {
  const data = req.body;
  // console.log(req.body)
  // const table = new Table({
  //   rows: [
  //     new TableRow({
  //       children: [
  //         new TableCell({
  //           children: [new Paragraph("S.No")],
  //         }),
  //         new TableCell({
  //           children: [new Paragraph("Name")],
  //         }),
  //         new TableCell({
  //           children: [new Paragraph("Age")],
  //         }),
  //         new TableCell({
  //           children: [new Paragraph("District")],
  //         }),
  //       ],
  //     }),
  //     ...data.map((item: any) =>
  //       new TableRow({
  //         children: [
  //           new TableCell({
  //             children: [new Paragraph(item.sNo.toString())],
  //           }),
  //           new TableCell({
  //             children: [new Paragraph(item.name)],
  //           }),
  //           new TableCell({
  //             children: [new Paragraph(item.age.toString())],
  //           }),
  //           new TableCell({
  //             children: [new Paragraph(item.district)],
  //           }),
  //         ],
  //       })
  //     ),
  //   ],
  // });
  // const table = new Table({
  //   rows: [
  //     // Header row with styling
  //     new TableRow({
  //       children: [
  //         new TableCell({
  //           children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.No", bold: true })] })],
  //           borders: {
  //             top: { style: BorderStyle.SINGLE, size: 2 },
  //             bottom: { style: BorderStyle.SINGLE, size: 2 },
  //             left: { style: BorderStyle.SINGLE, size: 2 },
  //             right: { style: BorderStyle.SINGLE, size: 2 },
  //           },
  //           margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //         }),
  //         new TableCell({
  //           children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Name", bold: true })] })],
  //           borders: {
  //             top: { style: BorderStyle.SINGLE, size: 2 },
  //             bottom: { style: BorderStyle.SINGLE, size: 2 },
  //             left: { style: BorderStyle.SINGLE, size: 2 },
  //             right: { style: BorderStyle.SINGLE, size: 2 },
  //           },
  //           margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //         }),
  //         new TableCell({
  //           children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Age", bold: true })] })],
  //           borders: {
  //             top: { style: BorderStyle.SINGLE, size: 2 },
  //             bottom: { style: BorderStyle.SINGLE, size: 2 },
  //             left: { style: BorderStyle.SINGLE, size: 2 },
  //             right: { style: BorderStyle.SINGLE, size: 2 },
  //           },
  //           margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //         }),
  //         new TableCell({
  //           children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "District", bold: true })] })],
  //           borders: {
  //             top: { style: BorderStyle.SINGLE, size: 2 },
  //             bottom: { style: BorderStyle.SINGLE, size: 2 },
  //             left: { style: BorderStyle.SINGLE, size: 2 },
  //             right: { style: BorderStyle.SINGLE, size: 2 },
  //           },
  //           margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //         }),
  //       ],
  //     }),
  //     // Data rows with styling
  //     ...data.map((item: any) =>
  //       new TableRow({
  //         children: [
  //           new TableCell({
  //             children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.sNo.toString() })] })],
  //             borders: {
  //               top: { style: BorderStyle.SINGLE, size: 1 },
  //               bottom: { style: BorderStyle.SINGLE, size: 1 },
  //               left: { style: BorderStyle.SINGLE, size: 1 },
  //               right: { style: BorderStyle.SINGLE, size: 1 },
  //             },
  //             margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //           }),
  //           new TableCell({
  //             children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.name })] })],
  //             borders: {
  //               top: { style: BorderStyle.SINGLE, size: 1 },
  //               bottom: { style: BorderStyle.SINGLE, size: 1 },
  //               left: { style: BorderStyle.SINGLE, size: 1 },
  //               right: { style: BorderStyle.SINGLE, size: 1 },
  //             },
  //             margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //           }),
  //           new TableCell({
  //             children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.age.toString() })] })],
  //             borders: {
  //               top: { style: BorderStyle.SINGLE, size: 1 },
  //               bottom: { style: BorderStyle.SINGLE, size: 1 },
  //               left: { style: BorderStyle.SINGLE, size: 1 },
  //               right: { style: BorderStyle.SINGLE, size: 1 },
  //             },
  //             margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //           }),
  //           new TableCell({
  //             children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.district })] })],
  //             borders: {
  //               top: { style: BorderStyle.SINGLE, size: 1 },
  //               bottom: { style: BorderStyle.SINGLE, size: 1 },
  //               left: { style: BorderStyle.SINGLE, size: 1 },
  //               right: { style: BorderStyle.SINGLE, size: 1 },
  //             },
  //             margins: { top: 100, bottom: 100, left: 100, right: 100 },
  //           }),
  //         ],
  //       })
  //     ),
  //   ],
  //   width: {
  //     size: 10000,
  //     type: WidthType.DXA,
  //   },
  // });

  // const doc = new Document({
  //   sections: [
  //     {
  //       children: [table],
  //     },
  //   ],
  // });

  // const buffer = await Packer.toBuffer(doc);
  // Load the template document from the assets folder
  const templatePath = path.join(__dirname, "Assets", "Templates.docx");


  // // Set headers for file download
  // res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  // res.setHeader("Content-Disposition", "attachment; filename=GeneratedTemplate.docx");
  // res.send(buffer);

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

  // Create a new Docxtemplater instance
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // Replace placeholders with actual values
  // doc.render({
  //   User: data.title,
  //   price: data.price,
  //   details: data.details
  // });
  doc.render(data)
  // Generate the modified document
  const buf = doc.getZip().generate({ type: 'nodebuffer' });

  // Save the modified document to a new file
  fs.writeFileSync(outputPath, buf);
  // Set headers for file download
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", "attachment; filename=GeneratedTemplate.docx");
  res.send(buf);

  console.log('Document created successfully!');


  // // Step 1: Fetch the template file from SharePoint
  // async function getTemplateFile(accessToken: string, siteId: string, libraryId: string, itemId: string) {
  //   const fileEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${libraryId}/items/${itemId}/content`;

  //   try {
  //     const response = await axios.get(fileEndpoint, {
  //       headers: {
  //         Authorization: `Bearer ${accessToken}`,
  //       },
  //       responseType: 'arraybuffer', // Important to get the file as binary data
  //     });

  //     return response.data;
  //   } catch (error: any) {
  //     console.error("Error fetching template file:", error.response?.data || error.message);
  //     throw error;
  //   }
  // }

  // // Step 2: Upload the generated document back to SharePoint
  // async function uploadFileToSharePoint(accessToken: string, siteId: string, libraryId: string, fileName: string, fileContent: Buffer) {
  //   const uploadEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${libraryId}/root:/${fileName}:/content`;

  //   try {
  //     const response = await axios.put(uploadEndpoint, fileContent, {
  //       headers: {
  //         Authorization: `Bearer ${accessToken}`,
  //         'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  //       },
  //     });

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

  //   // Assume you have the item ID of the template file
  //   const templateItemId = "25"; // Replace with your template file ID in SharePoint

  //   // Get the template file from SharePoint
  //   const templateFileContent = await getTemplateFile(accessToken, siteId, libraryId, templateItemId);

  //   // Create a new PizZip instance to read the binary content
  //   const zip = new PizZip(templateFileContent);

  //   // Create a new Docxtemplater instance
  //   const doc = new Docxtemplater(zip, {
  //     paragraphLoop: true,
  //     linebreaks: true,
  //   });

  //   // Replace placeholders with actual values
  //   doc.render({
  //     User: data.title,
  //     price: data.price,
  //     details: data.details
  //   });

  //   // Generate the modified document
  //   const buf = doc.getZip().generate({ type: 'nodebuffer' });

  //   // Define the name for the generated document
  //   const generatedFileName = `GeneratedDocument_${Date.now()}.docx`;

  //   // Upload the generated document back to SharePoint
  //   const uploadedFile = await uploadFileToSharePoint(accessToken, siteId, libraryId, generatedFileName, buf);

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

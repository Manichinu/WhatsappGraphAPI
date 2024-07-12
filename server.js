"use strict";
/**
 * Copyright (c) Meta Platforms, Inc. and affiliates.
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 */
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var express_1 = require("express");
var axios_1 = require("axios");
var cors_1 = require("cors");
var dotenv_1 = require("dotenv");
dotenv_1.default.config();
var app = (0, express_1.default)();
app.use((0, cors_1.default)());
app.use(express_1.default.json());
var _a = process.env, WEBHOOK_VERIFY_TOKEN = _a.WEBHOOK_VERIFY_TOKEN, GRAPH_API_TOKEN = _a.GRAPH_API_TOKEN, PORT = _a.PORT;
app.post("/webhook", function (req, res) { return __awaiter(void 0, void 0, void 0, function () {
    var message, postItem, business_phone_number_id;
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u;
    return __generator(this, function (_v) {
        switch (_v.label) {
            case 0:
                // log incoming messages
                console.log("Incoming webhook message:", JSON.stringify(req.body, null, 2));
                message = (_e = (_d = (_c = (_b = (_a = req.body.entry) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b.changes[0]) === null || _c === void 0 ? void 0 : _c.value) === null || _d === void 0 ? void 0 : _d.messages) === null || _e === void 0 ? void 0 : _e[0];
                // postItem = {
                //     url: "https://prod-134.westus.logic.azure.com:443/workflows/54a65ff633684999b86c7d2067a4ed82/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KNvJymBCdH3oU-BzKHvXT3BXmWju_IzMyCNkCevuEwE",
                //     method: "POST",
                //     timeout: 0,
                //     headers: {
                //         'Access-Control-Allow-Origin': '*',
                //         "Accept": "application/json; odata=nometadata",
                //         "Content-Type": "application/json; odata=nometadata"
                //     },
                //     data: {
                //         Message: message === null || message === void 0 ? void 0 : message.text.body,
                //         Number: (_k = (_j = (_h = (_g = (_f = req.body.entry) === null || _f === void 0 ? void 0 : _f[0]) === null || _g === void 0 ? void 0 : _g.changes[0]) === null || _h === void 0 ? void 0 : _h.value) === null || _j === void 0 ? void 0 : _j.messages) === null || _k === void 0 ? void 0 : _k[0].from,
                //         Name: (_q = (_p = (_o = (_m = (_l = req.body.entry) === null || _l === void 0 ? void 0 : _l[0]) === null || _m === void 0 ? void 0 : _m.changes[0]) === null || _o === void 0 ? void 0 : _o.value) === null || _p === void 0 ? void 0 : _p.contacts) === null || _q === void 0 ? void 0 : _q[0].profile.name
                //     }
                // };
                // (0, axios_1.default)(postItem)
                //     .then(function (response) {
                //     console.log('Response:', response.data);
                // })
                //     .catch(function (error) {
                //     console.error('Error:', error);
                // });
                if (!((message === null || message === void 0 ? void 0 : message.type) === "text")) return [3 /*break*/, 2];
                business_phone_number_id = (_u = (_t = (_s = (_r = req.body.entry) === null || _r === void 0 ? void 0 : _r[0].changes) === null || _s === void 0 ? void 0 : _s[0].value) === null || _t === void 0 ? void 0 : _t.metadata) === null || _u === void 0 ? void 0 : _u.phone_number_id;
                // send a reply message as per the docs here https://developers.facebook.com/docs/whatsapp/cloud-api/reference/messages
                // await axios({
                //   method: "POST",
                //   url: `https://graph.facebook.com/v18.0/${business_phone_number_id}/messages`,
                //   headers: {
                //     Authorization: `Bearer ${GRAPH_API_TOKEN}`,
                //   },
                //   data: {
                //     messaging_product: "whatsapp",
                //     to: message.from,
                //     text: { body: "Echo: " + message.text.body },
                //     context: {
                //       message_id: message.id, // shows the message as a reply to the original user message
                //     },
                //   },
                // });
                // mark incoming message as read
                return [4 /*yield*/, (0, axios_1.default)({
                        method: "POST",
                        url: "https://graph.facebook.com/v18.0/".concat(business_phone_number_id, "/messages"),
                        headers: {
                            Authorization: "Bearer ".concat(GRAPH_API_TOKEN),
                        },
                        data: {
                            messaging_product: "whatsapp",
                            status: "read",
                            message_id: message.id,
                        },
                    })];
            case 1:
                // send a reply message as per the docs here https://developers.facebook.com/docs/whatsapp/cloud-api/reference/messages
                // await axios({
                //   method: "POST",
                //   url: `https://graph.facebook.com/v18.0/${business_phone_number_id}/messages`,
                //   headers: {
                //     Authorization: `Bearer ${GRAPH_API_TOKEN}`,
                //   },
                //   data: {
                //     messaging_product: "whatsapp",
                //     to: message.from,
                //     text: { body: "Echo: " + message.text.body },
                //     context: {
                //       message_id: message.id, // shows the message as a reply to the original user message
                //     },
                //   },
                // });
                // mark incoming message as read
                _v.sent();
                _v.label = 2;
            case 2:
                res.sendStatus(200);
                return [2 /*return*/];
        }
    });
}); });
// accepts GET requests at the /webhook endpoint. You need this URL to setup webhook initially.
// info on verification request payload: https://developers.facebook.com/docs/graph-api/webhooks/getting-started#verification-requests
app.get("/webhook", function (req, res) {
    var mode = req.query["hub.mode"];
    var token = req.query["hub.verify_token"];
    var challenge = req.query["hub.challenge"];
    // check the mode and token sent are correct
    if (mode === "subscribe" && token === WEBHOOK_VERIFY_TOKEN) {
        // respond with 200 OK and challenge token from the request
        res.status(200).send(challenge);
        console.log("Webhook verified successfully!");
    }
    else {
        // respond with '403 Forbidden' if verify tokens do not match
        res.sendStatus(403);
    }
});
app.get("/", function (req, res) {
    res.send("<pre>Nothing to see here.\nCheckout README.md to start.</pre>");
});
app.listen(PORT, function () {
    console.log("Server is listening on port: ".concat(PORT));
});

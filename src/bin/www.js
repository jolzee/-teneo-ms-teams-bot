/**
 * Copyright 2018 Artificial Solutions. All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

const restify = require("restify");
const dotenv = require("dotenv");
import "core-js/stable";
import "regenerator-runtime/runtime";
dotenv.config();

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
  BotFrameworkAdapter,
  MemoryStorage,
  ConversationState,
  UserState,
} = require("botbuilder");

// prevent ReferenceError: Headers is not defined
const fetch = require("node-fetch");
global.Headers = fetch.Headers;

// This bot's main dialog.
const { TeneoBot } = require("../bots/teneoBot");
const { MainAuthDialog } = require("../authDialogs/mainAuthDialog");

// Define state store for your bot.
// See https://aka.ms/about-bot-state to learn more about bot state.
const memoryStorage = new MemoryStorage();

// Create conversation state with in-memory storage provider.
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

const mainAuthDialog = new MainAuthDialog();
const conversationReferences = {};

// Create HTTP server
let server = restify.createServer();
server.use(restify.plugins.queryParser());

server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log(`${server.name} listening to ${server.url}`);
});

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about .bot file its use and bot configuration .
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
});

// Create the main dialog.
const teneoBot = new TeneoBot(
  conversationReferences,
  conversationState,
  userState,
  mainAuthDialog
);

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  );
  // Clear out state
  await conversationState.delete(context);
};

// Listen for incoming requests.
server.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    // Route to main dialog.
    // await teneoBot.onTurn(context);
    await teneoBot.run(context);
  });
});

// Listen for incoming notifications and send proactive messages to users.
server.get("/api/notify", async (req, res) => {
  for (const conversationReference of Object.values(conversationReferences)) {
    await adapter.continueConversation(
      conversationReference,
      async (turnContext) => {
        // const notificationMessage = req.query.msg;
        // If you encounter permission-related errors when sending this message, see
        // https://aka.ms/BotTrustServiceUrl
        await turnContext.sendActivity(
          "This is a proactive message... If you eat something & nobody see you eat it, it has no calories."
        );
      }
    );
  }

  res.setHeader("Content-Type", "text/html");
  res.writeHead(200);
  res.write(
    "<html><body><h1>Proactive messages have been sent.</h1></body></html>"
  );
  res.end();
});

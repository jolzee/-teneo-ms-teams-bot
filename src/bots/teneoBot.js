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

const { DialogBot } = require("./dialogBot");
const { ActivityTypes } = require("botbuilder");
const TIE = require("@artificialsolutions/tie-api-client");
const dotenv = require("dotenv");
dotenv.config();

// Teneo engine url
const teneoEngineUrl = process.env.TENEO_ENGINE_URL;
const googleSheetId = process.env.GOOGLE_SHEET_ID;

// property to store sessionId in conversation state object
const SESSION_ID_PROPERTY = "sessionId";

const WELCOMED_USER = "welcomedUserProperty";

// initialize a Teneo client for interacting with TeneoEengine
const teneoApi = TIE.init(teneoEngineUrl);

class TeneoBot extends DialogBot {
  /**
   *
   * @param {ConversationState} conversation state object
   */
  constructor(
    conversationReferences,
    conversationState,
    userState,
    mainAuthDialog
  ) {
    super(conversationState, userState, mainAuthDialog);

    // Dependency injected dictionary for storing ConversationReference objects used in NotifyController to proactively message users
    this.conversationReferences = conversationReferences;

    // Creates a new state accessor property.
    // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.
    this.sessionIdProperty = conversationState.createProperty(
      SESSION_ID_PROPERTY
    );
    this.welcomedUserProperty = conversationState.createProperty(WELCOMED_USER);

    this.conversationState = conversationState;

    this.onConversationUpdate(async (context, next) => {
      this.addConversationReference(context.activity);
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id !== context.activity.recipient.id) {
          const welcomeMessage =
            "Proactive Greeting. Hi! In future I will expose a url '/api/notify?msg=This is a notification' that will proactively message everyone who has previously messaged this bot.";
          await context.sendActivity(welcomeMessage);
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMessage(async (context, next) => {
      this.addConversationReference(context.activity);

      // Echo back what the user said
      await context.sendActivity(`You sent '${context.activity.text}'`);
      await next();
    });
  }

  /**
   * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
   */
  async run(context) {
    await super.run(context);
    this.onTurn(context);
  }
  /**
   *
   * @param {TurnContext} on turn context object.
   */
  async onTurn(turnContext) {
    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.

    if (turnContext.activity.type === ActivityTypes.Message) {
      // send user input to engine and store sessionId in state in case not stored yet
      await this.handleMessage(turnContext);
    } else if (turnContext.activity.type === ActivityTypes.ConversationUpdate) {
      // Conversation update activities describe a change in a conversation's members, description, existence, or otherwise.
      // We want to send a welcome message to conversation members when they join the conversation

      console.log(turnContext.activity.membersAdded);
      // Iterate over all new members added to the conversation
      for (const idx in turnContext.activity.membersAdded) {
        // Only sent message to conversation members who aren't the bot
        if (
          turnContext.activity.membersAdded[idx].id !==
          turnContext.activity.recipient.id
        ) {
          // send empty input to engine to receive Teneo greeting message and store sessionId in state
          await this.handleMessage(turnContext);
        }
      }
    } else {
      console.log(`[${turnContext.activity.type} event detected]`);
    }
    // Save state changes
    await this.conversationState.saveChanges(turnContext);
  }

  getChunks(answerText) {
    let finalAnswerText = "";
    const chunks = answerText.split("||");
    chunks.forEach((chunk) => {
      const trimmedChunk = chunk.trim();
      if (trimmedChunk) {
        finalAnswerText += `||${trimmedChunk}`;
      }
    });
    if (finalAnswerText.startsWith("||")) {
      finalAnswerText = finalAnswerText.substring(2);
    }
    return finalAnswerText.trim().split("||");
  }

  addConversationReference(activity) {
    const conversationReference = TurnContext.getConversationReference(
      activity
    );
    this.conversationReferences[
      conversationReference.conversation.id
    ] = conversationReference;
  }

  /**
   *
   * @param {TurnContext} on turn context object.
   */
  async handleMessage(turnContext) {
    const fullName = turnContext._activity.from.name; // Full name
    const firstName = fullName.split(" ")[0];
    const lastName = fullName.substring(
      0,
      fullName.length > firstName.length
        ? firstName.length + 1
        : firstName.length
    );

    const country = turnContext._activity.entities[0].country; // something like "US"
    const locale = turnContext._activity.entities[0].locale;

    // console.log(`Turn Context Info:`, JSON.stringify(turnContext, null, 2));
    const message = turnContext.activity;
    // console.log(message);
    try {
      let messageText = "";
      if (message.text) {
        messageText = message.text;
      }

      console.log(
        `Got message '${messageText}' from channel ${message.channelId}`
      );

      // find engine session id
      const sessionId = await this.sessionIdProperty.get(turnContext);

      let inputDetails = {
        text: messageText,
        channel: "botframework-" + message.channelId,
        sheetId: googleSheetId,
        displayName: fullName,
        lastName: lastName,
        givenName: firstName,
        name: firstName,
        countryCode: country,
        locale: locale,
      };

      if (message.attachments) {
        inputDetails["botframeworkAttachments"] = JSON.stringify(
          message.attachments
        );
      }

      // send message to engine using sessionId
      const teneoResponse = await teneoApi.sendInput(sessionId, inputDetails);

      console.log(
        `Got Teneo Engine response '${teneoResponse.output.text}' for session ${teneoResponse.sessionId}`
      );

      // store egnine sessionId in conversation state
      await this.sessionIdProperty.set(turnContext, teneoResponse.sessionId);

      // split the reply by possible chunk separator || reply text from engine
      let chunks = this.getChunks(teneoResponse.output.text);

      for (let index = 0; index < chunks.length; index++) {
        const chunkAnswer = chunks[index];
        const reply = {};

        reply.text = chunkAnswer;

        // only send bot framework actions for the last chunk
        if (index + 1 === chunks.length) {
          // check if an output parameter 'msbotframework' exists in engine response
          // if so, check if it should be added as attachment/card or suggestion action
          if (teneoResponse.output.parameters.msbotframework) {
            try {
              const extension = JSON.parse(
                teneoResponse.output.parameters.msbotframework
              );

              // suggested actions have an 'actions' key
              if (extension.actions) {
                reply.suggestedActions = extension;
              } else {
                // we assume the extension code matches that of an attachment or rich card
                reply.attachments = [extension];
              }
            } catch (attachError) {
              console.error(`Failed when parsing attachment JSON`, attachError);
            }
          }
        }

        // send response to bot framework.
        await turnContext.sendActivity(reply);
      }
    } catch (error) {
      console.error(
        `Failed when sending input to Teneo Engine @ ${teneoEngineUrl}`,
        error
      );
    }
  }
}

module.exports.TeneoBot = TeneoBot;

'use strict';

const path = require('path');
const fs = require('fs');
const dotenv = require('dotenv');

const _root = path.join(__dirname, '..');
const _envPath = path.join(_root, '.env');
if (fs.existsSync(_envPath)) {
  dotenv.config({ path: _envPath });
} else {
  dotenv.config();
}

const { BotFrameworkAdapter, ActivityHandler } = require('botbuilder');
const { getBotChannelAuthTenant } = require('./botChannelTenant');
const { handleTeamsMessage } = require('./botMessageHandlers');

class TeamsAdoBot extends ActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      try {
        await handleTeamsMessage(context);
      } catch (err) {
        console.error(err);
        await context.sendActivity(`Error: ${err.message}`);
      }
      await next();
    });
  }
}

let adapter;
let bot;

function getAdapter() {
  if (!adapter) {
    const tenantId = getBotChannelAuthTenant();
    adapter = new BotFrameworkAdapter({
      appId: process.env.MicrosoftAppId,
      appPassword: process.env.MicrosoftAppPassword,
      ...(tenantId ? { channelAuthTenant: tenantId } : {}),
    });
  }
  return adapter;
}

function getBot() {
  if (!bot) {
    bot = new TeamsAdoBot();
  }
  return bot;
}

/**
 * Bot Framework webhook: pass Node-style req/res (Restify, Express, Vercel).
 * @param {import('http').IncomingMessage} req
 * @param {import('http').ServerResponse} res
 */
async function processBotMessage(req, res) {
  const ad = getAdapter();
  const b = getBot();
  await ad.processActivity(req, res, (context) => b.run(context));
}

module.exports = {
  processBotMessage,
  TeamsAdoBot,
};

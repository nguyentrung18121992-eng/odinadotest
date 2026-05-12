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

const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');
const { getBotChannelAuthTenant } = require('./botChannelTenant');

function requireEnv(name) {
  const v = process.env[name];
  if (!v) {
    throw new Error(`Missing environment variable: ${name}`);
  }
  return v;
}

/**
 * Proactive Teams message via Bot Framework Connector (same app as the bot).
 * Use conversation id from an incoming activity (e.g. channel thread: 19:…@thread.tacv2).
 */
class TeamsOutreach {
  constructor(opts = {}) {
    this.appId = opts.appId || requireEnv('MicrosoftAppId');
    this.appPassword = opts.appPassword || requireEnv('MicrosoftAppPassword');
    this.tenantId = opts.tenantId || getBotChannelAuthTenant();
    const base =
      opts.serviceUrl ||
      process.env.TEAMS_SERVICE_URL ||
      'https://smba.trafficmanager.net/teams/';
    this.serviceUrl = `${base.replace(/\/?$/, '/')}`;
  }

  /**
   * @param {string} chatId - Bot conversation id (often 19:…@thread.tacv2 for a channel thread).
   * @param {string} message - Text; markdown if textFormat is markdown.
   * @returns {Promise<{ success: boolean, error?: string, id?: string }>}
   */
  async sendToBot(chatId, message) {
    try {
      const credentials = new MicrosoftAppCredentials(
        this.appId,
        this.appPassword,
        this.tenantId
      );
      const client = new ConnectorClient(credentials, { baseUri: this.serviceUrl });
      const activity = {
        type: 'message',
        text: message,
        textFormat: 'markdown',
      };
      const response = await client.conversations.sendToConversation(
        chatId,
        activity
      );
      return { success: true, id: response?.id };
    } catch (err) {
      const body =
        err.details?.body ||
        err.body ||
        err.response?.data ||
        err.response?.body;
      const extra =
        typeof body === 'string'
          ? body
          : body
            ? JSON.stringify(body)
            : '';
      const msg = [err.message || String(err), extra].filter(Boolean).join(' — ');
      return { success: false, error: msg };
    }
  }
}

module.exports = { TeamsOutreach };

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
const ado = require('./adoBoardsClient');

const defaultWorkItemType = process.env.ADO_DEFAULT_WORK_ITEM_TYPE || 'Task';

class TeamsAdoBot extends ActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      const text = (context.activity.text || '').trim();
      const lower = text.toLowerCase();

      try {
        if (lower.startsWith('create task ')) {
          const title = text.slice('create task '.length).trim();
          if (!title) {
            await context.sendActivity('Usage: create task <title>');
            return;
          }
          const wi = await ado.createWorkItem({
            workItemType: defaultWorkItemType,
            fields: {
              title,
              'System.Description': 'Created from Microsoft Teams bot.',
            },
          });
          const url = ado.workItemWebUrl(wi);
          await context.sendActivity(
            `Created work item **#${wi.id}**${url ? `\n${url}` : ''}`
          );
        } else if (lower.startsWith('update ')) {
          const rest = text.slice(7).trim();
          const match = rest.match(/^(\d+)\s+(.+)$/);
          if (!match) {
            await context.sendActivity(
              'Usage: `update <id> <state>` — example: `update 123 Active`'
            );
            return;
          }
          const id = Number(match[1], 10);
          const state = match[2].trim();
          const wi = await ado.updateWorkItem({
            id,
            fields: { 'System.State': state },
          });
          const url = ado.workItemWebUrl(wi);
          await context.sendActivity(
            `Updated **#${wi.id}** → \`${state}\`${url ? `\n${url}` : ''}`
          );
        } else {
          await context.sendActivity(
            [
              'Commands:',
              '- `create task <title>` — new board work item',
              '- `update <id> <state>` — set work item state',
            ].join('\n')
          );
        }
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
  // Use processActivity, not process(): process() runs ResponseT Zod on `res`; plain Node
  // ServerResponse fails unless fully shimmed. Local Restify responses work with either.
  await ad.processActivity(req, res, (context) => b.run(context));
}

module.exports = {
  processBotMessage,
  TeamsAdoBot,
};

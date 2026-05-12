'use strict';

require('dotenv').config();

const restify = require('restify');
const { BotFrameworkAdapter, ActivityHandler } = require('botbuilder');
const ado = require('./services/adoBoardsClient');

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
});

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

const bot = new TeamsAdoBot();

server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

const port = Number(process.env.PORT, 10) || 3978;
server.once('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(
      `Port ${port} is already in use. Close the other process or change PORT in .env.\n` +
        `Hint (Windows): netstat -ano | findstr ":${port}"  →  taskkill /PID <pid> /F`
    );
    process.exit(1);
    return;
  }
  console.error(err);
  process.exit(1);
});

server.listen(port, () => {
  console.log(`${server.name} listening on ${port}`);
});

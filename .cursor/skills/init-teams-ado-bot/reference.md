# Reference templates (greenfield repo)

**Cursor:** read `SKILL.md` in this skill folder first (when, description, verification loop), then use this file for pastable templates.

Use these as **verbatim** starting points in a **new** repository root. Paths are relative to repo root. Create folders `api/`, `services/`, `scripts/`, `public/` as needed.

After writing files, run `npm install` and `npm run verify` from the repo root.

---

## `package.json`

```json
{
  "name": "teamsado-bot",
  "version": "1.0.0",
  "private": true,
  "description": "Microsoft Teams bot → Azure Bot Service → Azure DevOps Boards",
  "main": "index.js",
  "engines": {
    "node": ">=18"
  },
  "scripts": {
    "start": "node --disable-warning=DEP0169 index.js",
    "dev": "node --disable-warning=DEP0169 --watch index.js",
    "smoke": "node --disable-warning=DEP0169 scripts/smoke-ado.js",
    "verify": "node --disable-warning=DEP0169 scripts/verify.js",
    "build": "npm run verify"
  },
  "dependencies": {
    "azure-devops-node-api": "^15.1.2",
    "botbuilder": "^4.23.1",
    "dotenv": "^16.4.5",
    "restify": "^11.1.0"
  }
}
```

---

## `vercel.json`

```json
{
  "buildCommand": "npm run build",
  "outputDirectory": "public",
  "env": {
    "NODE_OPTIONS": "--disable-warning=DEP0169"
  },
  "functions": {
    "api/messages.js": {
      "maxDuration": 10,
      "includeFiles": "services/**"
    }
  }
}
```

Commit the `public/` folder (at least `index.html`) so a `public` directory exists after `npm run build`. Vercel reports an error if **Project Settings → Output Directory** is `public` (or `vercel.json` sets it) but that folder is missing post-build.

---

## `public/index.html`

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Teams ADO bot</title>
    <style>
      body {
        font-family: system-ui, sans-serif;
        max-width: 36rem;
        margin: 2rem auto;
        padding: 0 1rem;
        line-height: 1.5;
      }
      code {
        background: #f0f0f0;
        padding: 0.15rem 0.35rem;
        border-radius: 4px;
      }
    </style>
  </head>
  <body>
    <h1>Teams ADO bot</h1>
    <p>
      Static placeholder so Vercel can use <code>public</code> as the output directory. The
      Azure Bot messaging endpoint is
      <code>/api/messages</code>
      (configure the full HTTPS URL in Azure Bot).
    </p>
    <p><a href="/api/messages">GET /api/messages</a> — health text</p>
  </body>
</html>
```

---

## `.gitignore`

```gitignore
# Local secrets — copy from `.env.example`
.env
.env.local
.env.*.local

node_modules/
npm-debug.log*
.DS_Store
```

---

## `.env.example`

```
# Copy this file to `.env` and fill in values. Never commit `.env`.
# On Vercel: Project Settings → Environment Variables (same names). Bot URL: https://<project>.vercel.app/api/messages

# --- Azure DevOps (Azure Boards) ---
ADO_ORG=

ADO_PROJECT=

ADO_PAT=

# ADO_SERVER_URL=https://dev.azure.com/myorg

ADO_DEFAULT_WORK_ITEM_TYPE=Task

# ADO_SEARCH_TOP=20

# --- Microsoft Bot Framework (Azure Bot / Teams) ---
MicrosoftAppId=
MicrosoftAppPassword=

# MicrosoftAppTenantId=
# MICROSOFT_APP_TENANT_ID=
# AZURE_TENANT_ID=

PORT=3978
```

---

## `index.js`

```javascript
'use strict';

const { processBotMessage } = require('./services/botApp');
const restify = require('restify');

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', async (req, res) => {
  await processBotMessage(req, res);
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
```

---

## `api/messages.js`

```javascript
'use strict';

/** Load ./services/botApp first so repo-root `.env` is applied (inline bootstrap there). */
const { processBotMessage } = require('../services/botApp');
const { wrapServerResponse } = require('../services/expressLikeServerResponse');

/**
 * Vercel Serverless Function — messaging endpoint for Azure Bot / Teams.
 * Configure Azure Bot messaging URL: https://<your-project>.vercel.app/api/messages
 */
module.exports = async (req, res) => {
  if (req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(
      'Teams ADO bot — POST JSON activities to this URL (Azure Bot messaging endpoint).'
    );
    return;
  }

  if (req.method !== 'POST') {
    res.writeHead(405, { Allow: 'GET, POST' });
    res.end('Method Not Allowed');
    return;
  }

  if (typeof req.body === 'string') {
    try {
      req.body = JSON.parse(req.body);
    } catch {
      res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Invalid JSON body');
      return;
    }
  }

  try {
    await processBotMessage(req, wrapServerResponse(res));
  } catch (err) {
    console.error(err);
    if (!res.headersSent) {
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Internal Server Error');
    }
  }
};
```

---

## `scripts/verify.js`

```javascript
#!/usr/bin/env node
'use strict';

/**
 * Loads all serverless + bot modules without starting HTTP. Exits 0 on success.
 */
require('../services/botApp');
require('../api/messages');
console.log('verify: ok');
```

---

## `services/botChannelTenant.js`

```javascript
'use strict';

/**
 * Azure AD tenant (directory) ID for Bot Framework outbound auth (MSAL).
 * @returns {string|undefined}
 */
function getBotChannelAuthTenant() {
  const t =
    process.env.MicrosoftAppTenantId ||
    process.env.MICROSOFT_APP_TENANT_ID ||
    process.env.AZURE_TENANT_ID;
  const v = t != null ? String(t).trim() : '';
  return v || undefined;
}

module.exports = { getBotChannelAuthTenant };
```

---

## `services/expressLikeServerResponse.js`

```javascript
'use strict';

/**
 * Express/Restify-style response wrapper for Node.js http.ServerResponse.
 * Bot Framework processActivity uses res.status(), res.send(), res.end();
 * @param {import('http').ServerResponse} raw
 */
function wrapServerResponse(raw) {
  let statusCode = 200;
  /** @type {string | null} */
  let payload = null;
  let json = false;

  return {
    status(code) {
      statusCode = code;
      return this;
    },
    header(name, value) {
      if (value === undefined) {
        return raw.getHeader(name);
      }
      raw.setHeader(name, value);
      return this;
    },
    send(body) {
      if (body === undefined || body === null) {
        return this;
      }
      if (typeof body === 'object' && !Buffer.isBuffer(body)) {
        payload = JSON.stringify(body);
        json = true;
      } else {
        payload = typeof body === 'string' ? body : String(body);
        json = false;
      }
      return this;
    },
    end(chunk) {
      if (chunk !== undefined && chunk !== null) {
        raw.statusCode = statusCode;
        if (!raw.getHeader('Content-Type')) {
          raw.setHeader('Content-Type', 'text/plain; charset=utf-8');
        }
        raw.end(chunk);
        return this;
      }
      raw.statusCode = statusCode;
      if (payload !== null) {
        if (json) {
          raw.setHeader('Content-Type', 'application/json');
        } else if (!raw.getHeader('Content-Type')) {
          raw.setHeader('Content-Type', 'text/plain; charset=utf-8');
        }
        raw.end(payload);
      } else {
        raw.end();
      }
      return this;
    },
  };
}

module.exports = { wrapServerResponse };
```

---

## `services/adoBoardsClient.js`

```javascript
'use strict';

const azdev = require('azure-devops-node-api');

/** @type {import('azure-devops-node-api').IWorkItemTrackingApi | null} */
let cachedWit = null;

function requireEnv(name) {
  const v = process.env[name];
  if (!v) {
    throw new Error(`Missing environment variable: ${name}`);
  }
  return v;
}

function getOrgUrl() {
  if (process.env.ADO_SERVER_URL) {
    return process.env.ADO_SERVER_URL.replace(/\/$/, '');
  }
  const org = requireEnv('ADO_ORG');
  return `https://dev.azure.com/${org}`;
}

function getConnection() {
  const token = requireEnv('ADO_PAT');
  const auth = azdev.getPersonalAccessTokenHandler(token);
  return new azdev.WebApi(getOrgUrl(), auth);
}

async function getWorkItemTrackingApi() {
  if (!cachedWit) {
    const connection = getConnection();
    cachedWit = await connection.getWorkItemTrackingApi();
  }
  return cachedWit;
}

function toPatchOperations(fields, op) {
  return Object.entries(fields).map(([key, value]) => {
    const path = key.startsWith('/fields/') ? key : `/fields/${key}`;
    return { op, path, value };
  });
}

async function createWorkItem({ workItemType, fields = {} }) {
  const project = requireEnv('ADO_PROJECT');
  const wit = await getWorkItemTrackingApi();
  const merged = { ...fields };
  if (merged.title != null && merged['System.Title'] == null) {
    merged['System.Title'] = merged.title;
  }
  delete merged.title;
  if (!merged['System.Title']) {
    throw new Error('createWorkItem requires fields.System.Title or fields.title');
  }
  const document = toPatchOperations(merged, 'add');
  return wit.createWorkItem(
    null,
    document,
    project,
    workItemType,
    false,
    false,
    false
  );
}

async function updateWorkItem({ id, fields, project }) {
  const wit = await getWorkItemTrackingApi();
  const proj = project ?? requireEnv('ADO_PROJECT');
  const document = toPatchOperations(fields, 'replace');
  return wit.updateWorkItem(null, document, id, proj, false, false, false);
}

async function deleteWorkItem({ id, destroy = false, project }) {
  const wit = await getWorkItemTrackingApi();
  const proj = project ?? requireEnv('ADO_PROJECT');
  await wit.deleteWorkItem(id, proj, destroy);
}

function workItemWebUrl(workItem) {
  return workItem?._links?.html?.href;
}

const SEARCH_FIELDS = [
  'System.Id',
  'System.Title',
  'System.State',
  'System.WorkItemType',
  'System.AssignedTo',
];

function escapeWiqlLiteral(value) {
  return String(value).replace(/'/g, "''");
}

function getSearchDefaultTop() {
  const n = Number(process.env.ADO_SEARCH_TOP, 10);
  if (Number.isFinite(n) && n > 0) {
    return Math.min(Math.floor(n), 50);
  }
  return 20;
}

async function searchWorkItems({ text, top, project }) {
  const proj = project ?? requireEnv('ADO_PROJECT');
  const wit = await getWorkItemTrackingApi();
  const limit = Math.min(Math.max(1, top ?? getSearchDefaultTop()), 50);
  const trimmed = (text || '').trim();
  if (!trimmed) {
    throw new Error('searchWorkItems requires non-empty text');
  }

  const projectLit = escapeWiqlLiteral(proj);
  let wiql;
  if (/^\d+$/.test(trimmed)) {
    wiql = `SELECT [System.Id]
FROM WorkItems
WHERE [System.TeamProject] = '${projectLit}'
  AND [System.Id] = ${trimmed}`;
  } else {
    const phrase = escapeWiqlLiteral(trimmed.replace(/\s+/g, ' '));
    wiql = `SELECT [System.Id]
FROM WorkItems
WHERE [System.TeamProject] = '${projectLit}'
  AND (
    [System.Title] CONTAINS WORDS '${phrase}'
    OR [System.Description] CONTAINS WORDS '${phrase}'
  )
ORDER BY [System.ChangedDate] DESC`;
  }

  const result = await wit.queryByWiql(
    { query: wiql },
    { project: proj },
    false,
    limit
  );

  const refs = result.workItems || [];
  const ids = refs.map((r) => r.id).filter((id) => id != null);
  if (ids.length === 0) {
    return [];
  }

  const items = await wit.getWorkItems(
    ids,
    SEARCH_FIELDS,
    undefined,
    undefined,
    undefined,
    proj
  );
  const list = items || [];
  const byId = new Map(list.map((wi) => [wi.id, wi]));
  return ids.map((id) => byId.get(id)).filter(Boolean);
}

function resetConnectionCache() {
  cachedWit = null;
}

module.exports = {
  getWorkItemTrackingApi,
  createWorkItem,
  updateWorkItem,
  deleteWorkItem,
  searchWorkItems,
  workItemWebUrl,
  resetConnectionCache,
};
```

---

## `services/botApp.js`

```javascript
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

async function processBotMessage(req, res) {
  const ad = getAdapter();
  const b = getBot();
  await ad.processActivity(req, res, (context) => b.run(context));
}

module.exports = {
  processBotMessage,
  TeamsAdoBot,
};
```

---

## `services/botMessageHandlers.js`

```javascript
'use strict';

const ado = require('./adoBoardsClient');

const defaultWorkItemType = process.env.ADO_DEFAULT_WORK_ITEM_TYPE || 'Task';

function parseIdAndRest(remainder) {
  const m = remainder.trim().match(/^(\d+)\s+([\s\S]+)$/);
  if (!m) {
    return null;
  }
  return { id: Number(m[1], 10), rest: m[2].trim() };
}

function parseIdOnly(remainder) {
  const m = remainder.trim().match(/^(\d+)$/);
  if (!m) {
    return null;
  }
  return Number(m[1], 10);
}

function helpText() {
  return [
    '**Commands**',
    '- `create task <title>` — new work item',
    '- `title <id> <new title>` — rename',
    '- `description <id> <text>` — set description',
    '- `assign <id> <email or name>` — set assignee',
    '- `unassign <id>` — clear assignee',
    '- `state <id> <state>` — change state (e.g. Active, Closed)',
    '- `update <id> <state>` — same as `state`',
    '- `search <keywords>` / `find <keywords>` — title & description (WIQL CONTAINS WORDS; optional ADO_SEARCH_TOP, max 50)',
    '- `search <id>` — when id is digits only, that work item in this project',
    '- `delete <id>` / `delete task <id>` — delete work item',
  ].join('\n');
}

const MAX_SEARCH_TITLE_CHARS = 120;

function fieldDisplay(wi, refName) {
  const f = wi.fields?.[refName];
  if (f == null) {
    return '';
  }
  if (typeof f === 'object' && f.displayName != null) {
    return String(f.displayName);
  }
  return String(f);
}

function truncate(s, max) {
  if (s.length <= max) {
    return s;
  }
  return `${s.slice(0, max - 1)}…`;
}

async function handleCreateTask(context, text) {
  const m = text.match(/^create task\s+/i);
  if (!m) {
    return false;
  }
  const title = text.slice(m[0].length).trim();
  if (!title) {
    await context.sendActivity('Usage: `create task <title>`');
    return true;
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
  return true;
}

async function handleSearch(context, text, lower) {
  const mSearch = text.match(/^search\s+/i);
  const mFind = text.match(/^find\s+/i);
  const m = mSearch || mFind;
  if (!m) {
    return false;
  }
  const q = text.slice(m[0].length).trim();
  if (!q) {
    await context.sendActivity(
      'Usage: `search <keywords>` or `find <keywords>` — open the link on each row for full description.'
    );
    return true;
  }
  try {
    const items = await ado.searchWorkItems({ text: q });
    if (items.length === 0) {
      await context.sendActivity(`No work items matched **${q}** in this project.`);
      return true;
    }
    const lines = items.map((wi) => {
      const id = wi.id;
      const wtype = fieldDisplay(wi, 'System.WorkItemType') || 'Work item';
      const state = fieldDisplay(wi, 'System.State') || '—';
      const title = truncate(
        fieldDisplay(wi, 'System.Title') || '(no title)',
        MAX_SEARCH_TITLE_CHARS
      );
      const assignee = fieldDisplay(wi, 'System.AssignedTo');
      const assignBit = assignee ? ` · ${assignee}` : '';
      const url = ado.workItemWebUrl(wi);
      const link = url ? `\n${url}` : '';
      return `• **#${id}** [${wtype}] *${state}* — ${title}${assignBit}${link}`;
    });
    const header = `**${items.length}** result(s) for \`${q}\``;
    await context.sendActivity([header, ...lines].join('\n'));
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    await context.sendActivity(`Search failed: ${msg}`);
  }
  return true;
}

async function handleDeleteTask(context, text, lower) {
  let id = null;
  if (lower.startsWith('delete task')) {
    const dm = text.match(/^delete\s+task\s+(\d+)\s*$/i);
    if (!dm) {
      await context.sendActivity('Usage: `delete task <id>`');
      return true;
    }
    id = Number(dm[1], 10);
  } else if (lower.startsWith('delete ')) {
    const dm = text.match(/^delete\s+(\d+)\s*$/i);
    if (!dm) {
      await context.sendActivity('Usage: `delete <id>`');
      return true;
    }
    id = Number(dm[1], 10);
  }
  if (id == null) {
    return false;
  }
  await ado.deleteWorkItem({ id, destroy: false });
  await context.sendActivity(`Deleted work item **#${id}** (sent to Recycle Bin).`);
  return true;
}

async function handleTitle(context, text, lower) {
  const m = text.match(/^title\s+/i);
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity('Usage: `title <id> <new title>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.Title': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Updated title on **#${wi.id}**${url ? `\n${url}` : ''}`
  );
  return true;
}

async function handleDescription(context, text, lower) {
  const m = text.match(/^description\s+/i);
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity('Usage: `description <id> <text>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.Description': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Updated description on **#${wi.id}**${url ? `\n${url}` : ''}`
  );
  return true;
}

async function handleAssign(context, text, lower) {
  const m = text.match(/^assign\s+/i);
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity('Usage: `assign <id> <email or display name>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.AssignedTo': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Assigned **#${wi.id}** to \`${parsed.rest}\`${url ? `\n${url}` : ''}`
  );
  return true;
}

async function handleUnassign(context, text, lower) {
  const m = text.match(/^unassign\s+/i);
  if (!m) {
    return false;
  }
  const id = parseIdOnly(text.slice(m[0].length));
  if (id == null || Number.isNaN(id)) {
    await context.sendActivity('Usage: `unassign <id>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id,
    fields: { 'System.AssignedTo': '' },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Unassigned **#${wi.id}**${url ? `\n${url}` : ''}`
  );
  return true;
}

async function handleState(context, text, lower) {
  const mState = text.match(/^state\s+/i);
  const mUpdate = text.match(/^update\s+/i);
  const m = mState || mUpdate;
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity(
      'Usage: `state <id> <state>` — example: `state 123 Active`'
    );
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.State': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `**#${wi.id}** → state \`${parsed.rest}\`${url ? `\n${url}` : ''}`
  );
  return true;
}

const handlers = [
  handleCreateTask,
  handleSearch,
  handleDeleteTask,
  handleTitle,
  handleDescription,
  handleAssign,
  handleUnassign,
  handleState,
];

async function handleTeamsMessage(context) {
  const text = (context.activity.text || '').trim();
  if (!text) {
    await context.sendActivity(helpText());
    return;
  }
  const lower = text.toLowerCase();

  for (const fn of handlers) {
    const handled = await fn(context, text, lower);
    if (handled) {
      return;
    }
  }

  await context.sendActivity(helpText());
}

module.exports = {
  handleTeamsMessage,
  helpText,
};
```

---

## `scripts/smoke-ado.js`

```javascript
'use strict';

/**
 * Verifies ADO_PAT / org / project without creating work items.
 * Usage: npm run smoke
 */
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

const azdev = require('azure-devops-node-api');

async function main() {
  const pat = process.env.ADO_PAT;
  if (!pat) {
    console.error('Missing ADO_PAT in .env');
    process.exit(1);
  }

  const orgUrl = process.env.ADO_SERVER_URL?.replace(/\/$/, '');
  const org = process.env.ADO_ORG;
  const base = orgUrl || (org ? `https://dev.azure.com/${org}` : '');
  if (!base) {
    console.error('Set ADO_ORG or ADO_SERVER_URL in .env');
    process.exit(1);
  }

  const auth = azdev.getPersonalAccessTokenHandler(pat);
  const api = new azdev.WebApi(base, auth);

  await api.connect();
  console.log('ADO connection OK:', base);

  const project = process.env.ADO_PROJECT;
  if (project) {
    const wit = await api.getWorkItemTrackingApi();
    const types = await wit.getWorkItemTypes(project);
    const names = types.map((t) => t.name).join(', ');
    console.log(`Project "${project}" — work item types: ${names}`);
  } else {
    console.log('(Set ADO_PROJECT to also list work item types)');
  }
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
```

---

## Personal Cursor skill (all repositories)

To use this skill **without** committing it to each repo, copy the folder:

`init-teams-ado-bot/`  (containing `SKILL.md` and this `reference.md`)

to:

- **Windows:** `%USERPROFILE%\.cursor\skills\init-teams-ado-bot\`
- **macOS / Linux:** `~/.cursor/skills/init-teams-ado-bot/`

Do **not** place skills under `~/.cursor/skills-cursor/` (Cursor internal).

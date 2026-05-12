---
name: init-teams-ado-bot
description: >-
  Scaffolds a Node.js Microsoft Teams bot (Bot Framework) integrated with Azure DevOps Boards,
  runnable locally via Restify and on Vercel as serverless POST /api/messages; runs npm verify/smoke
  and Vercel build checks in a loop until they succeed. Use when bootstrapping this stack in a new or
  existing repo, mirroring the canonical layout, or when the user wants init + build verification until green.
  In Cursor: keep this skill under .cursor/skills/init-teams-ado-bot/ (project) or ~/.cursor/skills/init-teams-ado-bot/
  (user); read reference.md in the same folder for paste-ready file templates on greenfield.
---

# Init: Teams Bot + Azure DevOps + Vercel

## Using this skill in Cursor

- **Project skill:** copy this folder to `<repo>/.cursor/skills/init-teams-ado-bot/` so Cursor can discover it from that workspace.
- **User skill:** copy to `~/.cursor/skills/init-teams-ado-bot/` (or the path your Cursor “skills” setting uses) so it applies in any new repository.
- **Greenfield (empty or unrelated repo):** before writing files, open **`reference.md`** next to this `SKILL.md` and paste the templates (package.json, services, api, scripts, vercel.json, `public/index.html`, .gitignore, `.env.example`) into the repo root. Then run the **Verification loop** below.
- **Existing repo that already matches the layout:** follow **Bootstrap order** and **Verification loop**; edit files in place instead of overwriting from `reference.md` unless the user asked for a full scaffold.

## Goal

Produce a minimal **Node 18+** app where:

- **Local:** Restify listens on `PORT`, exposes **`POST /api/messages`** for Bot Framework.
- **Vercel:** same path via **`api/messages.js`** serverless; **`GET /api/messages`** returns a short health string.
- **Azure DevOps:** PAT-backed client creates/updates/deletes work items from natural commands.
- **Secrets:** only in `.env` / host env — **never** real IDs, passwords, or PATs in `.env.example` or committed samples.

## Target layout

```text
package.json
index.js
vercel.json
public/index.html
api/messages.js
services/botApp.js
services/botMessageHandlers.js
services/botChannelTenant.js
services/adoBoardsClient.js
services/expressLikeServerResponse.js
scripts/smoke-ado.js
scripts/verify.js
.env.example
```

## Dependencies (`package.json`)

- `botbuilder`, `dotenv`, `restify`, `azure-devops-node-api`

Scripts (optional): `start` / `dev` / `smoke` with `node --disable-warning=DEP0169` if Node logs DEP0169 from Restify stack.

## Environment (`.env.example`)

Document (empty values only):

- **ADO:** `ADO_ORG`, `ADO_PROJECT`, `ADO_PAT`, optional `ADO_SERVER_URL`, `ADO_DEFAULT_WORK_ITEM_TYPE`, optional `ADO_SEARCH_TOP` (max rows for `search` / `find`, default 20, capped at 50)
- **Bot:** `MicrosoftAppId`, `MicrosoftAppPassword`
- **Single-tenant MSAL:** one of `MicrosoftAppTenantId`, `MICROSOFT_APP_TENANT_ID`, or `AZURE_TENANT_ID` — required when the app registration is *single-tenant* or AADSTS700016 appears.
- **Local:** `PORT`

## Bootstrap order (implementation)

1. **`services/botChannelTenant.js`** — `getBotChannelAuthTenant()` reads the three tenant env names above; returns `undefined` if unset (multi-tenant / default Bot Framework tenant).

2. **`services/adoBoardsClient.js`** — `getOrgUrl`, `getConnection`, `createWorkItem`, `updateWorkItem`, `deleteWorkItem` (wrap `IWorkItemTrackingApi.deleteWorkItem`), `searchWorkItems` (WIQL `queryByWiql` + `getWorkItems` with light fields only), `toPatchOperations` (JSON Patch `add`/`replace`), `workItemWebUrl`.

3. **`services/botMessageHandlers.js`** — export `handleTeamsMessage(context)`; discrete async handlers per command (`create task`, `search`/`find`, `title`, `description`, `assign`, `unassign`, `state`/`update`, `delete`); return boolean `handled` or call `helpText()`; use case-insensitive regex prefixes so `Create task` works.

4. **`services/botApp.js`** — at top: load `.env` from repo root (`path.join(__dirname, '..', '.env')`); `TeamsAdoBot` → `onMessage` → try/catch → `handleTeamsMessage`; `getAdapter()` builds `BotFrameworkAdapter` with `channelAuthTenant` when tenant helper returns a value; **`processActivity` not `process`** on the HTTP path (avoids Zod `ResponseT` on raw `ServerResponse`). Export `processBotMessage`.

5. **`services/expressLikeServerResponse.js`** — `wrapServerResponse(raw)` implementing **`status`**, **`header`**, **`send`**, **`end`** (Bot Framework `ResponseT` and `processActivity` expect Express/Restify-like `res`).

6. **`api/messages.js` (Vercel)** — `require` **`../services/botApp` first** (runs env bootstrap); `GET`/`405`/`400` use **`writeHead`/`end`** (no `res.status()` — Node core `ServerResponse`); `POST` → optional `req.body` string JSON parse → `processBotMessage(req, wrapServerResponse(res))` with try/catch; if response already sent, do not double-write on thrown errors after `processActivity`.

7. **`index.js` (local)** — `require('./services/botApp')` before Restify; `bodyParser`; `server.post('/api/messages', ...)` → `processBotMessage(req, res)` **without** wrap (Restify response is already compatible).

8. **`vercel.json`** — `buildCommand` / `outputDirectory` `public` when the Vercel project expects that folder after `npm run build`; **`public/index.html`** (or similar) must exist in the repo so the directory is present even though `build` only runs verify. **`functions.api/messages.js`**: `includeFiles: "services/**"`; optional `maxDuration`; optional root `env.NODE_OPTIONS: --disable-warning=DEP0169`.

9. **`scripts/smoke-ado.js`** — optional; load same dotenv root pattern; verify ADO connection without creating work items.

10. **`scripts/verify.js`** — `require` core modules (`services/botApp`, `api/messages`) without starting HTTP; **`package.json`** scripts **`verify`** and **`build`** (`build` → `npm run verify`) run it. Used by the mandatory verification loop below.

## Verification loop (mandatory — until success)

The agent MUST **not** declare the init complete until checks pass. On any non-zero exit or runtime exception: **read the output**, fix the code or configuration, and **re-run from step 1** of this loop.

### 1. Install

```bash
npm install
```

- Must exit **0**. If peer/audit noise only, OK; if install fails, fix `package.json` / lockfile / registry before continuing.

### 2. Module load (`verify` / `build`)

```bash
npm run verify
```

CI often expects `npm run build`; alias it to the same check when `package.json` defines `"build": "npm run verify"`.

- Must print `verify: ok` and exit **0**.
- **`scripts/verify.js`** should load at least: `../services/botApp` and `../api/messages` (relative to `scripts/`), so missing `services/` files or syntax errors fail fast **without** binding a port.
- If `verify` is missing, add it per item 10 and the `verify` npm script.

### 3. Azure DevOps smoke (conditional)

```bash
npm run smoke
```

- Run **only if** `.env` (or documented env) defines `ADO_ORG`, `ADO_PROJECT`, and `ADO_PAT`.
- If those are absent, **skip** and state: *smoke skipped — no ADO credentials*.
- If run: must exit **0**. On failure, fix PAT/URL/client or env until it passes.

### 4. Vercel build (conditional)

When the repo is linked to Vercel **or** the user asked to verify deployment packaging:

```bash
npx vercel build --yes
```

- If the CLI reports **`project_settings_required`**, run `npx vercel link` / `vercel pull` per CLI hints **or** skip and note: *run Vercel build after linking in CI*.
- Otherwise: build must complete **successfully** (fix `vercel.json`, `api/messages.js`, `includeFiles`, missing files).
- Do **not** treat GitHub push as a substitute for a local `vercel build` when the user asked for a green build.

### 5. Local server (optional final sanity)

Only if the user needs a runtime check: start `npm start` briefly; confirm log shows listening on `PORT`. Stop the process afterwards. **Do not** block the main loop on long-running servers unless requested.

### Completion criteria

| Step | Required? |
|------|-----------|
| `npm install` | Always |
| `npm run verify` (and `npm run build` if defined) | Always |
| `npm run smoke` | When ADO env present |
| `vercel build` | When Vercel linked or user requests deploy verification |

Treat **verify** as the minimum bar for “build OK” in a greenfield scaffold without ADO secrets.

## Azure / Bot checklist (handoff to user)

- Bot **Messaging endpoint:** `https://<host>/api/messages` (HTTPS required).
- **Single-tenant:** set tenant GUID in host env (Vercel + local).
- **GitHub:** enable secret / push protection so PATs and `MicrosoftAppPassword` never enter history.

## Common failures

| Symptom | Fix |
|--------|-----|
| `ResponseT` / Zod `Response` on Vercel | Shim `res.header`; or use `processActivity` + wrap |
| `AADSTS700016` … directory `Bot Framework` | Set `MicrosoftAppTenantId` (or alias) |
| `Cannot find module …/loadEnv` | Inline dotenv bootstrap in `botApp.js`; do not rely on a tiny unmatched file in serverless bundles |
| `FUNCTION_INVOCATION_FAILED` | Missing `includeFiles` or missing `services/` file in deployment |
| No Output Directory `public` after build | Add committed **`public/`** (e.g. `index.html`) and set `vercel.json` **`outputDirectory`** `public` (or clear the dashboard override); `npm run build` must leave that folder on disk |

## When the user says "init all this code"

1. **New repo or missing stack:** use **`reference.md`** for verbatim templates, then adjust only if the user specified different names or paths.
2. **Otherwise:** create/update the files in the layout above with the behaviors described in **Bootstrap order**.
3. Execute the **Verification loop** until all required steps succeed; paste or summarize command output on failure.
4. Do **not** add optional proactive Teams send scripts or Python mirrors unless explicitly requested.
5. Keep stable module names when extending an existing project: `adoBoardsClient`, `botMessageHandlers`, `wrapServerResponse`, `verify.js`.

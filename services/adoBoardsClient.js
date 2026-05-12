'use strict';

/**
 * Azure Boards via the official client: `azure-devops-node-api` (npm).
 * @see https://github.com/microsoft/azure-devops-node-api
 */

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

/**
 * @returns {Promise<import('azure-devops-node-api').IWorkItemTrackingApi>}
 */
async function getWorkItemTrackingApi() {
  if (!cachedWit) {
    const connection = getConnection();
    cachedWit = await connection.getWorkItemTrackingApi();
  }
  return cachedWit;
}

/**
 * Map friendly field keys to JSON Patch ops for create.
 * Keys can be "System.Title" or full ref "/fields/System.Title".
 * @param {Record<string, unknown>} fields
 * @param {'add' | 'replace'} op
 * @returns {import('azure-devops-node-api/interfaces/common/VSSInterfaces').JsonPatchOperation[]}
 */
function toPatchOperations(fields, op) {
  return Object.entries(fields).map(([key, value]) => {
    const path = key.startsWith('/fields/') ? key : `/fields/${key}`;
    return { op, path, value };
  });
}

/**
 * Create a work item on Azure Boards.
 * @param {object} options
 * @param {string} options.workItemType Work item type name (e.g. Task, Bug, User Story)
 * @param {Record<string, unknown>} options.fields Field reference name → value (e.g. System.Title)
 */
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

/**
 * Update an existing work item by id.
 * @param {object} options
 * @param {number} options.id
 * @param {Record<string, unknown>} options.fields Field ref → new value (use System.State, etc.)
 * @param {string} [options.project] Defaults to ADO_PROJECT
 */
async function updateWorkItem({ id, fields, project }) {
  const wit = await getWorkItemTrackingApi();
  const proj = project ?? requireEnv('ADO_PROJECT');
  const document = toPatchOperations(fields, 'replace');
  return wit.updateWorkItem(null, document, id, proj, false, false, false);
}

/**
 * Delete (remove) a work item. Uses recycle bin unless destroy=true.
 * @param {object} options
 * @param {number} options.id
 * @param {boolean} [options.destroy] Hard-delete when true
 * @param {string} [options.project]
 */
async function deleteWorkItem({ id, destroy = false, project }) {
  const wit = await getWorkItemTrackingApi();
  const proj = project ?? requireEnv('ADO_PROJECT');
  await wit.deleteWorkItem(id, proj, destroy);
}

/**
 * @param {import('azure-devops-node-api/interfaces/WorkItemTrackingInterfaces').WorkItem} workItem
 * @returns {string | undefined}
 */
function workItemWebUrl(workItem) {
  return workItem?._links?.html?.href;
}

/** Fields returned for search results (avoid heavy HTML description). */
const SEARCH_FIELDS = [
  'System.Id',
  'System.Title',
  'System.State',
  'System.WorkItemType',
  'System.AssignedTo',
];

/**
 * @param {string} value
 */
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

/**
 * WIQL search in `ADO_PROJECT`: by id if `text` is all digits, otherwise title/description CONTAINS WORDS.
 * @param {object} options
 * @param {string} options.text Non-empty keywords or numeric id
 * @param {number} [options.top] Max rows (capped at 50)
 * @param {string} [options.project] Defaults to ADO_PROJECT
 * @returns {Promise<import('azure-devops-node-api/interfaces/WorkItemTrackingInterfaces').WorkItem[]>}
 */
async function searchWorkItems({ text, top, project }) {
  const proj = project ?? requireEnv('ADO_PROJECT');
  const wit = await getWorkItemTrackingApi();
  const limit = Math.min(Math.max(1, top ?? getSearchDefaultTop()), 50);
  const trimmed = (text || '').trim();
  if (!trimmed) {
    throw new Error('searchWorkItems requires non-empty text');
  }

  const projectLit = escapeWiqlLiteral(proj);
  /** @type {string} */
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

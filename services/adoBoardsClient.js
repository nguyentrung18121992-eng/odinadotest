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
 * @param {import('azure-devops-node-api/interfaces/WorkItemTrackingInterfaces').WorkItem} workItem
 * @returns {string | undefined}
 */
function workItemWebUrl(workItem) {
  return workItem?._links?.html?.href;
}

function resetConnectionCache() {
  cachedWit = null;
}

module.exports = {
  getWorkItemTrackingApi,
  createWorkItem,
  updateWorkItem,
  workItemWebUrl,
  resetConnectionCache,
};

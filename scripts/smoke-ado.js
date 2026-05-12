'use strict';

/**
 * Verifies ADO_PAT / org / project without creating work items.
 * Usage: npm run smoke
 */
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

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

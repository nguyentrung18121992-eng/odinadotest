'use strict';

/**
 * Quick test: proactive message to a Teams conversation using Bot Framework Connector.
 * Same pattern as Python `TeamsSender.send_to_bot` in this repo.
 *
 * Env (see .env.example):
 *   MicrosoftAppId, MicrosoftAppPassword
 * Optional: MicrosoftAppTenantId or AZURE_TENANT_ID (single-tenant bot)
 * Optional: TEAMS_SERVICE_URL (default https://smba.trafficmanager.net/teams/)
 * Optional: TEAMS_TEST_CHAT_ID, TEAMS_TEST_MESSAGE
 *
 * Usage:
 *   node scripts/test-teams-send.js
 *   node scripts/test-teams-send.js "19:xxx@thread.tacv2" "Hello"
 */

const { TeamsOutreach } = require('../services/teamsOutreach');

async function main() {
  const chatId =
    process.argv[2] ||
    process.env.TEAMS_TEST_CHAT_ID ||
    '19:abc123@thread.tacv2';
  const message =
    process.argv[3] ||
    process.env.TEAMS_TEST_MESSAGE ||
    '**pQA AI Agent** sandbox test — if you see this, bot integration works.';

  const sender = new TeamsOutreach();
  const result = await sender.sendToBot(chatId, message);
  console.log(`Success: ${result.success}, Error: ${result.error || ''}`);
  if (result.id) {
    console.log(`Activity id: ${result.id}`);
  }
  if (!result.success) {
    process.exitCode = 1;
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

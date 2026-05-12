'use strict';

/**
 * Azure AD tenant (directory) ID for Bot Framework outbound auth (MSAL).
 * Required for single-tenant app registrations; omit for multi-tenant.
 * If unset, the SDK uses `botframework.com` and AAD returns AADSTS700016 for single-tenant apps.
 *
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

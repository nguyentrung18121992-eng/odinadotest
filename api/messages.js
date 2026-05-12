'use strict';

require('../config/loadEnv');

const { processBotMessage } = require('../services/botApp');
const { wrapServerResponse } = require('../services/expressLikeServerResponse');

/**
 * Vercel Serverless Function — messaging endpoint for Azure Bot / Teams.
 * Configure Azure Bot messaging URL: https://<your-project>.vercel.app/api/messages
 *
 * Vercel passes Node http.IncomingMessage / ServerResponse (no Express .status/.send).
 * Bot Framework expects Restify/Express-style res; we wrap for POST only.
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
    // processActivity() throws after sending 4xx bodies; don't fail the invocation.
    console.error(err);
    if (!res.headersSent) {
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Internal Server Error');
    }
  }
};

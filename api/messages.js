'use strict';

require('dotenv').config();

const { processBotMessage } = require('../services/botApp');

/**
 * Vercel Serverless Function — messaging endpoint for Azure Bot / Teams.
 * Configure Azure Bot messaging URL: https://<your-project>.vercel.app/api/messages
 */
module.exports = async (req, res) => {
  if (req.method === 'GET') {
    res
      .status(200)
      .setHeader('Content-Type', 'text/plain; charset=utf-8')
      .send(
        'Teams ADO bot — POST JSON activities to this URL (Azure Bot messaging endpoint).'
      );
    return;
  }

  if (req.method !== 'POST') {
    res.setHeader('Allow', 'GET, POST');
    res.status(405).end('Method Not Allowed');
    return;
  }

  if (typeof req.body === 'string') {
    try {
      req.body = JSON.parse(req.body);
    } catch {
      res.status(400).end('Invalid JSON body');
      return;
    }
  }

  await processBotMessage(req, res);
};

'use strict';

/**
 * Express/Restify-style response wrapper for Node.js http.ServerResponse.
 * Bot Framework processActivity uses res.status(), res.send(), res.end();
 * Vercel serverless provides raw ServerResponse (no .status / .send).
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

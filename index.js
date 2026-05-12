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

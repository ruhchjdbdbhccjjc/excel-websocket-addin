/**
 * scripts/wss-proxy.js
 * Lightweight zero-dependency WSS -> WS reverse proxy.
 * Bridges secure browser contexts (Excel on the Web / HTTPS) to local Kanata (ws://127.0.0.1:8890).
 * Uses locally trusted dev certificates from ~/.office-addin-dev-certs.
 */

const https = require('https');
const net = require('net');
const fs = require('fs');
const path = require('path');

const certDir = path.join(process.env.USERPROFILE, '.office-addin-dev-certs');
const certPath = path.join(certDir, 'localhost.crt');
const keyPath = path.join(certDir, 'localhost.key');

if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
  console.error('[wss-proxy] Trusted dev certs not found in', certDir);
  console.error('[wss-proxy] Run `npx office-addin-dev-certs install` to generate them.');
  process.exit(1);
}

const options = {
  key: fs.readFileSync(keyPath),
  cert: fs.readFileSync(certPath),
};

const LISTEN_PORT = parseInt(process.env.WSS_PORT, 10) || 8891;
const TARGET_PORT = parseInt(process.env.WS_PORT, 10) || 8890;

const server = https.createServer(options, (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/plain',
    'Access-Control-Allow-Origin': '*',
  });
  res.end('Kanata Secure WSS Proxy Active\n');
});

server.on('upgrade', (req, clientSocket, head) => {
  const targetSocket = net.connect(TARGET_PORT, '127.0.0.1', () => {
    let rawReq = `${req.method} ${req.url} HTTP/${req.httpVersion}\r\n`;
    for (let i = 0; i < req.rawHeaders.length; i += 2) {
      rawReq += `${req.rawHeaders[i]}: ${req.rawHeaders[i + 1]}\r\n`;
    }
    rawReq += '\r\n';
    targetSocket.write(rawReq);
    if (head && head.length) {
      targetSocket.write(head);
    }
    clientSocket.pipe(targetSocket);
    targetSocket.pipe(clientSocket);
  });

  targetSocket.on('error', (err) => {
    console.error(`[wss-proxy] Cannot connect to Kanata at 127.0.0.1:${TARGET_PORT}:`, err.message);
    clientSocket.destroy();
  });

  clientSocket.on('error', () => {
    targetSocket.destroy();
  });
});

server.listen(LISTEN_PORT, 'localhost', () => {
  console.log(`[wss-proxy] Secure WebSocket proxy listening on wss://localhost:${LISTEN_PORT}`);
  console.log(`[wss-proxy] Forwarding traffic to Kanata at ws://127.0.0.1:${TARGET_PORT}`);
});

const http = require('http');
const https = require('https');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const CERTS = path.join(ROOT, 'network-certs');
const PORT = parseInt(process.env.PORT || (process.argv[2] || '3000'), 10);

function lanIP() {
  const nets = os.networkInterfaces();
  let fallback = null;
  for (const name of Object.keys(nets)) {
    for (const n of nets[name]) {
      if (n.family === 'IPv4' && !n.internal) {
        if (!n.address.startsWith('169.254.')) return n.address;
        if (!fallback) fallback = n.address;
      }
    }
  }
  return fallback || '127.0.0.1';
}

function sh(cmd) {
  return execSync(cmd, { stdio: 'pipe', encoding: 'utf8' }).trim();
}

function ensureCerts(ip) {
  fs.mkdirSync(CERTS, { recursive: true });
  const caKey = path.join(CERTS, 'ca.key');
  const caCert = path.join(CERTS, 'ca.crt');
  const caCnf = path.join(CERTS, 'ca.cnf');
  const serverKey = path.join(CERTS, 'server.key');
  const serverCrt = path.join(CERTS, 'server.crt');
  const extPath = path.join(CERTS, 'server.ext');

  const caConfigContent = `
[ req ]
distinguished_name = req_distinguished_name
x509_extensions = v3_ca
prompt = no

[ req_distinguished_name ]
CN = Skylight LAN Dev CA

[ v3_ca ]
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid:always,issuer
basicConstraints = critical, CA:TRUE
keyUsage = critical, digitalSignature, cRLSign, keyCertSign
`;

  if (!fs.existsSync(caCert) || !fs.existsSync(caKey) || !fs.existsSync(caCnf)) {
    fs.writeFileSync(caCnf, caConfigContent);
    sh(`openssl req -x509 -newkey rsa:2048 -nodes -keyout "${caKey}" -out "${caCert}" -days 3650 -config "${caCnf}"`);
    console.log('[network] generated Apple-compliant CA: network-certs/ca.crt');
  }

  // Leaf cert: 365 days validity (Apple macOS strictly requires <= 398 days for TLS)
  const ext = `
basicConstraints = CA:FALSE
keyUsage = critical, digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = DNS:localhost, IP:127.0.0.1, IP:${ip}
`;

  const existingExt = fs.existsSync(extPath) ? fs.readFileSync(extPath, 'utf8') : '';
  if (!fs.existsSync(serverCrt) || !fs.existsSync(serverKey) || existingExt !== ext) {
    fs.writeFileSync(extPath, ext);
    sh(`openssl req -new -newkey rsa:2048 -nodes -keyout "${serverKey}" -out "${path.join(CERTS, 'server.csr')}" -subj "/CN=${ip}"`);
    sh(`openssl x509 -req -in "${path.join(CERTS, 'server.csr')}" -CA "${caCert}" -CAkey "${caKey}" -CAcreateserial -out "${serverCrt}" -days 365 -extfile "${extPath}"`);
    console.log('[network] generated Apple-compliant server certificate (365d) for localhost and ' + ip);
  }
}

function writeNetworkManifest(ip) {
  const xml = fs.readFileSync(path.join(ROOT, 'manifest.xml'), 'utf8');
  const out = xml.split('https://localhost:3000').join(`https://${ip}:${PORT}`);
  fs.writeFileSync(path.join(ROOT, 'manifest.network.xml'), out);
}

function contentType(name) {
  const map = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.png': 'image/png',
    '.xml': 'application/xml; charset=utf-8',
    '.crt': 'application/x-x509-ca-cert',
    '.key': 'application/octet-stream',
  };
  return map[path.extname(name)] || 'application/octet-stream';
}

function route(url) {
  const routes = {
    '/app.html': () => path.join(ROOT, 'src/app.html'),
    '/app.css': () => path.join(ROOT, 'src/app.css'),
    '/timezones.js': () => path.join(ROOT, 'src/timezones.js'),
    '/ics.js': () => path.join(ROOT, 'src/ics.js'),
    '/app.js': () => path.join(ROOT, 'src/app.js'),
    '/manifest.xml': () => path.join(ROOT, 'manifest.xml'),
    '/manifest.network.xml': () => path.join(ROOT, 'manifest.network.xml'),
    '/ca.crt': () => path.join(CERTS, 'ca.crt'),
    '/server.crt': () => path.join(CERTS, 'server.crt'),
  };
  if (routes[url]) return { file: routes[url](), name: url };
  if (url.startsWith('/assets/')) {
    return { file: path.join(ROOT, url), name: url };
  }
  if (url === '/' || url === '/index.html') {
    return { file: path.join(ROOT, 'src/app.html'), name: '/app.html', redirect: true };
  }
  return null;
}

const ip = lanIP();
ensureCerts(ip);
writeNetworkManifest(ip);

const server = https.createServer(
  {
    key: fs.readFileSync(path.join(CERTS, 'server.key')),
    cert: fs.readFileSync(path.join(CERTS, 'server.crt')),
  },
  (req, res) => {
    // Enable CORS for Office on the Web / iframes
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', '*');

    if (req.method === 'OPTIONS') {
      res.writeHead(204);
      return res.end();
    }

    const r = route(req.url.split('?')[0]);
    if (!r) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      return res.end('Not found');
    }
    fs.readFile(r.file, (err, data) => {
      if (err) {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        return res.end('Not found: ' + r.name);
      }
      res.writeHead(200, {
        'Content-Type': contentType(r.name),
        'Cache-Control': 'no-store',
        'X-Content-Type-Options': 'nosniff',
      });
      res.end(data);
    });
  }
);

server.listen(PORT, '0.0.0.0', () => {
  const ca = `https://${ip}:${PORT}/ca.crt`;
  const man = `https://${ip}:${PORT}/manifest.network.xml`;
  const app = `https://${ip}:${PORT}/app.html`;
  console.log('');
  console.log('[network] add-in server ready on ALL network interfaces');
  console.log('');
  console.log('  App page:      ' + app);
  console.log('  Manifest:      ' + man);
  console.log('  Trust anchor:  ' + ca);
  console.log('');
  console.log('REMOTE MACHINE — trust the CA once:');
  console.log('  Windows (Run as Administrator):');
  console.log('    curl.exe -k -o %TEMP%\\ca.crt ' + ca);
  console.log('    certutil -addstore -f Root %TEMP%\\ca.crt');
  console.log('  macOS:');
  console.log('    curl -k -o /tmp/ca.crt ' + ca);
  console.log('    security add-trusted-cert -d -r trustRoot -p ssl -k ~/Library/Keychains/login.keychain-db /tmp/ca.crt');
  console.log('    sudo security add-trusted-cert -d -r trustRoot -p ssl -k /Library/Keychains/System.keychain /tmp/ca.crt');
  console.log('');
  console.log('Then sideload in Outlook on that machine:');
  console.log('  Get Add-ins > My add-ins > Add a custom add-in > Add from URL');
  console.log('  Paste: ' + man);
  console.log('');
  console.log('Note: If Outlook is already open on macOS, restart it (Cmd+Q) to reload certificates.');
  console.log('');
  console.log('This machine must allow inbound TCP ' + PORT + ' through its firewall.');
});
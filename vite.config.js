import { defineConfig } from 'vite';
import fs from 'fs';
import path from 'path';
import os from 'os';
import https from 'https';
import { execFile } from 'child_process';

const STYLE_FILE = path.resolve(process.cwd(), 'my_personal_style.json');
const MAX_BODY_BYTES = 65536; // 64 KB hard cap for POST /api/style

// In-memory canonical store — no read-before-write race on concurrent POSTs
let styleDataCache = null;
function getStyleData() {
  if (!styleDataCache) {
    styleDataCache = fs.existsSync(STYLE_FILE)
      ? JSON.parse(fs.readFileSync(STYLE_FILE, 'utf8'))
      : { examples: [] };
  }
  return styleDataCache;
}
function persistStyleData() {
  fs.writeFileSync(STYLE_FILE, JSON.stringify(styleDataCache, null, 2), 'utf8');
}

async function getHttpsOptions() {
  try {
    const certsDir = path.join(os.homedir(), '.office-addin-dev-certs');
    const keyFile = path.join(certsDir, 'localhost.key');
    const certFile = path.join(certsDir, 'localhost.crt');

    if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
      return {
        key: fs.readFileSync(keyFile),
        cert: fs.readFileSync(certFile),
      };
    }
  } catch (e) {
    console.warn("Could not load dev certs. Run 'npx office-addin-dev-certs install' first.");
  }
  return false;
}

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

// Vite plugin: auto-scan past-works/ and generate index.json on every change
function pastWorksPlugin() {
  const worksDir = path.resolve(process.cwd(), 'past-works');

  function generateIndex() {
    if (!fs.existsSync(worksDir)) {
      fs.mkdirSync(worksDir, { recursive: true });
    }
    const files = fs.readdirSync(worksDir)
      .filter(f => /\.(docx|pdf|txt|md)$/i.test(f))
      .map(f => {
        const stat = fs.statSync(path.join(worksDir, f));
        return { name: f, size: stat.size, modified: stat.mtime.toISOString() };
      });
    const indexPath = path.join(worksDir, 'index.json');
    fs.writeFileSync(indexPath, JSON.stringify({ files }, null, 2), 'utf8');
    if (files.length) console.log(`[past-works] ${files.length} file(s) indexed.`);
  }

  return {
    name: 'past-works',
    buildStart() { generateIndex(); },
    configureServer(server) {
      generateIndex();
      server.watcher.add(worksDir);
      server.watcher.on('all', (event, filePath) => {
        if (filePath.startsWith(worksDir) && !filePath.endsWith('index.json')) {
          generateIndex();
          // Trigger HMR so the add-in re-fetches index.json
          server.ws.send({ type: 'full-reload' });
        }
      });
    }
  };
}

// Vite plugin: local style bank REST endpoints + DOI validation proxy
function styleBankPlugin() {
  return {
    name: 'style-bank',
    configureServer(server) {

      server.middlewares.use('/api/storage-sync', (req, res) => {
        if (req.method === 'OPTIONS') { res.writeHead(204, CORS_HEADERS); res.end(); return; }

        if (req.method !== 'POST') {
          res.writeHead(405, CORS_HEADERS);
          res.end();
          return;
        }

        let body = '';
        let bodyBytes = 0;
        req.on('data', chunk => {
          bodyBytes += chunk.length;
          if (bodyBytes > MAX_BODY_BYTES) {
            req.destroy();
            if (!res.headersSent) {
              res.writeHead(413, CORS_HEADERS);
              res.end(JSON.stringify({ error: 'payload too large' }));
            }
            return;
          }
          body += chunk;
        });
        req.on('end', () => {
          try {
            const { source } = JSON.parse(body || '{}');
            const trimmedSource = String(source || '').trim();
            if (!trimmedSource) {
              res.writeHead(400, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ error: 'source is required' }));
              return;
            }

            execFile(
              process.execPath,
              [path.resolve(process.cwd(), 'scripts', 'sync-storage-to-local.mjs'), trimmedSource, '--json'],
              { cwd: process.cwd(), timeout: 120000, maxBuffer: 1024 * 1024 },
              (error, stdout, stderr) => {
                if (error) {
                  res.writeHead(500, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
                  res.end(JSON.stringify({ ok: false, error: stderr || stdout || error.message }));
                  return;
                }
                const parsed = JSON.parse(stdout || '{}');
                res.writeHead(200, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: true, data: parsed }));
              }
            );
          } catch (e) {
            res.writeHead(500, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ ok: false, error: e.message }));
          }
        });
      });

      // ── /api/validate-doi?doi=10.xxxx/yyy  (server-side HEAD, no CORS) ──
      server.middlewares.use('/api/validate-doi', (req, res) => {
        if (req.method === 'OPTIONS') { res.writeHead(204, CORS_HEADERS); res.end(); return; }
        const url = new URL(req.url, 'https://localhost');
        const doi = url.searchParams.get('doi');
        if (!doi || !/^10\.\d{4,}\/\S+/.test(doi)) {
          res.writeHead(400, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: 'invalid doi' }));
          return;
        }
        const doiUrl = `https://doi.org/${doi}`;
        const doiReq = https.request(doiUrl, { method: 'HEAD', timeout: 5000 }, (doiRes) => {
          const ok = doiRes.statusCode < 400 || doiRes.statusCode === 301 || doiRes.statusCode === 302;
          res.writeHead(200, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok }));
        });
        doiReq.on('timeout', () => { doiReq.destroy(); res.writeHead(200, CORS_HEADERS); res.end(JSON.stringify({ ok: null })); });
        doiReq.on('error', () => { if (!res.headersSent) { res.writeHead(200, CORS_HEADERS); res.end(JSON.stringify({ ok: null })); } });
        doiReq.end();
      });

      // ── /api/style  GET + POST ───────────────────────────────────────────
      server.middlewares.use('/api/style', (req, res) => {
        if (req.method === 'OPTIONS') { res.writeHead(204, CORS_HEADERS); res.end(); return; }

        if (req.method === 'GET') {
          res.writeHead(200, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
          res.end(JSON.stringify(getStyleData()));
          return;
        }

        if (req.method === 'POST') {
          let body = '';
          let bodyBytes = 0;
          req.on('data', chunk => {
            bodyBytes += chunk.length;
            if (bodyBytes > MAX_BODY_BYTES) {
              req.destroy();
              if (!res.headersSent) { res.writeHead(413, CORS_HEADERS); res.end(JSON.stringify({ error: 'payload too large' })); }
              return;
            }
            body += chunk;
          });
          req.on('end', () => {
            try {
              const { text } = JSON.parse(body);
              if (!text || typeof text !== 'string' || text.trim().length < 10) {
                res.writeHead(400, CORS_HEADERS);
                res.end(JSON.stringify({ error: 'text must be at least 10 chars' }));
                return;
              }
              const sanitized = text.trim().slice(0, 2000);
              const data = getStyleData();
              data.examples.push({ savedAt: new Date().toISOString(), text: sanitized });
              persistStyleData();
              res.writeHead(200, { ...CORS_HEADERS, 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ ok: true, total: data.examples.length }));
            } catch (e) {
              res.writeHead(500, CORS_HEADERS);
              res.end(JSON.stringify({ error: e.message }));
            }
          });
          return;
        }

        res.writeHead(405, CORS_HEADERS);
        res.end();
      });
    }
  };
}

export default async () => defineConfig({
  plugins: [pastWorksPlugin(), styleBankPlugin()],
  server: {
    port: 3000,
    strictPort: true,
    https: await getHttpsOptions()
  }
});

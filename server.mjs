import http from 'node:http';
import { readFile, stat, mkdir, writeFile } from 'node:fs/promises';
import { createReadStream, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, 'public');
const dataDir = path.join(__dirname, 'data');
const projectsFile = path.join(dataDir, 'projects.json');
const port = Number(process.env.PORT || 8018);

const contentTypes = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8'
};

const OVERPASS_ENDPOINTS = [
  'https://overpass-api.de/api/interpreter',
  'https://overpass.kumi.systems/api/interpreter',
  'https://overpass.openstreetmap.ru/api/interpreter'
];

const execFileAsync = promisify(execFile);

async function ensureData() {
  await mkdir(dataDir, { recursive: true });
  if (!existsSync(projectsFile)) {
    await writeFile(projectsFile, '[]\n');
  }
}

function sendJson(res, status, payload) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS'
  });
  res.end(JSON.stringify(payload));
}

function notFound(res) {
  sendJson(res, 404, { error: 'Not found' });
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => {
      body += chunk;
      if (body.length > 5 * 1024 * 1024) {
        reject(new Error('Body too large'));
        req.destroy();
      }
    });
    req.on('end', () => {
      if (!body) return resolve({});
      try {
        resolve(JSON.parse(body));
      } catch (err) {
        reject(err);
      }
    });
    req.on('error', reject);
  });
}

function parseForm(body) {
  const params = new URLSearchParams(body);
  return Object.fromEntries(params.entries());
}

async function postOverpass(query) {
  let lastError = 'unknown_error';
  for (const endpoint of OVERPASS_ENDPOINTS) {
    try {
      const { stdout } = await execFileAsync('curl', [
        '--silent',
        '--show-error',
        '--max-time', '15',
        endpoint,
        '-H', 'content-type: application/x-www-form-urlencoded',
        '--data-urlencode', `data=${query}`
      ], { timeout: 20000, maxBuffer: 10 * 1024 * 1024 });
      const data = JSON.parse(stdout);
      return { ok: true, endpoint, data };
    } catch (err) {
      lastError = err?.message || 'fetch_failed';
    }
  }
  return { ok: false, error: lastError };
}

async function fetchElevations(lats, lngs) {
  const out = [];
  const BATCH = 100;
  for (let i = 0; i < lats.length; i += BATCH) {
    const latBatch = lats.slice(i, i + BATCH).map(v => Number(v).toFixed(6)).join(',');
    const lngBatch = lngs.slice(i, i + BATCH).map(v => Number(v).toFixed(6)).join(',');
    const openMeteoUrl = `https://api.open-meteo.com/v1/elevation?latitude=${latBatch}&longitude=${lngBatch}`;
    const euPts = [];
    for (let j = i; j < Math.min(i + BATCH, lats.length); j++) euPts.push([Number(lats[j]), Number(lngs[j])]);
    const euUrl = 'https://www.elevation-api.eu/v1/elevation?pts=' + encodeURIComponent(JSON.stringify(euPts));
    const openElevationUrl = 'https://api.open-elevation.com/api/v1/lookup?locations=' + encodeURIComponent(euPts.map(([a,b]) => a.toFixed(6) + ',' + b.toFixed(6)).join('|'));

    let done = false;

    try {
      const resp = await fetch(openMeteoUrl);
      if (resp.ok) {
        const data = await resp.json();
        if (Array.isArray(data.elevation)) {
          out.push(...data.elevation.map(Number));
          done = true;
        }
      }
    } catch {}

    if (!done) {
      try {
        const resp = await fetch(euUrl);
        const data = await resp.json();
        if (resp.ok && Array.isArray(data.elevations)) {
          out.push(...data.elevations.map(Number));
          done = true;
        }
      } catch {}
    }

    if (!done) {
      try {
        const resp = await fetch(openElevationUrl);
        const data = await resp.json();
        if (resp.ok && Array.isArray(data.results)) {
          out.push(...data.results.map(row => Number(row.elevation)));
          done = true;
        }
      } catch {}
    }

    if (!done) throw new Error('elevation_all_providers_failed');
  }
  return out;
}

async function loadProjects() {
  await ensureData();
  return JSON.parse(await readFile(projectsFile, 'utf8'));
}

async function saveProjects(projects) {
  await ensureData();
  await writeFile(projectsFile, JSON.stringify(projects, null, 2) + '\n');
}

async function handleApi(req, res, url) {
  if (req.method === 'OPTIONS') return sendJson(res, 204, {});

  if (url.pathname === '/api/overpass' && req.method === 'POST') {
    let raw = '';
    await new Promise((resolve, reject) => {
      req.on('data', chunk => { raw += chunk; });
      req.on('end', resolve);
      req.on('error', reject);
    });
    const form = parseForm(raw || '');
    const query = form.data;
    if (!query) return sendJson(res, 400, { error: 'Missing or invalid query' });
    const result = await postOverpass(query);
    if (!result.ok) return sendJson(res, 502, { error: result.error || 'Overpass failed' });
    return sendJson(res, 200, result.data);
  }

  if (url.pathname === '/api/check-subscription' && req.method === 'GET') {
    return sendJson(res, 200, {
      authenticated: false,
      plan: 'local',
      canUsePdfExport: true,
      canSaveCalculations: true,
      canShareProjects: false,
      source: 'local-shim'
    });
  }

  if (url.pathname === '/api/list-projects' && req.method === 'GET') {
    const projects = await loadProjects();
    return sendJson(res, 200, { projects, source: 'local-shim' });
  }

  if (url.pathname === '/api/save-project' && req.method === 'POST') {
    const body = await parseBody(req);
    const projects = await loadProjects();
    const id = body.projectId || `local_${Date.now()}`;
    const row = {
      id,
      name: body.name || 'Untitled Project',
      state: body.state || {},
      updatedAt: new Date().toISOString()
    };
    const idx = projects.findIndex(p => p.id === id);
    if (idx >= 0) projects[idx] = row; else projects.unshift(row);
    await saveProjects(projects);
    return sendJson(res, 200, { ok: true, projectId: id, source: 'local-shim' });
  }

  if (url.pathname === '/api/load-project' && req.method === 'GET') {
    const id = url.searchParams.get('id');
    const projects = await loadProjects();
    const project = projects.find(p => p.id === id);
    if (!project) return sendJson(res, 404, { error: 'Project not found' });
    return sendJson(res, 200, { project, source: 'local-shim' });
  }

  if (url.pathname === '/api/delete-project' && req.method === 'POST') {
    const body = await parseBody(req);
    const projects = await loadProjects();
    const filtered = projects.filter(p => p.id !== body.projectId);
    await saveProjects(filtered);
    return sendJson(res, 200, { ok: true, source: 'local-shim' });
  }

  if (url.pathname === '/api/buildings-hybrid' && req.method === 'POST') {
    const body = await parseBody(req);
    if (!body.query) return sendJson(res, 400, { error: 'Missing or invalid query' });
    const result = await postOverpass(body.query);
    if (!result.ok) return sendJson(res, 502, { error: result.error || 'Overpass failed' });
    return sendJson(res, 200, {
      elements: Array.isArray(result.data?.elements) ? result.data.elements : [],
      linz: {
        added: 0,
        skipped: true,
        reason: 'local-build-no-linz-dataset'
      },
      source: 'local-overpass'
    });
  }

  if (url.pathname === '/api/elevation-batch' && req.method === 'POST') {
    const body = await parseBody(req);
    if (!Array.isArray(body.lats) || !Array.isArray(body.lngs) || body.lats.length !== body.lngs.length) {
      return sendJson(res, 400, { error: 'lats and lngs must be equal-length arrays' });
    }
    try {
      const lats = body.lats.map(Number);
      const lngs = body.lngs.map(Number);
      const elevations = await fetchElevations(lats, lngs);
      return sendJson(res, 200, {
        elevations,
        source: 'open-meteo',
        refined: false
      });
    } catch (err) {
      return sendJson(res, 502, { error: err?.message || 'Elevation fetch failed' });
    }
  }

  if (url.pathname.startsWith('/api/')) {
    return sendJson(res, 501, { error: 'Route not implemented locally yet', route: url.pathname });
  }

  return notFound(res);
}

async function handleStatic(req, res, url) {
  let pathname = decodeURIComponent(url.pathname);
  if (pathname === '/') pathname = '/public/index.html';
  const safePath = path.normalize(path.join(__dirname, pathname));
  if (!safePath.startsWith(__dirname)) {
    res.writeHead(403);
    return res.end('Forbidden');
  }

  try {
    let candidatePath = safePath;
    try {
      await stat(candidatePath);
    } catch {
      const publicCandidate = path.normalize(path.join(publicDir, pathname.replace(/^\//, '')));
      if (publicCandidate.startsWith(publicDir) && existsSync(publicCandidate)) {
        candidatePath = publicCandidate;
      } else {
        throw new Error('not_found');
      }
    }

    const st = await stat(candidatePath);
    const finalPath = st.isDirectory() ? path.join(candidatePath, 'index.html') : candidatePath;
    const ext = path.extname(finalPath).toLowerCase();
    res.writeHead(200, { 'Content-Type': contentTypes[ext] || 'application/octet-stream' });
    createReadStream(finalPath).pipe(res);
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Not found');
  }
}

await ensureData();

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    if (url.pathname.startsWith('/api/')) return await handleApi(req, res, url);
    return await handleStatic(req, res, url);
  } catch (err) {
    return sendJson(res, 500, { error: err.message || 'Server error' });
  }
});

server.listen(port, () => {
  console.log(`structuralwind-local listening on http://127.0.0.1:${port}`);
});

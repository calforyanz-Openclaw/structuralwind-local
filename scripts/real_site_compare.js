#!/usr/bin/env node
const fs = require('fs');
const http = require('http');

const CDP_HTTP = 'http://127.0.0.1:18800';
const LIVE_URL = 'https://structuralwind.com/';
const LOCAL_URL = 'http://127.0.0.1:8018/';
const OUT_JSON = '/Users/jasonjia-claw/.openclaw/workspace/structuralwind-local/notes/real-site-compare.json';
const OUT_MD = '/Users/jasonjia-claw/.openclaw/workspace/structuralwind-local/notes/real-site-compare.md';

const GEOMETRY = {
  width: 20,
  depth: 15,
  height: 6,
  pitch: 15,
  roofType: 'gable',
  parapet: 0,
  overhang: 0.5,
  windSpeed: 45,
  svcVr: 32,
  terrainCat: 3,
  importance: 2,
  life: 50,
  loadCase: 'A',
  windAngle: 315,
  mapBuildingAngle: 0,
  openWW: 5,
  openLW: 5,
  openSW: 5,
  Kp: '1.0'
};

const SITES = [
  { name: 'Auckland CBD', lat: -36.8485, lng: 174.7633 },
  { name: 'Wellington CBD', lat: -41.2865, lng: 174.7762 },
  { name: 'Christchurch CBD', lat: -43.5321, lng: 172.6362 }
];

function httpJson(method, path) {
  return new Promise((resolve, reject) => {
    const req = http.request(CDP_HTTP + path, { method }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error(`Invalid JSON from ${path}: ${data.slice(0, 300)}`)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

async function openTarget(url) {
  return httpJson('PUT', '/json/new?' + encodeURIComponent(url));
}

async function evalExpr(wsUrl, expression, timeoutMs = 180000) {
  const ws = new WebSocket(wsUrl);
  await new Promise((resolve, reject) => {
    ws.onopen = resolve;
    ws.onerror = reject;
  });
  const result = await new Promise((resolve, reject) => {
    const id = 1;
    const timer = setTimeout(() => reject(new Error('timeout')), timeoutMs);
    ws.onmessage = ev => {
      const msg = JSON.parse(ev.data.toString());
      if (msg.id === id) {
        clearTimeout(timer);
        resolve(msg);
      }
    };
    ws.send(JSON.stringify({
      id,
      method: 'Runtime.evaluate',
      params: { expression, returnByValue: true, awaitPromise: true }
    }));
  });
  ws.close();
  if (result.result?.exceptionDetails) {
    throw new Error(JSON.stringify(result.result.exceptionDetails).slice(0, 400));
  }
  return result.result?.result?.value;
}

async function waitForReady(wsUrl, timeoutMs = 30000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const ready = await evalExpr(wsUrl, `(() => ({
        ready: document.readyState,
        hasS: typeof S !== 'undefined',
        hasCalc: typeof calc !== 'undefined',
        hasAuto: typeof autoDetectAllMultipliers !== 'undefined'
      }))()`, 10000);
      if (ready.ready === 'complete' && ready.hasS && ready.hasCalc && ready.hasAuto) return;
    } catch (_) {}
    await new Promise(r => setTimeout(r, 500));
  }
  throw new Error('Timed out waiting for app readiness');
}

function buildExpression(geometry, site) {
  return `
(async () => {
  const geometry = ${JSON.stringify(geometry)};
  const site = ${JSON.stringify(site)};
  Object.assign(S, geometry);
  S.lat = site.lat;
  S.lng = site.lng;
  S.address = site.name;
  S.detectedBuildingsPerSector = null;
  S.detectedBuildingsList = null;
  S.terrainRecalcCtx = null;
  S.detectedNearBuildings = null;
  S.detectedSectorWater = null;
  S.detectedSectorOpen = null;
  S.detectedElevations = null;
  S.detectedSiteElev = null;
  S.detectedSampleDistances = null;
  S.terrainDataSource = 'unknown';
  try {
    const rr = autoDetectRegion(S.lat, S.lng);
    if (rr) applyRegion(rr);
  } catch (_) {}
  await autoDetectAllMultipliers({});
  return {
    site: site.name,
    region: S.region,
    terrainDataSource: S.terrainDataSource,
    detectedSiteElev: S.detectedSiteElev,
    tcDir: Array.isArray(S.TC_dir) ? S.TC_dir.slice() : null,
    result: {
      Vsit: S.R?.Vsit ?? null,
      qz: S.R?.qz ?? null,
      Mz: S.R?.Mz ?? null,
      Ms: S.R?.Ms ?? null,
      Mt: S.R?.Mt ?? null,
      Md: S.R?.Md ?? null,
      tc: S.R?.tc ?? null,
      windward_p: S.R?.faces?.windward?.p ?? null,
      leeward_p: S.R?.faces?.leeward?.p ?? null,
      roof_ww_p: S.R?.faces?.roof_ww?.p ?? null,
      roof_lw_p: S.R?.faces?.roof_lw?.p ?? null
    }
  };
})()`;
}

function numDiff(a, b) {
  if (typeof a !== 'number' || typeof b !== 'number') return null;
  return Number((a - b).toFixed(12));
}

function pctDiff(a, b) {
  if (typeof a !== 'number' || typeof b !== 'number' || a === 0) return null;
  return Number((((b - a) / a) * 100).toFixed(3));
}

function firstThree(arr) {
  return Array.isArray(arr) ? arr.slice(0, 3).map(v => Number(v.toFixed ? v.toFixed(2) : v)) : null;
}

function summarize(liveRows, localRows) {
  const rows = [];
  for (const live of liveRows) {
    const local = localRows.find(x => x.site === live.site);
    if (!local) continue;
    rows.push({
      site: live.site,
      region: `${live.region}/${local.region}`,
      source: `${live.terrainDataSource}/${local.terrainDataSource}`,
      elev: `${live.detectedSiteElev}/${local.detectedSiteElev}`,
      tcDirHead: `${JSON.stringify(firstThree(live.tcDir))} / ${JSON.stringify(firstThree(local.tcDir))}`,
      liveVsit: live.result.Vsit,
      localVsit: local.result.Vsit,
      vsitDiff: numDiff(live.result.Vsit, local.result.Vsit),
      vsitPct: pctDiff(live.result.Vsit, local.result.Vsit),
      liveMt: live.result.Mt,
      localMt: local.result.Mt,
      mtDiff: numDiff(live.result.Mt, local.result.Mt),
      liveMs: live.result.Ms,
      localMs: local.result.Ms,
      msDiff: numDiff(live.result.Ms, local.result.Ms),
      liveQz: live.result.qz,
      localQz: local.result.qz,
      qzDiff: numDiff(live.result.qz, local.result.qz),
      roofWWDiff: numDiff(live.result.roof_ww_p, local.result.roof_ww_p),
      windwardDiff: numDiff(live.result.windward_p, local.result.windward_p)
    });
  }
  return rows;
}

function writeReports(raw, summary) {
  fs.writeFileSync(OUT_JSON, JSON.stringify({ geometry: GEOMETRY, sites: SITES, raw, summary }, null, 2) + '\n');
  const lines = [];
  lines.push('# Real Site Compare Report');
  lines.push('');
  lines.push('Fixed geometry for all real-site comparisons:');
  lines.push(`- width ${GEOMETRY.width} m, depth ${GEOMETRY.depth} m, height ${GEOMETRY.height} m, pitch ${GEOMETRY.pitch}°, roof ${GEOMETRY.roofType}`);
  lines.push(`- windSpeed ${GEOMETRY.windSpeed} m/s, svcVr ${GEOMETRY.svcVr} m/s, loadCase ${GEOMETRY.loadCase}, windAngle ${GEOMETRY.windAngle}°`);
  lines.push('');
  lines.push('| Site | Region (live/local) | Source (live/local) | Elev (live/local) | TC_dir head (live/local) | Vsit live | Vsit local | ΔVsit | ΔVsit % | Mt live | Mt local | ΔMt | Ms live | Ms local | ΔMs | qz live | qz local | Δqz | Δroof WW p | Δwindward p |');
  lines.push('|---|---|---|---:|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|');
  for (const r of summary) {
    lines.push(`| ${r.site} | ${r.region} | ${r.source} | ${r.elev} | ${r.tcDirHead} | ${r.liveVsit.toFixed(6)} | ${r.localVsit.toFixed(6)} | ${r.vsitDiff.toFixed(6)} | ${r.vsitPct.toFixed(3)}% | ${r.liveMt.toFixed(6)} | ${r.localMt.toFixed(6)} | ${r.mtDiff.toFixed(6)} | ${r.liveMs.toFixed(6)} | ${r.localMs.toFixed(6)} | ${r.msDiff.toFixed(6)} | ${r.liveQz.toFixed(6)} | ${r.localQz.toFixed(6)} | ${r.qzDiff.toFixed(6)} | ${r.roofWWDiff.toFixed(6)} | ${r.windwardDiff.toFixed(6)} |`);
  }
  lines.push('');
  lines.push('## Notes');
  lines.push('- Differences here are map/data-source path differences, not formula-core differences.');
  lines.push('- `terrainDataSource` may differ even when outputs match or nearly match because one side can reuse cached terrain/elevation data.');
  fs.writeFileSync(OUT_MD, lines.join('\n') + '\n');
}

(async () => {
  const liveTarget = await openTarget(LIVE_URL);
  const localTarget = await openTarget(LOCAL_URL);
  await waitForReady(liveTarget.webSocketDebuggerUrl);
  await waitForReady(localTarget.webSocketDebuggerUrl);
  const live = [];
  const local = [];
  for (const site of SITES) {
    console.log('running', site.name, 'live');
    live.push(await evalExpr(liveTarget.webSocketDebuggerUrl, buildExpression(GEOMETRY, site), 120000));
    console.log('running', site.name, 'local');
    local.push(await evalExpr(localTarget.webSocketDebuggerUrl, buildExpression(GEOMETRY, site), 120000));
  }
  const summary = summarize(live, local);
  writeReports({ live, local }, summary);
  console.log(JSON.stringify({ outJson: OUT_JSON, outMd: OUT_MD, rows: summary.length }, null, 2));
})();

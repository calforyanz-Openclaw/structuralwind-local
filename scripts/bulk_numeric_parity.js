#!/usr/bin/env node
const fs = require('fs');
const http = require('http');

const CDP_HTTP = 'http://127.0.0.1:18800';
const LIVE_URL = 'https://structuralwind.com/';
const LOCAL_URL = 'http://127.0.0.1:8018/';
const OUT_JSON = '/Users/jasonjia-claw/.openclaw/workspace/structuralwind-local/notes/bulk-numeric-parity.json';
const OUT_MD = '/Users/jasonjia-claw/.openclaw/workspace/structuralwind-local/notes/bulk-numeric-parity.md';

function httpJson(method, path) {
  return new Promise((resolve, reject) => {
    const req = http.request(CDP_HTTP + path, { method }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error(`Invalid JSON from ${path}: ${data.slice(0,200)}`)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

async function openTarget(url) {
  return httpJson('PUT', '/json/new?' + encodeURIComponent(url));
}

async function evalExpr(wsUrl, expression, timeoutMs = 20000) {
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
    throw new Error(JSON.stringify(result.result.exceptionDetails).slice(0, 500));
  }
  return result.result?.result?.value;
}

async function waitForReady(wsUrl, timeoutMs = 30000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const ready = await evalExpr(wsUrl, `(() => ({
        href: location.href,
        ready: document.readyState,
        hasS: typeof S !== 'undefined',
        hasCalc: typeof calc !== 'undefined'
      }))()` , 10000);
      if (ready && ready.ready === 'complete' && ready.hasS && ready.hasCalc) return ready;
    } catch (_) {}
    await new Promise(r => setTimeout(r, 500));
  }
  throw new Error('Timed out waiting for app readiness');
}

function mulberry32(a) {
  return function() {
    let t = a += 0x6D2B79F5;
    t = Math.imul(t ^ t >>> 15, t | 1);
    t ^= t + Math.imul(t ^ t >>> 7, t | 61);
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  };
}

function randRange(rand, min, max, digits = 3) {
  const val = min + (max - min) * rand();
  return Number(val.toFixed(digits));
}

function choice(rand, arr) {
  return arr[Math.floor(rand() * arr.length)];
}

function generateCases(n = 100, seed = 20260429) {
  const rand = mulberry32(seed);
  const roofTypes = ['gable', 'hip', 'flat', 'monoslope'];
  const loadCases = ['A', 'B', 'C', 'D'];
  const regions = ['NZ1', 'NZ2', 'NZ3', 'NZ4'];
  const terrains = [1, 1.5, 2, 2.5, 3, 4];
  const cases = [];
  for (let i = 0; i < n; i++) {
    const terrainCat = choice(rand, terrains);
    const cfg = {
      name: `case_${String(i + 1).padStart(3, '0')}`,
      width: randRange(rand, 8, 60, 3),
      depth: randRange(rand, 8, 80, 3),
      height: randRange(rand, 3, 25, 3),
      pitch: randRange(rand, 0, 35, 3),
      roofType: choice(rand, roofTypes),
      parapet: randRange(rand, 0, 2.0, 3),
      overhang: randRange(rand, 0, 1.2, 3),
      windSpeed: randRange(rand, 30, 70, 3),
      svcVr: randRange(rand, 20, 45, 3),
      terrainCat,
      importance: choice(rand, [1, 2, 3, 4]),
      life: choice(rand, [5, 10, 25, 50, 100]),
      loadCase: choice(rand, loadCases),
      windAngle: choice(rand, [0, 15, 30, 45, 60, 75, 90, 120, 135, 150, 180, 225, 270, 315, 345]),
      mapBuildingAngle: randRange(rand, 0, 359, 3),
      openWW: randRange(rand, 0, 50, 3),
      openLW: randRange(rand, 0, 50, 3),
      openSW: randRange(rand, 0, 50, 3),
      region: choice(rand, regions),
      ari: choice(rand, [25, 50, 100, 250, 500, 1000]),
      Kp: choice(rand, ['1.0', '0.9', '0.8', '0.7']),
      Md: Array.from({ length: 8 }, () => randRange(rand, 0.85, 1.0, 4)),
      TC_dir: Array.from({ length: 8 }, () => choice(rand, terrains)),
      Ms: Array.from({ length: 8 }, () => randRange(rand, 0.85, 1.2, 4)),
      Mt: Array.from({ length: 8 }, () => randRange(rand, 0.9, 1.25, 4)),
      Mt_hill: Array.from({ length: 8 }, () => 1),
      Mlee: Array.from({ length: 8 }, () => randRange(rand, 0.9, 1.15, 4))
    };
    if (cfg.roofType === 'flat') cfg.pitch = 0;
    cases.push(cfg);
  }
  return cases;
}

function buildExpression(cases) {
  return `(() => {
    const cases = ${JSON.stringify(cases)};
    const get = (obj, path, fallback = null) => {
      try {
        return path.reduce((acc, key) => acc == null ? undefined : acc[key], obj) ?? fallback;
      } catch (_) {
        return fallback;
      }
    };
    function pick(obj) {
      return {
        VR: obj.R.VR,
        Vsit: obj.R.Vsit,
        Mz: obj.R.Mz,
        Ms: obj.R.Ms,
        Mt: obj.R.Mt,
        Md: obj.R.Md,
        Mlee: obj.R.Mlee,
        qz: obj.R.qz,
        Cpi: obj.R.Cpi,
        Kv: obj.R.Kv,
        h: obj.R.h,
        Kr: obj.R.Kr,
        CpRW_max: obj.R.CpRW_max,
        tc: obj.R.tc,
        effW: obj.R.effW,
        effD: obj.R.effD,
        windward_p: get(obj, ['R','faces','windward','p']),
        leeward_p: get(obj, ['R','faces','leeward','p']),
        roof_ww_p: get(obj, ['R','faces','roof_ww','p']),
        roof_lw_p: get(obj, ['R','faces','roof_lw','p']),
        roof_cw_p: get(obj, ['R','faces','roof_cw','p']),
        side1_zone1_p: get(obj, ['R','faces','sidewall1','zones',0,'p']),
        side1_zone2_p: get(obj, ['R','faces','sidewall1','zones',1,'p']),
        side1_zone3_p: get(obj, ['R','faces','sidewall1','zones',2,'p']),
        roofClause: get(obj, ['R','faces','roof_ww','clause'])
      };
    }
    const out = [];
    for (const cfg of cases) {
      Object.assign(S, cfg);
      calc();
      out.push({ name: cfg.name, result: pick(S) });
    }
    return out;
  })()`;
}

function diffNumber(a, b) {
  return Math.abs(a - b);
}

function compare(liveRows, localRows, tolerance = 1e-9) {
  const localMap = new Map(localRows.map(r => [r.name, r.result]));
  const mismatches = [];
  for (const row of liveRows) {
    const other = localMap.get(row.name);
    if (!other) {
      mismatches.push({ name: row.name, reason: 'missing_local' });
      continue;
    }
    const bad = {};
    for (const key of Object.keys(row.result)) {
      const a = row.result[key];
      const b = other[key];
      if (typeof a === 'number' && typeof b === 'number') {
        const d = diffNumber(a, b);
        if (d > tolerance) bad[key] = { live: a, local: b, absDiff: d };
      } else if (a !== b) {
        bad[key] = { live: a, local: b };
      }
    }
    if (Object.keys(bad).length) mismatches.push({ name: row.name, diff: bad });
  }
  return mismatches;
}

function writeReports(report) {
  fs.writeFileSync(OUT_JSON, JSON.stringify(report, null, 2) + '\n');
  const lines = [];
  lines.push('# Bulk Numeric Parity Report');
  lines.push('');
  lines.push(`- Cases: ${report.caseCount}`);
  lines.push(`- Seed: ${report.seed}`);
  lines.push(`- Mismatches: ${report.mismatchCount}`);
  lines.push(`- Tolerance: ${report.tolerance}`);
  lines.push('');
  lines.push('## Coverage');
  lines.push('- Randomized geometry, roof types, load cases, regions, terrain categories, openings, wind angles');
  lines.push('- Randomized directional arrays for `Md`, `TC_dir`, `Ms`, `Mt`, `Mlee`');
  lines.push('- Compared live and local via direct DevTools evaluation of `calc()` outputs');
  lines.push('');
  if (report.mismatchCount === 0) {
    lines.push('## Result');
    lines.push('All 100 generated cases matched exactly within tolerance.');
  } else {
    lines.push('## Result');
    lines.push(`Found ${report.mismatchCount} mismatched cases.`);
    lines.push('');
    lines.push('## First mismatches');
    for (const item of report.mismatches.slice(0, 10)) {
      lines.push(`- ${item.name}: ${JSON.stringify(item.diff || item.reason)}`);
    }
  }
  lines.push('');
  lines.push('## Notes');
  lines.push('- This validates the numeric engine for controlled/manual state injection.');
  lines.push('- Next risk surface remains map-driven auto-detection and external data parity.');
  fs.writeFileSync(OUT_MD, lines.join('\n') + '\n');
}

(async () => {
  const seed = 20260429;
  const cases = generateCases(100, seed);
  const liveTarget = await openTarget(LIVE_URL);
  const localTarget = await openTarget(LOCAL_URL);
  await waitForReady(liveTarget.webSocketDebuggerUrl);
  await waitForReady(localTarget.webSocketDebuggerUrl);
  const expr = buildExpression(cases);
  const liveRows = await evalExpr(liveTarget.webSocketDebuggerUrl, expr, 30000);
  const localRows = await evalExpr(localTarget.webSocketDebuggerUrl, expr, 30000);
  const tolerance = 1e-9;
  const mismatches = compare(liveRows, localRows, tolerance);
  const report = {
    seed,
    caseCount: cases.length,
    tolerance,
    mismatchCount: mismatches.length,
    mismatches,
    sampleCases: cases.slice(0, 5)
  };
  writeReports(report);
  console.log(JSON.stringify({
    caseCount: report.caseCount,
    mismatchCount: report.mismatchCount,
    outJson: OUT_JSON,
    outMd: OUT_MD
  }));
})().catch(err => {
  console.error(err);
  process.exit(1);
});

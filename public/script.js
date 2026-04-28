/* ═══════════════════════════════════════════════════════════════════
   Wind Analysis Tool
   AS/NZS 1170.2:2021 compliant calculations
   With Firebase Auth + Local-First Workspace
   ═══════════════════════════════════════════════════════════════════ */

// ═════════════ LOCAL WORKSPACE MODE ═════════════
let firebaseApp, firebaseAuth, firestore, googleProvider, microsoftProvider;
let currentUser = null;
/** Prefer this for gating and Bearer tokens: `currentUser` is only set in onAuthStateChanged, which can lag behind firebaseAuth.currentUser right after load. */
function getActiveFirebaseUser(){
  if(currentUser) return currentUser;
  if(typeof firebaseAuth !== 'undefined' && firebaseAuth && firebaseAuth.currentUser) return firebaseAuth.currentUser;
  return null;
}

/** Resolves when Firebase Auth has finished restoring the session from persistence. */
function ensureFirebaseAuthReady(){
  return Promise.resolve();
}

function updateIfcUploadSigninHint(){
  var el = document.getElementById('ifc-signin-hint');
  if(!el) return;
  el.textContent = 'IFC/BIM files load in your browser — no sign-in required.';
  el.style.color = 'var(--text2)';
}

let userSubscription = { active: true, plan: 'local', teamInvitee: false };

const PLAN_LEVELS = { local: 1 };
const PLAN_LABELS = { local: 'Local Workspace' };
const PLAN_MAX_MEMBERS = { local: 1 };
const PLAN_FEATURES = {
  local: { pdf: true, save: true, shared: false, api: false, batch: false, sso: false, templates: false, branding: false, activityLog: false, cloudAi: true }
};
function hasPlanFeature(feature){ return PLAN_FEATURES[userSubscription.plan]?.[feature] || false; }
function isPaidPlan(){ return true; }
function hasSharedProjects(){ return false; }
function canAccessSharedProjects(){ return false; }

const TERRAIN_POLAR_FREE_USES = 10;
function terrainPolarStorageKey(){
  if(!currentUser || !currentUser.uid) return null;
  return 'cw_terrain_polar_uses_' + currentUser.uid;
}
function isTerrainPolarTab(tab){
  return tab === 'Mzcat' || tab === 'Ms' || tab === 'Mt';
}
function getTerrainPolarUseCount(){
  const k = terrainPolarStorageKey();
  if(!k) return 0;
  try{
    const v = localStorage.getItem(k);
    const n = parseInt(v, 10);
    return Number.isFinite(n) && n >= 0 ? n : 0;
  }catch(e){ return 0; }
}
function incrementTerrainPolarUseCount(){
  if(!currentUser || !currentUser.uid || isPaidPlan()) return;
  const k = terrainPolarStorageKey();
  if(!k) return;
  try{
    const n = getTerrainPolarUseCount();
    localStorage.setItem(k, String(n + 1));
  }catch(e){}
}
function canUseTerrainPolarInteraction(){
  if(!isTerrainPolarTab(S.activeDirTab)) return true;
  if(isPaidPlan()) return true;
  if(!currentUser) return false;
  return getTerrainPolarUseCount() < TERRAIN_POLAR_FREE_USES;
}

const MAP_OVERLAY_FREE_USES = 10;
/** overlayId: 'tc' | 'ms' | 'mt' — separate free-tier limits per map overlay button (each button press counts). */
function mapOverlayUsesStorageKey(overlayId){
  if(!currentUser || !currentUser.uid) return null;
  return 'cw_map_overlay_uses_' + overlayId + '_' + currentUser.uid;
}
function getMapOverlayPressCount(overlayId){
  const k = mapOverlayUsesStorageKey(overlayId);
  if(!k) return 0;
  try{
    const v = localStorage.getItem(k);
    const n = parseInt(v, 10);
    return Number.isFinite(n) && n >= 0 ? n : 0;
  }catch(e){ return 0; }
}
function incrementMapOverlayPressCount(overlayId){
  if(!currentUser || !currentUser.uid || isPaidPlan()) return;
  const k = mapOverlayUsesStorageKey(overlayId);
  if(!k) return;
  try{
    const n = getMapOverlayPressCount(overlayId);
    localStorage.setItem(k, String(n + 1));
  }catch(e){
    console.warn('map overlay use counter: localStorage failed', e);
  }
}
/** Each map overlay button press: false = show auth or upgrade and abort toggle. */
function mapOverlayPressAllowed(overlayId){
  return true;
}

/** Pro / not signed in (no quota) / still within free overlay press quota — map-based analysis may run. */
function mapOverlayFreeTierAllowsMapAnalysis(overlayId){
  if(isPaidPlan()) return true;
  if(!currentUser) return true;
  return getMapOverlayPressCount(overlayId) < MAP_OVERLAY_FREE_USES;
}

/** Turn off TC/Ms/Mt map layers when the free-tier press count no longer allows that overlay (e.g. pin moved while "on"). */
function deactivateMapOverlaysIfFreeTierExceeded(){
  if(isPaidPlan() || !currentUser) return;
  if(S.overlayTC && !mapOverlayFreeTierAllowsMapAnalysis('tc')){
    S.overlayTC = false;
    clearTCZones();
    const btn = document.getElementById('btn-ov-tc');
    if(btn) btn.classList.remove('active');
  }
  if(S.overlayMs && !mapOverlayFreeTierAllowsMapAnalysis('ms')){
    S.overlayMs = false;
    clearMsOverlay();
    const btn = document.getElementById('btn-ov-ms');
    if(btn) btn.classList.remove('active');
    const legend = document.getElementById('map-ms-legend');
    if(legend) legend.style.display = 'none';
  }
  if(S.overlayMt && !mapOverlayFreeTierAllowsMapAnalysis('mt')){
    S.overlayMt = false;
    clearMtOverlay();
    const btn = document.getElementById('btn-ov-mt');
    if(btn) btn.classList.remove('active');
    const legend = document.getElementById('map-mt-legend');
    if(legend) legend.style.display = 'none';
  }
}

// Local build: Firebase removed.

// ═════════════ STATE ═════════════
const S = {
  width:20, depth:15, height:6, pitch:15,
  roofType:'gable',
  parapet:0, overhang:0.5,
  windSpeed:45, terrainCat:3, importance:2, loadCase:'A',
  Kp:'1.0', // Permeable cladding reduction factor (Table 5.8)
  windAngle:315,
  openWW:5, openLW:5, openSW:5,
  // site
  lat:-36.8485, lng:174.7633, address:'', elevation:'',
  region:'NZ1', ari:500, svcVr:32,
  life:50,
  /** When false, V_R follows 1170.2 Table 3.1 at ARI from 1170.0 Table 3.3 (IL + design life). */
  vrManual:false,
  mapBuildingAngle:0,
  // 8-direction multipliers (N,NE,E,SE,S,SW,W,NW)
  dirs:['N','NE','E','SE','S','SW','W','NW'],
  Md:  [0.90,0.95,0.95,0.95,0.90,1.00,1.00,0.95],
  Mzcat:[0.83,0.83,0.83,0.83,0.83,0.83,0.83,0.83],
  Ms:  [1,1,1,1,1,1,1,1],
  Mt:  [1,1,1,1,1,1,1,1],       // combined topographic multiplier (Mh ⊗ Mlee per region)
  Mt_hill:[1,1,1,1,1,1,1,1],   // hill-shape Mh from elevation only (Cl 4.4.2)
  Mlee:[1,1,1,1,1,1,1,1],
  leeZone:null,
  leeOverride:false,
  manualShieldBuildings:[],  // user-added shielding buildings [{lat,lng,height,breadth}]
  Vsit_dir:[0,0,0,0,0,0,0,0],
  TC_dir:[null,null,null,null,null,null,null,null],
  overlayTC:false,
  overlayMs:false,
  overlayMt:false,
  // auto-detection raw data (for PDF report)
  detectedBuildingsPerSector:null,
  detectedBuildingsList:null,
  /** Cached context from last Overpass detect — allows TC/Ms recompute without re-fetch */
  terrainRecalcCtx:null,
  detectedNearBuildings:null,
  detectedSectorWater:null,
  detectedSectorOpen:null,
  detectedElevations:null,
  detectedSiteElev:null,
  detectedSampleDistances:null,
  mhSub:null,
  mhDetailsSub:null,
  elevBearingsSub:null,
  nSubDirs:32,
  detectionTimestamp:null,
  /** OSM/Overpass data quality for map TC/Ms: live | cached | fallback | loading | unknown */
  terrainDataSource:'unknown',
  // toggles
  showHeatmap:true, showParticles:false, showPressureArrows:false,
  showInternalPressure:false, showDimensions:true,
  showLabels:true, showShadows:true,
  showGrid:true,
  darkMode:true,
  viewMode:'solid',
  // active tab
  activeTab:'site',
  activeStructureSubtab:'model',
  activeSiteSubtab:'location',
  activeDirTab:'Mzcat',
  /** Open sector popover index (polar UI); null when closed. */
  dirPolarPopoverIdx:null,
  activePressureTab:'heatmap',
  kvOverride:false,
  /** When kvOverride: parsed m³ from #inp-vol-kv (Cl 5.3.4). */
  volKvManual:NaN,
  // lock/unlock
  analysisLocked:false,
  lockedResults:null,
  lockedSiteLat:null,
  lockedSiteLng:null,
  // pressure map
  showPressureMap:false,
  /** User opt-in: optional cloud refinement for low-confidence uploaded mesh classification (IFC/BIM). */
  ifcAiAssist:false,
  // computed
  R:{},
  // user
  user:{ name:'Guest User', plan:'free', signedIn:false },
};

let scene, camera, renderer, controls;
let grpBuild, grpDim, grpLabel, grpArrows, grpInternal, grpWind;
let particleSys = null;
let gridHelper;
let raycaster, mouse;
let hoveredFace = null;
/** Face (#face-info) tooltip: anchor position per mesh uuid so it does not track every mousemove. */
let faceInfoAnchorUuid = null;
let faceInfoPointerInside = false;
let faceInfoHideTimer = null;
let faceMap = new Map();
let pressureTableHoverTr = null;
let compassDrag = false;
let clock = 0;

// Leaflet
let leafletMap, mapMarker, buildingPolygon, mapSatelliteLayer, mapStreetLayer, mapCurrentLayer='street';
let tcOverlayLayers = [];
let msOverlayLayers = [];
let msManualLayers = [];  // separate layer group for manual shielding markers
let mtOverlayLayers = [];

/** Prevents overlapping auto-detect runs (heavy network + main-thread work). */
let autoDetectInFlight = false;
/** If true, run one more detect after the current one finishes (rapid pin moves / double-clicks). */
let autoDetectQueued = false;
/** Options passed through to the queued `autoDetectAllMultipliers` retry (e.g. fromPinMove). */
let autoDetectQueuedOpts = null;
/** Bumps when the site pin moves or Detect is re-triggered — stale async work must not apply results. */
let siteDetectGeneration = 0;
/** AbortController for the active auto-detect network work (Overpass); pin move aborts this. */
let activeDetectAbort = null;
/** True while Overpass/OSM work for TC + M<sub>s</sub> (and M<sub>z,cat</sub>) is still running. */
let detectPendingOsm = false;
/** True while elevation sampling for M<sub>t</sub> is still running. */
let detectPendingElev = false;

// ── Overpass / OSM reliability (client-side; optional proxy: window.CW_OVERPASS_PROXY_URL) ──
const CW_OVERPASS_PUBLIC_ENDPOINTS = [
  'https://overpass.kumi.systems/api/interpreter',
  'https://overpass-api.de/api/interpreter',
  'https://overpass.openstreetmap.ru/api/interpreter'
];
const CW_OVERPASS_CACHE_PREFIX = 'cw_osm_el_v1';
const CW_OVERPASS_CACHE_TTL_MS = 15 * 60 * 1000;
/** Max age for emergency reuse when every live fetch fails (still same cache key). */
const CW_OVERPASS_CACHE_STALE_MAX_MS = 7 * 24 * 60 * 60 * 1000;
const CW_OVERPASS_DEADLINE_MS = 120000;
const CW_OVERPASS_PER_ATTEMPT_MS = 28000;
const CW_OVERPASS_ATTEMPTS_PER_HOST = 3;

function cwOverpassCacheKey(lat, lng, radius, fastDetect){
  const hybrid = (typeof window !== 'undefined' && window.CW_BUILDINGS_HYBRID_URL) ? 'h1' : 'h0';
  return [Number(lat).toFixed(4), Number(lng).toFixed(4), String(radius), fastDetect ? 'fast' : 'full', hybrid].join('|');
}

/** POST /api/buildings-hybrid — same-origin when set (Overpass + optional LINZ on server). */
function cwBuildingsHybridUrl(){
  if(typeof window === 'undefined' || !window.CW_BUILDINGS_HYBRID_URL) return '';
  const u = String(window.CW_BUILDINGS_HYBRID_URL).trim();
  if(!u || (u.indexOf('http') !== 0 && u.charAt(0) !== '/')) return '';
  return u;
}

/** Human-readable LINZ line for terrain status (server returns `linz` from /api/buildings-hybrid). */
function formatTerrainLiveDetail(viaProxy, linz){
  const parts = [];
  if(viaProxy) parts.push('via proxy');
  if(linz && typeof linz === 'object'){
    if(linz.skipped){
      if(linz.reason === 'not_configured') parts.push('LINZ: API key not set (CW_LINZ_API_KEY on server)');
      else if(linz.reason === 'outside_nz_bbox') parts.push('LINZ: N/A outside NZ');
      else if(linz.reason) parts.push('LINZ: skipped (' + linz.reason + ')');
    } else {
      const err = linz.reason ? ' — ' + linz.reason : '';
      parts.push('LINZ +' + (linz.added | 0) + ' buildings' + err);
    }
  }
  return parts.length ? parts.join(' — ') : null;
}
function readCwOsmCache(ck){
  try{
    const raw = sessionStorage.getItem(CW_OVERPASS_CACHE_PREFIX + ':' + ck);
    if(!raw) return null;
    const o = JSON.parse(raw);
    if(!o || typeof o.t !== 'number' || !Array.isArray(o.elements)) return null;
    if(Date.now() - o.t > CW_OVERPASS_CACHE_TTL_MS) return null;
    return o.elements;
  } catch(e){ return null; }
}
/** Same key as readCwOsmCache but ignores freshness TTL (used only after all Overpass attempts fail). */
function readCwOsmCacheStaleFallback(ck){
  try{
    const raw = sessionStorage.getItem(CW_OVERPASS_CACHE_PREFIX + ':' + ck);
    if(!raw) return null;
    const o = JSON.parse(raw);
    if(!o || typeof o.t !== 'number' || !Array.isArray(o.elements)) return null;
    if(Date.now() - o.t > CW_OVERPASS_CACHE_STALE_MAX_MS) return null;
    return o.elements;
  } catch(e){ return null; }
}
function writeCwOsmCache(ck, elements){
  try{
    const payload = JSON.stringify({ t: Date.now(), elements });
    if(payload.length > 4500000) return;
    sessionStorage.setItem(CW_OVERPASS_CACHE_PREFIX + ':' + ck, payload);
  } catch(e){}
}
function sleepMs(ms){ return new Promise(r=>setTimeout(r, ms)); }

function setTerrainDataStatus(source, detail){
  S.terrainDataSource = source || 'unknown';
  const rows = {
    live: { text: 'Terrain data: live (OpenStreetMap)', variant: 'terrain-data-status--live' },
    cached: { text: 'Terrain data: cached OSM (network skipped)', variant: 'terrain-data-status--cached' },
    fallback: { text: 'Terrain data: conservative defaults — OSM unavailable', variant: 'terrain-data-status--fallback' },
    loading: { text: 'Terrain data: loading…', variant: 'terrain-data-status--loading' },
    unknown: { text: '', variant: '' }
  };
  const row = rows[source] || rows.unknown;
  const targets = [
    { id: 'map-terrain-data-status', base: 'map-terrain-data-status' },
    { id: 'panel-terrain-data-status', base: 'panel-terrain-data-status' }
  ];
  const fullText = row.text && detail ? row.text + ' — ' + detail : row.text;
  const title = fullText || '';
  for(const t of targets){
    const el = document.getElementById(t.id);
    if(!el) continue;
    if(!row.text){
      el.style.display = 'none';
      el.textContent = '';
      el.className = t.base;
      el.removeAttribute('title');
      continue;
    }
    el.style.display = 'block';
    el.className = t.base + (row.variant ? ' ' + row.variant : '');
    el.textContent = fullText;
    el.title = title;
  }
}

/** Firebase Bearer for same-origin detection APIs when CW_DETECTION_API_REQUIRE_AUTH is enabled. */
async function cwDetectionAuthHeaders(opts){
  opts = opts || {};
  const forceRefresh = !!opts.forceRefresh;
  try{
    if(typeof firebase === 'undefined' || !firebaseAuth) return {};
    const u = getActiveFirebaseUser();
    if(!u) return {};
    const idToken = forceRefresh ? await u.getIdToken(true) : await u.getIdToken();
    return idToken ? { Authorization: 'Bearer '+idToken } : {};
  } catch(e){
    return {};
  }
}

function cwOverpassEndpointList(){
  const out = [];
  const proxy = typeof window !== 'undefined' && window.CW_OVERPASS_PROXY_URL;
  if(proxy && typeof proxy === 'string'){
    const u = proxy.trim();
    if(u.indexOf('http') === 0 || u.charAt(0) === '/'){
      out.push({ url: u, isProxy: true });
    }
  }
  const pubs = CW_OVERPASS_PUBLIC_ENDPOINTS.slice();
  for(let i = pubs.length - 1; i > 0; i--){
    const j = Math.floor(Math.random() * (i + 1));
    const t = pubs[i]; pubs[i] = pubs[j]; pubs[j] = t;
  }
  pubs.forEach(url=>{ out.push({ url, isProxy: false }); });
  return out;
}

/**
 * POST Overpass QL with retries, jittered backoff, and global deadline.
 * When lat/lng are provided and CW_BUILDINGS_HYBRID_URL is set, tries /api/buildings-hybrid first (LINZ merge on server).
 * Returns { elements, cancelled, reason?, viaProxy?, linz? }.
 */
async function fetchOverpassElementsForSite({ query, runGen, ac, lat, lng }){
  const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
  const deadline = t0 + CW_OVERPASS_DEADLINE_MS;
  const authHdrs = await cwDetectionAuthHeaders();
  const hybridUrl = cwBuildingsHybridUrl();
  if(hybridUrl && Number.isFinite(lat) && Number.isFinite(lng)){
    for(let attempt = 0; attempt < 2; attempt++){
      if(runGen !== siteDetectGeneration) return { elements: null, cancelled: true };
      if(ac.signal.aborted) return { elements: null, cancelled: true };
      const now = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
      if(now > deadline) break;
      if(attempt > 0) await sleepMs(400 + Math.random() * 500);

      const attemptCtrl = new AbortController();
      const timeoutId = setTimeout(()=>attemptCtrl.abort(), CW_OVERPASS_PER_ATTEMPT_MS);
      const onParentAbort = ()=>{ clearTimeout(timeoutId); attemptCtrl.abort(); };
      ac.signal.addEventListener('abort', onParentAbort);
      try{
        const hdrs = { 'Content-Type': 'application/json' };
        if(authHdrs.Authorization) hdrs['Authorization'] = authHdrs.Authorization;
        const resp = await fetch(hybridUrl, {
          method:'POST',
          headers: hdrs,
          body: JSON.stringify({ query, lat, lng }),
          signal: attemptCtrl.signal
        });
        clearTimeout(timeoutId);
        ac.signal.removeEventListener('abort', onParentAbort);
        if(resp.ok){
          const data = await resp.json().catch(()=>null);
          if(data && Array.isArray(data.elements)){
            if(data.linz){
              console.info('Wind Analysis: LINZ hybrid status', data.linz);
            }
            return {
              elements: data.elements,
              cancelled: false,
              reason: null,
              endpoint: hybridUrl,
              viaProxy: true,
              linz: data.linz || null
            };
          }
        } else {
          console.warn('Wind Analysis: /api/buildings-hybrid HTTP', resp.status, '— falling back to Overpass');
        }
      } catch(e){
        clearTimeout(timeoutId);
        ac.signal.removeEventListener('abort', onParentAbort);
        if(runGen !== siteDetectGeneration) return { elements: null, cancelled: true };
        if(ac.signal.aborted) return { elements: null, cancelled: true };
        console.warn('Wind Analysis: hybrid buildings fetch failed — falling back to Overpass', hybridUrl, e && e.message);
      }
    }
  }

  const endpoints = cwOverpassEndpointList();
  let lastReason = 'failed';

  for(let ei = 0; ei < endpoints.length; ei++){
    const ep = endpoints[ei];
    for(let attempt = 0; attempt < CW_OVERPASS_ATTEMPTS_PER_HOST; attempt++){
      if(runGen !== siteDetectGeneration) return { elements: null, cancelled: true };
      if(ac.signal.aborted) return { elements: null, cancelled: true };
      const now = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
      if(now > deadline){
        return { elements: null, cancelled: false, reason: 'timeout' };
      }
      if(attempt > 0) await sleepMs(400 + Math.random() * 900 + attempt * 180);

      const attemptCtrl = new AbortController();
      const timeoutId = setTimeout(()=>attemptCtrl.abort(), CW_OVERPASS_PER_ATTEMPT_MS);
      const onParentAbort = ()=>{ clearTimeout(timeoutId); attemptCtrl.abort(); };
      ac.signal.addEventListener('abort', onParentAbort);
      try{
        const hdrs = { 'Content-Type': 'application/x-www-form-urlencoded' };
        if(ep.isProxy && authHdrs.Authorization) Object.assign(hdrs, authHdrs);
        const resp = await fetch(ep.url, {
          method:'POST',
          headers: hdrs,
          body: 'data=' + encodeURIComponent(query),
          signal: attemptCtrl.signal
        });
        clearTimeout(timeoutId);
        ac.signal.removeEventListener('abort', onParentAbort);
        if(!resp.ok){
          lastReason = 'http_' + resp.status;
          if(ep.isProxy && resp.status === 401){
            return { elements: null, cancelled: false, reason: 'unauthorized' };
          }
          continue;
        }
        const data = await resp.json().catch(()=>null);
        if(!data){ lastReason = 'bad_json'; continue; }
        const elements = Array.isArray(data.elements) ? data.elements : [];
        return { elements, cancelled: false, reason: null, endpoint: ep.url, viaProxy: !!ep.isProxy, linz: null };
      } catch(e){
        clearTimeout(timeoutId);
        ac.signal.removeEventListener('abort', onParentAbort);
        if(runGen !== siteDetectGeneration) return { elements: null, cancelled: true };
        if(ac.signal.aborted) return { elements: null, cancelled: true };
        lastReason = (e && e.name === 'AbortError') ? 'aborted_attempt' : (e && e.message ? String(e.message).slice(0, 80) : 'fetch_error');
        console.warn('Overpass fetch failed', ep.url, 'attempt', attempt, e && e.message);
      }
    }
  }
  return { elements: null, cancelled: false, reason: lastReason };
}

// 3D Model Upload
let uploadedModelGroup = null;
let uploadedModelVisible = true;
let uploadedModelName = '';
let parametricVisible = true;

// Uploaded model wind analysis
let uploadBBox = null;           // {width, depth, height}
let uploadFaceMap = new Map();   // mesh.uuid -> faceKey (windward/leeward/etc)
/** Per-mesh classification diagnostics (MLP + analytic) for upload hover / QA */
let uploadClassDiag = new Map();
const IFC_AI_REFINE_CONF_MAX = 0.52;
// Phase 9: meshes whose chosen face-key confidence (Phase 8c) sits below this
// threshold are routed to Cloud AI for a second-opinion classification. Matches
// the face-inspector "(low)" tag threshold so users see the same signal the
// refinement queue uses. Only applied to wall/roof faceKeys.
const IFC_AI_REFINE_FACE_CONF_MIN = 0.15;
const UPLOAD_FACE_OVERRIDE_STORAGE = 'structuralwind_upload_face_overrides';
/** Cache for IFC AI batch responses: key = modelName + wind sector */
let ifcAiResultCache = new Map();
let detailedPdfPreviewUrl = null;
const IFC_AI_CACHE_MAX = 48;
const IFC_AI_SESSION_STORAGE = 'structuralwind_ifc_ai_assist';
const IFC_AI_CLASS_LABELS = ['interior','exterior_wall','roof','floor_ground'];
/** Must stay in sync with netlify/functions/ifc-ai-mesh.js IFC_AI_MESH_MAX (default 300). */
const IFC_AI_MESH_CHUNK = 300;
let uploadOrigMaterials = new Map(); // mesh.uuid -> original material
let grpUploadOverlay = null;     // THREE.Group for dims/labels on uploaded model
/** Largest axis of the scaled uploaded model (m) — orbit limits + fog use this so small IFCs are not clipped by minDistance 4. */
let uploadModelCamExtent = 0;
/** Last hovered mesh uuid when viewing uploaded IFC (for manual zone override). */
let lastUploadHoverUuid = null;

// Internal pressure permeability condition state (must be before IIFE init)
let permCondition = 'auto';

function readKvOverrideFromUi(){
  const volGeom = S.width * S.depth * S.height;
  const t = document.getElementById('kv-override-toggle');
  const i = document.getElementById('inp-vol-kv');
  S.kvOverride = !!(t && t.checked);
  if(i){
    if(!S.kvOverride){
      i.readOnly = true;
      i.value = volGeom > 0 ? String(Number(volGeom.toFixed(1))) : '0';
    } else {
      i.readOnly = false;
    }
  }
  const m = i ? parseFloat(i.value) : NaN;
  S.volKvManual = Number.isFinite(m) && m > 0 ? m : NaN;
}

/** Envelope volume and V used in Cl 5.3.4 (manual V when override on). */
function prepareVolumeForKv(){
  readKvOverrideFromUi();
  const volGeom = S.width * S.depth * S.height;
  const vol = (S.kvOverride && Number.isFinite(S.volKvManual) && S.volKvManual > 0)
    ? S.volKvManual
    : volGeom;
  const overridden = S.kvOverride && Math.abs(vol - volGeom) > 1e-3;
  return { volGeom, vol, overridden };
}

/** AS/NZS 1170.2 Cl 5.3.4 — Kv from opening ratio r = 100 × A^1.5 / V */
function computeKvAuto(vol, A){
  let Kv = 1.0;
  let ratio = 0;
  let zone = 'none';
  if(A > 0 && vol > 0){
    ratio = 100 * Math.pow(A, 1.5) / vol;
    if(ratio < 0.09){ Kv = 0.85; zone = 'r_lt_0_09'; }
    else if(ratio <= 3){ Kv = 1.01 + 0.15 * Math.abs(Math.log10(ratio)); zone = 'r_mid'; }
    else { Kv = 1.085; zone = 'r_gt_3'; }
  }
  return { Kv, ratio, zone };
}

function updateKvEquationPopupHtml(vol, volGeom, A, ratio, zone, KvAuto, KvFinal, qz, cpi1, kci1, pi1, volumeOverridden){
  const pop = document.getElementById('kv-equation-popup');
  if(!pop) return;
  if(A <= 0 || vol <= 0){
    pop.innerHTML = '<p><strong>Cl 5.3.4</strong> — With no effective opening area, <i>r</i> is undefined; <i>K<sub>v</sub></i> = 1.0.</p>'
      + (qz != null && cpi1 != null && kci1 != null && pi1 != null
        ? `<p style="margin-top:8px"><strong>Cl 5.3</strong> — <i>p<sub>i</sub></i> = <i>q<sub>z</sub></i><i>C<sub>p,i</sub></i><i>K<sub>c,i</sub></i><i>K<sub>v</sub></i> = ${qz.toFixed(3)}×${cpi1.toFixed(2)}×${kci1.toFixed(1)}×${KvFinal.toFixed(3)} = <strong>${pi1.toFixed(2)} kPa</strong> (case 1)</p>`
        : '');
    return;
  }
  const rStr = ratio.toFixed(4);
  let branchHtml = '';
  if(zone === 'r_lt_0_09'){
    branchHtml = '<p>For <i>r</i> &lt; 0.09 → <i>K<sub>v</sub></i> = <strong>0.85</strong></p>';
  } else if(zone === 'r_mid'){
    const mid = (1.01 + 0.15 * Math.abs(Math.log10(ratio))).toFixed(4);
    branchHtml = `<p>For 0.09 ≤ <i>r</i> ≤ 3 → <i>K<sub>v</sub></i> = 1.01 + 0.15|log<sub>10</sub><i>r</i>| = 1.01 + 0.15×|log<sub>10</sub>(${rStr})| = <strong>${mid}</strong></p>`;
  } else if(zone === 'r_gt_3'){
    branchHtml = '<p>For <i>r</i> &gt; 3 → <i>K<sub>v</sub></i> = <strong>1.085</strong></p>';
  }
  const piLine = qz != null && cpi1 != null && kci1 != null && pi1 != null
    ? `<p style="margin-top:8px;padding-top:8px;border-top:1px solid var(--border)"><strong>Cl 5.3 — internal pressure (case 1)</strong><br>`
      + `<i>p<sub>i</sub></i> = <i>q<sub>z</sub></i> <i>C<sub>p,i</sub></i> <i>K<sub>c,i</sub></i> <i>K<sub>v</sub></i> = ${qz.toFixed(3)} × ${cpi1.toFixed(2)} × ${kci1.toFixed(1)} × ${KvFinal.toFixed(3)} = <strong>${pi1.toFixed(2)} kPa</strong></p>`
    : '';
  pop.innerHTML = `<p><strong>Cl 5.3.4 — Opening ratio</strong></p>`
    + `<p><i>r</i> = 100 <i>A</i><sup>1.5</sup> / <i>V</i> = 100 × ${A.toFixed(2)}<sup>1.5</sup> / ${vol.toFixed(1)} = <strong>${rStr}</strong></p>`
    + branchHtml
    + `<p>Computed <i>K<sub>v</sub></i> = <strong>${KvAuto.toFixed(4)}</strong>${volumeOverridden && volGeom > 0 ? ` <span style="color:var(--text2)">(manual <i>V</i> = ${vol.toFixed(1)} m³; envelope ${volGeom.toFixed(1)} m³)</span>` : ''}</p>`
    + piLine;
}

function setupKvEquationHover(){
  const sec = document.querySelector('.kv-override-section');
  const pop = document.getElementById('kv-equation-popup');
  if(!sec || !pop || sec.dataset.kvHoverInit) return;
  sec.dataset.kvHoverInit = '1';
  const place = ()=>{
    const r = sec.getBoundingClientRect();
    const w = Math.min(380, Math.max(260, r.width + 24));
    pop.style.width = w + 'px';
    pop.style.left = Math.max(8, Math.min(innerWidth - w - 8, r.left + r.width / 2 - w / 2)) + 'px';
    const ph = pop.offsetHeight || 200;
    let top = r.top - ph - 8;
    if(top < 8) top = r.bottom + 8;
    pop.style.top = top + 'px';
  };
  const show = () => {
    pop.hidden = false;
    requestAnimationFrame(place);
  };
  const hide = () => { pop.hidden = true; };
  sec.addEventListener('mouseenter', show);
  sec.addEventListener('mouseleave', hide);
  window.addEventListener('resize', ()=>{ if(!pop.hidden) place(); });
}

function onKvOverrideToggle(){
  readKvOverrideFromUi();
  if(S.analysisLocked) return;
  onInput();
}

function onVolKvChange(){
  if(S.analysisLocked) return;
  onInput();
}

// ═════════════ INIT ═════════════
(function init(){
  if(typeof THREE === 'undefined'){
    console.error('Wind Analysis: THREE.js did not load — check CDN / network.');
    return;
  }
  const box = document.getElementById('canvas-container');
  if(!box){
    console.error('Wind Analysis: #canvas-container not found.');
    return;
  }
  scene = new THREE.Scene();
  setSceneBg();

  camera = new THREE.PerspectiveCamera(50, innerWidth/innerHeight, 0.1, 500);
  camera.position.set(32,22,38);

  renderer = new THREE.WebGLRenderer({antialias:true, alpha:true, preserveDrawingBuffer:true});
  renderer.setPixelRatio(Math.min(devicePixelRatio,2));
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  renderer.toneMapping = THREE.ACESFilmicToneMapping;
  renderer.toneMappingExposure = 1.15;
  box.appendChild(renderer.domElement);

  controls = new THREE.OrbitControls(camera, renderer.domElement);
  controls.enableDamping = true;
  controls.dampingFactor = 0.08;
  controls.maxPolarAngle = Math.PI/2 - 0.04;
  controls.minDistance = 4;
  controls.maxDistance = 160;

  addLights();
  addGround();

  grpBuild   = new THREE.Group(); scene.add(grpBuild);
  grpDim     = new THREE.Group(); scene.add(grpDim);
  grpLabel   = new THREE.Group(); scene.add(grpLabel);
  grpArrows  = new THREE.Group(); scene.add(grpArrows);
  grpInternal= new THREE.Group(); scene.add(grpInternal);
  grpWind    = new THREE.Group(); scene.add(grpWind);

  gridHelper = new THREE.GridHelper(120,60,0x444466,0x222244);
  gridHelper.position.y = 0.01;
  scene.add(gridHelper);

  raycaster = new THREE.Raycaster();
  mouse = new THREE.Vector2(-9,-9);

  renderer.domElement.addEventListener('mousemove', onPointerMove);
  renderer.domElement.addEventListener('click', onPointerClick);
  renderer.domElement.addEventListener('mouseleave', (ev)=>{
    const fi = document.getElementById('face-info');
    const rt = ev.relatedTarget;
    if(fi && rt && (fi === rt || (typeof fi.contains === 'function' && fi.contains(rt)))) return;
    hideFaceInfoTooltip();
  });
  const faceInfoEl = document.getElementById('face-info');
  if(faceInfoEl){
    faceInfoEl.addEventListener('pointerenter', ()=>{
      faceInfoPointerInside = true;
      clearFaceInfoHideTimer();
    });
    faceInfoEl.addEventListener('pointerleave', (ev)=>{
      const rt = ev.relatedTarget;
      if(rt && typeof faceInfoEl.contains === 'function' && faceInfoEl.contains(rt)) return;
      faceInfoPointerInside = false;
      scheduleFaceInfoHide();
    });
    const fiOverrideSel = document.getElementById('fi-override-select');
    if(fiOverrideSel) fiOverrideSel.addEventListener('focus', clearFaceInfoHideTimer);
  }
  window.addEventListener('resize', onResize);
  window.addEventListener('keydown', e=>{
    if(e.key==='Escape'){
      closePaymentModal();
      closeAuthOverlay();
      hidePressureHoverTip();
      closePressureCalcModal();
      hideFaceInfoTooltip();
    }
  });

  setupCompass();
  initPressureHoverTipDelegate();
  initPressureCalcClickDelegate();
  initPressureTableRowHoverDelegate();
  setupKvEquationHover();
  switchMainTab('site');
  try{
    onInput();
  } catch(err){
    console.error('Wind Analysis: onInput failed during init —', err);
    try{ calc(); rebuild(); } catch(e2){ console.error('Wind Analysis: fallback calc failed', e2); }
  }
  animate();

  // Init map on next frame(s) after site tab layout — no long delay so the pin is usable immediately
  requestAnimationFrame(()=>{
    try{
      initMap();
      if(leafletMap){
        requestAnimationFrame(()=>{
          leafletMap.invalidateSize();
          leafletMap.setView([S.lat, S.lng], leafletMap.getZoom() || 17);
          // Extra passes — #map-container can still be 0×0 after one frame on some layouts (pin clicks no-op until Ms/fitBounds).
          setTimeout(()=>{ if(leafletMap){ leafletMap.invalidateSize(); } }, 120);
          setTimeout(()=>{ if(leafletMap){ leafletMap.invalidateSize(); } }, 400);
        });
      }
      const detectedRegion = autoDetectRegion(S.lat, S.lng);
      if(detectedRegion){ applyRegion(detectedRegion); onInput(); }
      checkLeeZone();
      // Same TC/Ms/Mt fetch as overlay buttons — start as soon as the app is up so terrain data is ready on first interaction
      setTimeout(function(){ try{ scheduleInitialTerrainAutoDetectFromLoad(); }catch(e){ console.error('Wind Analysis: initial terrain detect —', e); } }, 0);
    } catch(err){
      console.error('Wind Analysis: map / region init failed —', err);
    }
  });

  // Initialize Firebase auth listener
  initAuthListener();
  void ensureFirebaseAuthReady().then(function(){
    try{ updateIfcUploadSigninHint(); }catch(e){}
  });

  // Check for checkout result in URL
  checkCheckoutResult();
})();

function onResize(){
  const box = document.getElementById('canvas-container');
  if(!box || !renderer) return;
  const canvas = renderer.domElement;
  const rect = canvas.getBoundingClientRect();
  const w = rect.width;
  const h = Math.max(rect.height, 1);
  if(w < 2 || h < 2) return;
  camera.aspect = w / h;
  camera.updateProjectionMatrix();
  renderer.setSize(w, h);
  if(leafletMap) leafletMap.invalidateSize();
}

function setModel3DStackVisible(visible){
  const stack = document.getElementById('model-center-stack');
  if(!stack) return;
  stack.style.display = visible ? 'flex' : 'none';
  stack.setAttribute('aria-hidden', visible ? 'false' : 'true');
}

// ═════════════ TAB SWITCHING ═════════════
function switchMainTab(tab){
  S.activeTab = tab;
  ['structure','site','about'].forEach(t=>{
    const el = document.getElementById('tc-'+t);
    const btn = document.getElementById('mtab-'+t);
    if(el) el.style.display = t===tab ? '' : 'none';
    if(btn) btn.classList.toggle('active', t===tab);
  });

  // Show/hide 3D canvas + tabulated pressures (center stack)
  const compass = document.getElementById('wind-compass');
  const legend = document.getElementById('legend');
  if(tab==='structure'){
    const showCanvas = (S.activeStructureSubtab||'model') === 'model';
    setModel3DStackVisible(showCanvas);
    if(compass) compass.style.display = 'none';
    if(legend) legend.style.display = (showCanvas && S.showPressureMap) ? '' : 'none';
    updateAnalysisTab();
    if(showCanvas){
      setTimeout(onResize, 50);
      recalcPressures();
    }
    if(S.activeStructureSubtab === 'document') updateDocPreview();
  } else {
    setModel3DStackVisible(false);
    if(compass) compass.style.display = 'none';
    if(legend) legend.style.display = 'none';
  }

  // Init map when switching to site (map only on Location subtab)
  if(tab==='site' && leafletMap && (S.activeSiteSubtab||'location')==='location'){
    setTimeout(()=>leafletMap.invalidateSize(), 100);
  }
  if(tab==='site' && S.activeSiteSubtab==='topography'){
    setTimeout(renderTopoProfiles, 60);
  }
}

// ═════════════ STRUCTURE SUBTAB SWITCHING ═════════════
function switchStructureSubtab(tab){
  S.activeStructureSubtab = tab;
  ['model','document'].forEach(t=>{
    const el = document.getElementById('struct-'+t);
    const btn = document.getElementById('sstab-'+t);
    if(el) el.style.display = t===tab ? '' : 'none';
    if(btn) btn.classList.toggle('active', t===tab);
  });

  const compass = document.getElementById('wind-compass');
  const legend = document.getElementById('legend');
  const faceInfo = document.getElementById('face-info');

  if(tab==='model'){
    setModel3DStackVisible(true);
    if(compass) compass.style.display = 'none';
    if(legend) legend.style.display = S.showPressureMap ? '' : 'none';
    setTimeout(onResize, 50);
    updateAnalysisTab();
    recalcPressures();
  } else {
    setModel3DStackVisible(false);
    if(compass) compass.style.display = 'none';
    if(legend) legend.style.display = 'none';
    hideFaceInfoTooltip();
  }
  if(tab==='document') updateDocPreview();
}

// ═════════════ SITE (LOCATION) SUBTAB SWITCHING ═════════════
function switchSiteSubtab(tab){
  S.activeSiteSubtab = tab;
  const locPanel = document.getElementById('site-panel-location');
  const topoPanel = document.getElementById('struct-topography');
  const btnLoc = document.getElementById('site-subtab-location');
  const btnTopo = document.getElementById('site-subtab-topography');
  if(locPanel) locPanel.style.display = tab === 'location' ? '' : 'none';
  if(topoPanel) topoPanel.style.display = tab === 'topography' ? '' : 'none';
  if(btnLoc) btnLoc.classList.toggle('active', tab === 'location');
  if(btnTopo) btnTopo.classList.toggle('active', tab === 'topography');
  if(tab === 'topography') setTimeout(renderTopoProfiles, 60);
  if(tab === 'location' && leafletMap) setTimeout(() => leafletMap.invalidateSize(), 100);
}

// ═════════════ VIEW TOPOGRAPHY PROFILES ═════════════
function toggleAllTopoCards(expand){
  const cards = document.querySelectorAll('#topo-profiles .topo-card');
  cards.forEach(card => {
    if(expand) card.classList.add('expanded');
    else card.classList.remove('expanded');
  });
  if(expand) setTimeout(drawVisibleTopoCanvases, 100);
}

function toggleTopoCard(card){
  card.classList.toggle('expanded');
  if(card.classList.contains('expanded')){
    setTimeout(drawVisibleTopoCanvases, 50);
  }
}

function togglePressureTablesDrawer(){
  const wrap = document.getElementById('model-pressure-tables');
  const card = document.getElementById('pressure-tables-drawer');
  const head = document.getElementById('pressure-tables-drawer-header');
  if(!wrap || !card) return;
  const open = wrap.classList.toggle('drawer-expanded');
  card.classList.toggle('expanded', open);
  if(head) head.setAttribute('aria-expanded', open ? 'true' : 'false');
  const hint = wrap.querySelector('.pressure-drawer-hint');
  if(hint) hint.textContent = open ? 'Click to collapse' : 'Click to expand';
  setTimeout(onResize, 50);
}

function pressureTablesDrawerKeydown(ev){
  if(!ev || (ev.key !== 'Enter' && ev.key !== ' ')) return;
  ev.preventDefault();
  togglePressureTablesDrawer();
}

function drawVisibleTopoCanvases(){
  const container = document.getElementById('topo-profiles');
  if(!container) return;
  const profiles = S.detectedProfiles;
  const mhDetails = S.mhDetailsSub;
  const siteElev = S.detectedSiteElev;
  if(!profiles || !mhDetails) return;
  const nSub = profiles.length;
  const ridgeRise = (S.pitch||0) > 0 && S.roofType !== 'flat' ? Math.tan((S.pitch||0)*Math.PI/180) * (S.depth||10) / 2 : 0;
  const h = (S.height||6) + (S.parapet||0) + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);

  const cards = container.querySelectorAll('.topo-card.expanded');
  cards.forEach(card => {
    const canvas = card.querySelector('canvas');
    if(!canvas || canvas.dataset.drawn === '1') return;
    const si = parseInt(canvas.dataset.subIdx, 10);
    if(isNaN(si) || !profiles[si]) return;
    try {
      drawTopoCanvas(canvas, profiles[si], mhDetails[si], siteElev, h, si, nSub, profiles);
      canvas.dataset.drawn = '1';
    } catch(e) {
      console.error('[Topo] drawTopoCanvas error for dir', si, e);
    }
  });
}

function renderTopoProfiles(){
  const container = document.getElementById('topo-profiles');
  if(!container) return;

  const profiles = S.detectedProfiles;
  const mhDetails = S.mhDetailsSub;
  const mhSub = S.mhSub;
  const siteElev = S.detectedSiteElev;

  if(!profiles || !profiles.length || !mhDetails || !mhDetails.length){
    container.innerHTML = '<p style="color:var(--text2);text-align:center;padding:40px">Run site detection to generate topographic profiles.</p>';
    return;
  }

  const nSub = profiles.length;
  const filter = document.getElementById('topo-filter');
  const mode = filter ? filter.value : 'cardinal';

  const dir32Names = ['N','NbE','NNE','NEbN','NE','NEbE','ENE','EbN',
    'E','EbS','ESE','SEbE','SE','SEbS','SSE','SbE',
    'S','SbW','SSW','SWbS','SW','SWbW','WSW','WbS',
    'W','WbN','WNW','NWbW','NW','NWbN','NNW','NbW'];
  const dir32Full = ['North','North by East','North North East','North East by North',
    'North East','North East by East','East North East','East by North',
    'East','East by South','East South East','South East by East',
    'South East','South East by South','South South East','South by East',
    'South','South by West','South South West','South West by South',
    'South West','South West by West','West South West','West by South',
    'West','West by North','West North West','North West by West',
    'North West','North West by North','North North West','North by West'];
  const dir8Full = ['North','North East','East','South East','South','South West','West','North West'];

  let indices = [];
  if(mode === 'cardinal'){
    for(let i=0; i<8; i++) indices.push(i*4);
  } else if(mode === 'governing'){
    for(let i=0; i<8; i++){
      const c = i*4;
      let bestIdx = c, bestMh = 0;
      for(let off=-2; off<=2; off++){
        const idx = ((c+off)%nSub+nSub)%nSub;
        if(mhSub[idx] > bestMh){ bestMh = mhSub[idx]; bestIdx = idx; }
      }
      indices.push(bestIdx);
    }
  } else {
    for(let i=0; i<nSub; i++) indices.push(i);
  }

  const sectorOf = si => Math.round(si / 4) % 8;

  const gps = S.lat ? S.lat.toFixed(6) + ', ' + S.lng.toFixed(6) : '';
  const ridgeRise = (S.pitch||0) > 0 && S.roofType !== 'flat' ? Math.tan((S.pitch||0)*Math.PI/180) * (S.depth||10) / 2 : 0;
  const h = (S.height||6) + (S.parapet||0) + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);

  container.innerHTML = '';
  indices.forEach(si => {
    const prof = profiles[si];
    const det = mhDetails[si] || {Mh:1, H:0, Lu:0, slope:0, x:0, L1:0, L2:0, crestElev:siteElev, crestDist:0, isEsc:false};
    const bearing = (si * 360 / nSub).toFixed(1);
    const sectorIdx = sectorOf(si);
    const mtVal = (S.Mt && S.Mt[sectorIdx]) || 1;
    const dirShort = dir32Names[si] || si;
    const dirLong = dir32Full[si] || ('Dir ' + si);
    const sectorName = dir8Full[sectorIdx] || '';

    const card = document.createElement('div');
    card.className = 'topo-card';

    const header = document.createElement('div');
    header.className = 'topo-card-header';
    header.onclick = function(){ toggleTopoCard(card); };

    const titleH = document.createElement('h4');
    titleH.innerHTML = `<span class="expand-icon">▶</span>${dirLong.toUpperCase()} (${dirShort} — ${bearing}°) TOPOGRAPHY PROFILE`;
    if(gps) titleH.innerHTML += ` <span style="color:var(--text3);font-weight:normal;font-size:10px">(GPS: ${gps})</span>`;

    const stats = document.createElement('div');
    stats.className = 'topo-card-stats';
    const topoType = det.isEsc ? 'ESCARPMENT' : (det.H > 1 ? 'HILL' : 'FLAT');
    stats.innerHTML = `
      <span>TOPOGRAPHY: <span class="val">${topoType}</span></span>
      <span>h: <span class="val">${h.toFixed(2)} m</span></span>
      <span>L<sub>u</sub>: <span class="val">${det.Lu.toFixed(2)} m</span></span>
      <span>H: <span class="val">${det.H.toFixed(2)} m</span></span>
      <span>UPWIND SLOPE: <span class="val">${det.slope.toFixed(2)}</span></span>
      <span>L<sub>1</sub>: <span class="val">${det.L1.toFixed(1)} m</span></span>
      <span>L<sub>2</sub>: <span class="val">${det.L2.toFixed(1)} m</span></span>
      <span>x: <span class="val">${det.x.toFixed(1)} m</span></span>
      <span>Mh: <span class="val">${det.Mh.toFixed(4)}</span></span>
      <span>Mt (${sectorName}): <span class="val" style="color:${mtVal>1.05?'var(--warn)':'var(--success)'}">${mtVal.toFixed(4)}</span></span>
    `;

    header.appendChild(titleH);
    header.appendChild(stats);
    card.appendChild(header);

    const body = document.createElement('div');
    body.className = 'topo-card-body';
    const canvas = document.createElement('canvas');
    canvas.setAttribute('width', '1200');
    canvas.setAttribute('height', '440');
    canvas.dataset.subIdx = si;
    canvas.style.cssText = 'display:block;width:100%;height:220px;background:#0d1526';
    body.appendChild(canvas);
    card.appendChild(body);
    container.appendChild(card);
  });
}

function drawTopoCanvas(canvas, profile, det, siteElev, buildH, subIdx, nSub, allProfiles, opts){
  opts = opts || {};
  const parent = canvas.parentElement;
  const grandparent = parent ? parent.parentElement : null;
  const parentW = grandparent ? grandparent.clientWidth : (parent ? parent.clientWidth : 800);
  const cw = (opts.fixedWidth != null && opts.fixedWidth > 20) ? opts.fixedWidth : (parentW > 20 ? parentW : 800);
  const ch = 220;
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.round(cw * dpr);
  canvas.height = Math.round(ch * dpr);
  canvas.style.width = cw + 'px';
  canvas.style.height = ch + 'px';
  const ctx = canvas.getContext('2d');
  if(!ctx) return;
  ctx.scale(dpr, dpr);

  ctx.fillStyle = '#0d1526';
  ctx.fillRect(0, 0, cw, ch);

  if(!det) det = {Mh:1, H:0, Lu:0, slope:0, x:0, L1:0, L2:0, crestElev:siteElev, crestDist:0, isEsc:false};

  // Combine upwind (this direction) + downwind (opposite) into a single cross-section
  const oppIdx = (subIdx + Math.floor(nSub/2)) % nSub;
  const upProf = profile || [];
  const dnProf = allProfiles ? allProfiles[oppIdx] : null;

  // Build merged points: negative x = downwind, 0 = site, positive x = upwind
  const pts = [];
  if(dnProf && dnProf.length > 1){
    for(let i=dnProf.length-1; i>=1; i--){
      pts.push({ x: -dnProf[i].dist, y: dnProf[i].elev });
    }
  }
  pts.push({ x: 0, y: siteElev });
  if(upProf.length > 1){
    for(let i=1; i<upProf.length; i++){
      pts.push({ x: upProf[i].dist, y: upProf[i].elev });
    }
  }

  if(pts.length < 3){
    ctx.fillStyle = '#607d8b';
    ctx.font = '12px Segoe UI, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Insufficient elevation data for this direction', cw/2, ch/2);
    return;
  }

  // Axis ranges
  const maxDist = 5000;
  const xMin = -maxDist, xMax = maxDist;
  const allElevs = pts.map(p=>p.y).filter(v=>isFinite(v));
  const eMin = Math.min(...allElevs);
  const eMax = Math.max(...allElevs);
  const elevRange = Math.max(eMax - eMin, 20);
  const yMin = eMin - elevRange * 0.15;
  const yMax = eMax + elevRange * 0.3;

  const padL = 52, padR = 16, padT = 16, padB = 28;
  const pw = cw - padL - padR;
  const ph = ch - padT - padB;

  const toX = v => padL + (v - xMin) / (xMax - xMin) * pw;
  const toY = v => padT + (1 - (v - yMin) / (yMax - yMin)) * ph;

  // ── Local topographic zone shading (Figures 4.3–4.5) ──
  if(det.Mh > 1 && (det.L1 > 0 || det.L2 > 0)){
    const crestDist = det.crestDist; // positive = upwind
    // L1 zone: from crest ± L1 (peak zone for steep slopes)
    if(det.L1 > 0){
      const l1Left = toX(crestDist - det.L1);
      const l1Right = toX(crestDist + det.L1);
      ctx.fillStyle = 'rgba(255,82,82,.06)';
      ctx.fillRect(Math.max(l1Left, padL), padT, Math.min(l1Right, cw-padR) - Math.max(l1Left, padL), ph);
    }
    // L2 zone: extends downwind from crest (toward site)
    if(det.L2 > 0){
      const l2Left = toX(crestDist - det.L2);
      const l2Right = toX(crestDist + det.L2);
      ctx.fillStyle = 'rgba(255,171,0,.04)';
      ctx.fillRect(Math.max(l2Left, padL), padT, Math.min(l2Right, cw-padR) - Math.max(l2Left, padL), ph);
    }
  }

  // Grid lines
  ctx.strokeStyle = '#1e2d4a';
  ctx.lineWidth = 0.5;
  const elevStep = niceStep(elevRange, 4);
  const eGridStart = Math.ceil(yMin / elevStep) * elevStep;
  ctx.font = '9px Segoe UI, sans-serif';
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  for(let e = eGridStart; e <= yMax; e += elevStep){
    const y = toY(e);
    if(y < padT || y > ch - padB) continue;
    ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(cw-padR, y); ctx.stroke();
    ctx.fillStyle = '#607d8b';
    ctx.fillText(e.toFixed(0) + ' m', padL - 4, y);
  }
  const distStep = niceStep(xMax - xMin, 6);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  for(let d = Math.ceil(xMin/distStep)*distStep; d <= xMax; d += distStep){
    const x = toX(d);
    if(x < padL || x > cw - padR) continue;
    ctx.beginPath(); ctx.moveTo(x, padT); ctx.lineTo(x, ch-padB); ctx.stroke();
    ctx.fillStyle = '#607d8b';
    ctx.fillText((d/1000).toFixed(1) + ' km', x, ch - padB + 4);
  }

  // Axis labels
  ctx.fillStyle = '#90a4ae';
  ctx.font = '9px Segoe UI, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('DISTANCE FROM SITE', padL + pw/2, ch - 3);
  ctx.save();
  ctx.translate(10, padT + ph/2);
  ctx.rotate(-Math.PI/2);
  ctx.fillText('ELEVATION (m)', 0, 0);
  ctx.restore();

  // Draw terrain fill
  ctx.beginPath();
  let started = false;
  pts.forEach(p => {
    if(p.x < xMin || p.x > xMax) return;
    const sx = toX(p.x), sy = toY(p.y);
    if(!started){ ctx.moveTo(sx, sy); started = true; }
    else ctx.lineTo(sx, sy);
  });
  const lastVisible = pts.filter(p=>p.x>=xMin && p.x<=xMax);
  if(lastVisible.length > 0){
    ctx.lineTo(toX(lastVisible[lastVisible.length-1].x), toY(yMin));
    ctx.lineTo(toX(lastVisible[0].x), toY(yMin));
  }
  ctx.closePath();
  const grad = ctx.createLinearGradient(0, padT, 0, ch-padB);
  grad.addColorStop(0, 'rgba(139,119,101,.45)');
  grad.addColorStop(0.5, 'rgba(100,80,60,.3)');
  grad.addColorStop(1, 'rgba(60,50,40,.15)');
  ctx.fillStyle = grad;
  ctx.fill();

  // Terrain line
  ctx.beginPath();
  started = false;
  pts.forEach(p => {
    if(p.x < xMin || p.x > xMax) return;
    const sx = toX(p.x), sy = toY(p.y);
    if(!started){ ctx.moveTo(sx, sy); started = true; }
    else ctx.lineTo(sx, sy);
  });
  ctx.strokeStyle = '#a08060';
  ctx.lineWidth = 1.5;
  ctx.stroke();

  // ── Site marker (blue diamond) ──
  const siteX = toX(0), siteY = toY(siteElev);
  ctx.fillStyle = '#2196F3';
  ctx.beginPath();
  ctx.moveTo(siteX, siteY - 6);
  ctx.lineTo(siteX + 5, siteY);
  ctx.lineTo(siteX, siteY + 6);
  ctx.lineTo(siteX - 5, siteY);
  ctx.closePath();
  ctx.fill();
  ctx.strokeStyle = '#fff';
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.fillStyle = '#2196F3';
  ctx.font = 'bold 9px Segoe UI, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('Site', siteX, siteY - 10);

  // ── Crest marker and annotations if Mh > 1 ──
  if(det.Mh > 1 && det.crestDist > 0){
    const crestX = toX(det.crestDist);
    const crestY = toY(det.crestElev);

    // Crest marker (red triangle)
    ctx.fillStyle = '#ff5252';
    ctx.beginPath();
    ctx.moveTo(crestX, crestY - 7);
    ctx.lineTo(crestX + 5, crestY + 2);
    ctx.lineTo(crestX - 5, crestY + 2);
    ctx.closePath();
    ctx.fill();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 0.8;
    ctx.stroke();
    ctx.fillStyle = '#ff5252';
    ctx.font = 'bold 9px Segoe UI, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Crest', crestX, crestY - 11);

    // ── H annotation (vertical dashed line at crest) ──
    if(det.H > 1){
      const baseElev = det.crestElev - det.H;
      const baseY = toY(baseElev);
      ctx.setLineDash([3,3]);
      ctx.strokeStyle = '#ffab00';
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(crestX, crestY + 4);
      ctx.lineTo(crestX, baseY);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = '#ffab00';
      ctx.font = '8px Segoe UI, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText('H=' + det.H.toFixed(1) + 'm', crestX + 4, (crestY + baseY)/2);

      // ── H/2 level horizontal dashed line ──
      const halfElev = det.crestElev - det.H/2;
      const halfY = toY(halfElev);
      ctx.setLineDash([4,4]);
      ctx.strokeStyle = 'rgba(0,210,255,.4)';
      ctx.lineWidth = 0.8;
      ctx.beginPath();
      ctx.moveTo(padL, halfY);
      ctx.lineTo(cw - padR, halfY);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = 'rgba(0,210,255,.6)';
      ctx.font = '8px Segoe UI, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText('H/2', cw - padR - 20, halfY - 4);

      // ── Lu bracket (crest to half-height point upwind) ──
      if(det.Lu > 10){
        const luEndX = toX(det.crestDist + det.Lu);
        const bracketY = crestY - 18;
        ctx.strokeStyle = '#4CAF50';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(crestX, bracketY);
        ctx.lineTo(Math.min(luEndX, cw - padR), bracketY);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(crestX, bracketY-3); ctx.lineTo(crestX, bracketY+3); ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(Math.min(luEndX, cw-padR), bracketY-3); ctx.lineTo(Math.min(luEndX, cw-padR), bracketY+3); ctx.stroke();
        ctx.fillStyle = '#4CAF50';
        ctx.font = '8px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('Lu=' + det.Lu.toFixed(0) + 'm', Math.min((crestX+luEndX)/2, cw-padR-30), bracketY - 5);
      }

      // ── x bracket (site to crest distance) ──
      if(det.x > 5){
        const xBracketY = toY(siteElev) + 14;
        ctx.strokeStyle = '#e040fb';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(siteX, xBracketY);
        ctx.lineTo(crestX, xBracketY);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(siteX, xBracketY-3); ctx.lineTo(siteX, xBracketY+3); ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(crestX, xBracketY-3); ctx.lineTo(crestX, xBracketY+3); ctx.stroke();
        ctx.fillStyle = '#e040fb';
        ctx.font = '8px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('x=' + det.x.toFixed(0) + 'm', (siteX+crestX)/2, xBracketY + 10);
      }

      // ── L1, L2 zone boundary lines (vertical dashed) ──
      [['L1', det.L1, '#ff5252'], ['L2', det.L2, '#ffab00']].forEach(([lbl, dist, col]) => {
        if(dist <= 0) return;
        // Downwind boundary from crest
        const bxDn = toX(det.crestDist - dist);
        if(bxDn >= padL && bxDn <= cw - padR){
          ctx.setLineDash([2,4]);
          ctx.strokeStyle = col;
          ctx.lineWidth = 0.7;
          ctx.beginPath();
          ctx.moveTo(bxDn, padT);
          ctx.lineTo(bxDn, ch - padB);
          ctx.stroke();
          ctx.setLineDash([]);
          ctx.fillStyle = col;
          ctx.font = '7px Segoe UI, sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(lbl, bxDn, ch - padB - 4);
        }
        // Upwind boundary from crest
        const bxUp = toX(det.crestDist + dist);
        if(bxUp >= padL && bxUp <= cw - padR){
          ctx.setLineDash([2,4]);
          ctx.strokeStyle = col;
          ctx.lineWidth = 0.7;
          ctx.beginPath();
          ctx.moveTo(bxUp, padT);
          ctx.lineTo(bxUp, ch - padB);
          ctx.stroke();
          ctx.setLineDash([]);
          ctx.fillStyle = col;
          ctx.font = '7px Segoe UI, sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(lbl, bxUp, ch - padB - 4);
        }
      });
    }
  } else if(det.Mh > 1 && det.crestDist === 0){
    // Site IS the crest
    ctx.fillStyle = '#ff5252';
    ctx.font = 'bold 9px Segoe UI, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Crest (Site)', siteX, siteY - 20);
  }

  // ── Wind direction arrow ──
  const arrY = padT + 8;
  ctx.fillStyle = '#00d2ff';
  ctx.strokeStyle = '#00d2ff';
  ctx.lineWidth = 1.5;
  const arrLeft = toX(0) - 30;
  const arrRight = toX(0) + 60;
  ctx.beginPath();
  ctx.moveTo(arrRight, arrY);
  ctx.lineTo(arrLeft, arrY);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(arrLeft, arrY);
  ctx.lineTo(arrLeft + 8, arrY - 4);
  ctx.lineTo(arrLeft + 8, arrY + 4);
  ctx.closePath();
  ctx.fill();
  ctx.font = '8px Segoe UI, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('Wind', (arrLeft+arrRight)/2, arrY - 7);

  // ── Elevation sample dots ──
  pts.forEach(p => {
    if(p.x < xMin || p.x > xMax || p.x === 0) return;
    const sx = toX(p.x), sy = toY(p.y);
    const diff = p.y - siteElev;
    let col;
    if(diff < -20) col = '#22aa44';
    else if(diff < -5) col = '#66bb66';
    else if(diff <= 5) col = '#cccc44';
    else if(diff <= 20) col = '#ee8833';
    else col = '#dd3333';
    ctx.fillStyle = col;
    ctx.beginPath();
    ctx.arc(sx, sy, 2.5, 0, Math.PI*2);
    ctx.fill();
  });

  // ── Legend (top-right corner) ──
  const legW = 110, legH = 50;
  const legX = cw - padR - legW - 6, legY = padT + 4;
  ctx.fillStyle = 'rgba(13,21,38,.88)';
  ctx.strokeStyle = 'rgba(46,63,95,.6)';
  ctx.lineWidth = 1;
  ctx.beginPath();
  if(ctx.roundRect) ctx.roundRect(legX - 4, legY - 2, legW + 8, legH, 4);
  else ctx.rect(legX - 4, legY - 2, legW + 8, legH);
  ctx.fill();
  ctx.stroke();
  ctx.font = '8px Segoe UI, sans-serif';
  const legItems = [
    ['#2196F3', '◆ Site'],
    ['#ff5252', '▲ Crest'],
    ['#4CAF50', '— Lu (half-height dist)'],
    ['#ffab00', '| H (hill height)'],
    ['#e040fb', '— x (site to crest)']
  ];
  legItems.forEach((item, i) => {
    ctx.fillStyle = item[0];
    ctx.textAlign = 'left';
    ctx.fillText(item[1], legX, legY + 8 + i * 9);
  });

  // ── Stats overlay box (right side) ──
  const topoType = det.isEsc ? 'ESCARPMENT' : (det.H > 1 ? 'HILL' : 'FLAT');
  const sectorOf = si => Math.round(si / 4) % 8;
  const sectorIdx = sectorOf(subIdx);
  const mtVal = (S.Mt && S.Mt[sectorIdx]) || 1;

  const boxW = 140, boxH = 100;
  const boxX = cw - padR - boxW - 6;
  const boxY = ch - padB - boxH - 6;
  ctx.fillStyle = 'rgba(13,21,38,.9)';
  ctx.strokeStyle = 'rgba(0,210,255,.25)';
  ctx.lineWidth = 1;
  ctx.beginPath();
  if(ctx.roundRect) ctx.roundRect(boxX, boxY, boxW, boxH, 4);
  else ctx.rect(boxX, boxY, boxW, boxH);
  ctx.fill();
  ctx.stroke();

  ctx.font = 'bold 9px Segoe UI, sans-serif';
  ctx.fillStyle = '#00d2ff';
  ctx.textAlign = 'left';
  ctx.fillText('TOPOGRAPHY:', boxX + 8, boxY + 14);
  ctx.fillStyle = '#e8eaf6';
  ctx.font = 'bold 9px Segoe UI, sans-serif';
  ctx.fillText(topoType, boxX + 82, boxY + 14);

  const statsLines = [
    ['H:', det.H.toFixed(2) + ' m'],
    ['Lu:', det.Lu.toFixed(2) + ' m'],
    ['UPWIND SLOPE:', det.slope.toFixed(2)],
    ['x:', det.x.toFixed(1) + ' m'],
    ['Mh:', det.Mh.toFixed(4)],
    ['Mt:', mtVal.toFixed(4)]
  ];
  ctx.font = '8px Segoe UI, sans-serif';
  statsLines.forEach((line, i) => {
    const ly = boxY + 26 + i * 12;
    ctx.fillStyle = '#90a4ae';
    ctx.textAlign = 'left';
    ctx.fillText(line[0], boxX + 8, ly);
    ctx.fillStyle = i === 5 ? (mtVal > 1.05 ? '#ffab00' : '#00e676') : '#e8eaf6';
    ctx.font = i === 5 ? 'bold 9px Segoe UI, sans-serif' : '8px Segoe UI, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(line[1], boxX + boxW - 8, ly);
    ctx.font = '8px Segoe UI, sans-serif';
  });
}

function niceStep(range, targetTicks){
  const raw = range / targetTicks;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  if(norm <= 1.5) return mag;
  if(norm <= 3) return 2*mag;
  if(norm <= 7) return 5*mag;
  return 10*mag;
}

// ═════════════ LOCK / UNLOCK ANALYSIS ═════════════
function syncAnalysisLockButtons(){
  const locked = S.analysisLocked;
  ['btn-lock','btn-lock-site'].forEach(id=>{
    const b = document.getElementById(id);
    if(!b) return;
    b.textContent = locked ? '🔒 Locked' : '🔓 Unlocked';
    b.classList.toggle('locked', locked);
  });
  const latEl = document.getElementById('inp-lat');
  const lngEl = document.getElementById('inp-lng');
  if(latEl) latEl.readOnly = locked;
  if(lngEl) lngEl.readOnly = locked;
  if(locked && S.lockedSiteLat != null && S.lockedSiteLng != null){
    if(latEl) latEl.value = S.lockedSiteLat.toFixed(6);
    if(lngEl) lngEl.value = S.lockedSiteLng.toFixed(6);
  }
  if(leafletMap && leafletMap.getContainer){
    leafletMap.getContainer().classList.toggle('site-pin-locked', locked);
  }
}

function toggleAnalysisLock(){
  S.analysisLocked = !S.analysisLocked;
  if(S.analysisLocked){
    S.lockedResults = JSON.parse(JSON.stringify(S.R));
    S.lockedSiteLat = S.lat;
    S.lockedSiteLng = S.lng;
    syncAnalysisLockButtons();
    toast('Analysis locked — results and site pin frozen');
  } else {
    S.lockedResults = null;
    S.lockedSiteLat = null;
    S.lockedSiteLng = null;
    syncAnalysisLockButtons();
    calc();
    rebuild();
    refreshDirectionalWindUI();
    recalcPressures();
    toast('Analysis unlocked — live updates restored');
  }
}

// ═════════════ LIGHTS ═════════════
function addLights(){
  scene.add(new THREE.HemisphereLight(0x87ceeb,0x362d1b,0.5));
  scene.add(new THREE.AmbientLight(0x404060,0.3));
  const d = new THREE.DirectionalLight(0xfff5e6,1);
  d.position.set(35,45,25);
  d.castShadow=true;
  d.shadow.mapSize.width=d.shadow.mapSize.height=2048;
  d.shadow.camera.near=0.5; d.shadow.camera.far=160;
  d.shadow.camera.left=d.shadow.camera.bottom=-55;
  d.shadow.camera.right=d.shadow.camera.top=55;
  d.shadow.bias=-0.001;
  scene.add(d);
  const f=new THREE.DirectionalLight(0x8888cc,0.3);
  f.position.set(-25,18,-12);
  scene.add(f);
}
function setSceneBg(){
  if(!scene)return;
  const c=S.darkMode?0x0a0a1a:0xdbe9f4;
  scene.background=new THREE.Color(c);
  const useUploadFog = uploadModelCamExtent > 0 && uploadedModelGroup && uploadedModelVisible && !parametricVisible;
  if(useUploadFog){
    const ex = uploadModelCamExtent;
    scene.fog = new THREE.Fog(c, Math.min(85, Math.max(12, ex * 1.4)), Math.max(420, ex * 20));
  } else {
    scene.fog=new THREE.Fog(c,90,220);
  }
}
function addGround(){
  const g=new THREE.Mesh(new THREE.PlaneGeometry(220,220),new THREE.MeshStandardMaterial({color:0x2d5a27,roughness:.9}));
  g.rotation.x=-Math.PI/2; g.receiveShadow=true; g.name='ground';
  scene.add(g);
}

// ═══════════════════════════════════════════════
//   LEAFLET MAP
// ═══════════════════════════════════════════════
/**
 * Call when the user moves the pin or hits Detect again: abort in-flight map requests and
 * invalidate any work still running so the UI never has to "wait for analysis" before the next move.
 */
function invalidateInFlightSiteDetect(){
  siteDetectGeneration++;
  if(activeDetectAbort){
    try{ activeDetectAbort.abort(); }catch(e){}
    activeDetectAbort = null;
  }
  autoDetectQueued = false;
  autoDetectInFlight = false;
}

/**
 * Leaflet can report 0×0 or stale size if the map initializes before the layout finishes.
 * Without invalidateSize, map click / latlng are wrong until something forces a layout (e.g. overlay fitBounds).
 */
function ensureLeafletMapSizedForInteraction(){
  if(!leafletMap) return;
  try{
    const z = leafletMap.getSize();
    if(!z || z.x < 4 || z.y < 4) leafletMap.invalidateSize();
  }catch(e){}
}

/** Keep Leaflet in sync whenever #map-container gets a real size (first paint, flex, font load, window resize). */
function setupLeafletMapResizeObserver(){
  const host = document.getElementById('map-container');
  if(!host || typeof ResizeObserver === 'undefined' || host._cwResizeObs) return;
  host._cwResizeObs = true;
  const ro = new ResizeObserver(()=>{
    if(leafletMap) leafletMap.invalidateSize();
  });
  ro.observe(host);
}

/**
 * Cold load: run the same Overpass + elevation pipeline as "Detect" / map overlay buttons so
 * TC/Ms/Mt terrain data is populated immediately (user can move the pin without waiting on a manual trigger).
 */
function scheduleInitialTerrainAutoDetectFromLoad(){
  if(S.analysisLocked) return;
  if(!S.lat || !S.lng) return;
  try{
    syncSiteCoordsFromMapPin();
  }catch(e){}
  autoDetectAllMultipliers();
}

/** Keep the blue site marker on S.lat/S.lng before TC/Ms/Mt draws (avoids overlays centered on an old position). */
function alignMapMarkerWithSiteState(){
  if(!leafletMap || !mapMarker) return;
  try{
    const ll = mapMarker.getLatLng();
    if(!ll || !Number.isFinite(ll.lat) || !Number.isFinite(ll.lng)) return;
    if(Math.abs(ll.lat - S.lat) > 1e-6 || Math.abs(ll.lng - S.lng) > 1e-6){
      mapMarker.setLatLng([S.lat, S.lng]);
    }
  } catch(e){}
}

/**
 * After pin coordinates are committed and updateMapBuilding() has run, run calc / region / auto-detect.
 * Deferred with double requestAnimationFrame so the browser paints the new pin immediately on click/drag
 * (otherwise heavy sync calc + detect start blocks the main thread and the pin appears stuck until analysis progresses).
 */
function scheduleSiteFollowUpAfterPinMove(){
  requestAnimationFrame(()=>{
    requestAnimationFrame(()=>{
      if(S.analysisLocked) return;
      reverseGeocode(S.lat, S.lng);
      const detectedRegion = autoDetectRegion(S.lat, S.lng);
      if(detectedRegion){ applyRegion(detectedRegion); console.log('Region auto-detected:', detectedRegion); }
      checkLeeZone();
      calc();
      refreshDirectionalWindUI();
      autoDetectAllMultipliers({ fromPinMove: true });
    });
  });
}

/** Move site pin to lat/lng: same pipeline as clicking the map (terrain reset, detect, etc.). */
function applySiteLocationFromMap(lat, lng){
  if(S.analysisLocked){
    toast('Unlock analysis to move the site pin');
    return;
  }
  invalidateInFlightSiteDetect();
  deactivateMapOverlaysIfFreeTierExceeded();
  S.lat = lat;
  S.lng = lng;
  const la = document.getElementById('inp-lat');
  const ln = document.getElementById('inp-lng');
  if(la) la.value = S.lat.toFixed(6);
  if(ln) ln.value = S.lng.toFixed(6);
  try{ alignMapMarkerWithSiteState(); }catch(e){}
  resetSiteTerrainForNewPinLocation();
  detectPendingOsm = true;
  detectPendingElev = true;
  setTerrainDataStatus('loading');
  refreshDirectionalWindUI();
  toast('⏳ Sampling terrain…');
  requestAnimationFrame(()=>{
    updateMapBuilding();
    scheduleSiteFollowUpAfterPinMove();
  });
}

/**
 * Move site pin on map click. Uses Leaflet's `map.on('click')` and `e.latlng` so coordinates match
 * Leaflet's internal view state. The old DOM-capture + mouseEventToLatLng + stopImmediatePropagation
 * path broke first clicks until something else called invalidateSize (e.g. overlay buttons).
 */
function bindMapContainerClickToMoveSitePin(){
  if(!leafletMap) return;
  const c = leafletMap.getContainer();
  if(c._cwSitePinClickBound) return;
  c._cwSitePinClickBound = true;

  leafletMap.on('click', function(e){
    if(S.analysisLocked){
      toast('Unlock analysis to move the site pin');
      return;
    }
    const orig = e.originalEvent;
    const tgt = orig && orig.target;
    if(tgt && tgt.closest){
      if(tgt.closest('.leaflet-control')) return;
      if(tgt.closest('.leaflet-popup')) return;
      if(tgt.closest('.leaflet-tooltip')) return;
      if(tgt.closest('.cw-ms-manual-marker')) return;
    }
    ensureLeafletMapSizedForInteraction();
    const ll = e.latlng;
    if(!ll || !Number.isFinite(ll.lat) || !Number.isFinite(ll.lng)) return;
    if(typeof leafletMap.closePopup === 'function') leafletMap.closePopup();
    applySiteLocationFromMap(ll.lat, ll.lng);
  });
}

function initMap(){
  if(typeof L === 'undefined'){
    console.warn('Wind Analysis: Leaflet (L) not loaded — map disabled.');
    return;
  }
  const mapEl = document.getElementById('leaflet-map');
  if(!mapEl || leafletMap) return;

  mapStreetLayer = L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',{
    attribution:'&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/">CARTO</a>',
    maxZoom:20, subdomains:'abcd'
  });
  mapSatelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{
    attribution:'Esri', maxZoom:19
  });

  leafletMap = L.map('leaflet-map',{
    center:[S.lat, S.lng], zoom:17, zoomControl:true, layers:[mapStreetLayer]
  });

  setupLeafletMapResizeObserver();
  leafletMap.whenReady(function(){
    ensureLeafletMapSizedForInteraction();
    requestAnimationFrame(function(){ ensureLeafletMapSizedForInteraction(); });
  });

  bindMapContainerClickToMoveSitePin();

  leafletMap.on('mousemove', function(e){
    const el1 = document.getElementById('map-cur-lat');
    const el2 = document.getElementById('map-cur-lng');
    const el3 = document.getElementById('map-cur-dist');
    if(el1) el1.textContent = e.latlng.lat.toFixed(6);
    if(el2) el2.textContent = e.latlng.lng.toFixed(6);
    if(el3){
      const d = leafletMap.distance(e.latlng, L.latLng(S.lat, S.lng));
      el3.textContent = d < 1000 ? d.toFixed(0)+' m' : (d/1000).toFixed(2)+' km';
    }
  });

  // Right-click to add manual shielding buildings
  initMsContextMenu();

  updateMapBuilding();
}

function updateMapBuilding(){
  if(!leafletMap) return;
  deactivateMapOverlaysIfFreeTierExceeded();
  if(mapMarker) leafletMap.removeLayer(mapMarker);
  if(buildingPolygon) leafletMap.removeLayer(buildingPolygon);
  // Remove old overlays
  leafletMap.eachLayer(l=>{ if(l._isCW) leafletMap.removeLayer(l); });

  const w=S.width, d=S.depth;
  const angle = S.mapBuildingAngle * Math.PI / 180;
  const mPerDegLat = 111320;
  const mPerDegLng = 111320 * Math.cos(S.lat * Math.PI / 180);

  const corners = [
    [-w/2,-d/2],[w/2,-d/2],[w/2,d/2],[-w/2,d/2]
  ].map(([x,z])=>{
    const rx = x*Math.cos(angle) - z*Math.sin(angle);
    const rz = x*Math.sin(angle) + z*Math.cos(angle);
    return [S.lat + rz/mPerDegLat, S.lng + rx/mPerDegLng];
  });

  const qz = S.R.qz || 1;
  const intensity = Math.min(qz/3, 1);
  const r = Math.round(100 + intensity*155);
  const g = Math.round(100*(1-intensity));
  const b = Math.round(200*(1-intensity));

  buildingPolygon = L.polygon(corners,{
    interactive:false,
    color:`rgb(${r},${g},${b})`, fillColor:`rgb(${r},${g},${b})`,
    fillOpacity:0.5, weight:2, dashArray:'5,5'
  }).addTo(leafletMap);
  buildingPolygon._isCW = true;

  // Ridge
  if(S.roofType==='gable'||S.roofType==='hip'){
    const ridge = [[-w/2,0],[w/2,0]].map(([x,z])=>{
      const rx=x*Math.cos(angle)-z*Math.sin(angle);
      const rz=x*Math.sin(angle)+z*Math.cos(angle);
      return [S.lat+rz/mPerDegLat, S.lng+rx/mPerDegLng];
    });
    const rl = L.polyline(ridge,{interactive:false,color:'#ffaa00',weight:2,dashArray:'3,6'}).addTo(leafletMap);
    rl._isCW = true;
  }

  // interactive:false — clicks pass through to the map so one click moves the pin (no drag required)
  const sitePinIcon = L.divIcon({
    className: 'cw-site-pin-map-icon',
    html: '<div style="width:14px;height:14px;background:#e53935;border:2px solid #fff;border-radius:50%;transform:translate(-50%,-50%);box-shadow:0 0 6px rgba(229,57,53,.85)"></div>',
    iconAnchor: [7, 7]
  });
  mapMarker = L.marker([S.lat, S.lng], { interactive: false, icon: sitePinIcon, zIndexOffset: 600 }).addTo(leafletMap);
  mapMarker._isCW = true;
  mapMarker.bindPopup(`
    <div style="font-family:Segoe UI,sans-serif;font-size:12px;line-height:1.6">
      <b style="color:#1e40af">Building Site</b><br>
      📐 ${S.width}m × ${S.depth}m × ${S.height}m<br>
      🏠 ${S.roofType} roof @ ${S.pitch}°<br>
      🌪️ V<sub>R</sub> = ${S.windSpeed} m/s | Region ${S.region}<br>
      ${S.R.qz ? '📊 q<sub>z</sub> = '+S.R.qz.toFixed(3)+' kPa' : ''}
    </div>`);

  // Orientation arrow — shows "front" (windward face) direction on map
  const frontMid = [(corners[2][0]+corners[3][0])/2,(corners[2][1]+corners[3][1])/2];
  const orientDist = Math.max(w,d)*0.4;
  const orientEnd = [
    S.lat + Math.cos(angle)*orientDist/mPerDegLat,
    S.lng + Math.sin(angle)*orientDist/mPerDegLng
  ];
  const orientArrow = L.polyline([frontMid, orientEnd],{
    interactive:false,color:'#00e676',weight:3,opacity:0.9
  }).addTo(leafletMap);
  orientArrow._isCW = true;
  const orientHead = L.circleMarker(orientEnd,{interactive:false,radius:5,color:'#00e676',fillColor:'#00e676',fillOpacity:1,weight:0}).addTo(leafletMap);
  orientHead._isCW = true;
  // "Front" label — offset perpendicular so it clears the site pin when front faces the marker
  const perpA = angle + Math.PI / 2;
  const sideBumpM = Math.min(18, Math.max(w, d) * 0.22);
  const frontLabelPos = [
    orientEnd[0] + Math.cos(perpA) * sideBumpM / mPerDegLat,
    orientEnd[1] + Math.sin(perpA) * sideBumpM / mPerDegLng
  ];
  const frontIcon = L.divIcon({className:'',
    html:`<div style="transform:translate(-50%,-50%);background:rgba(0,230,118,.85);color:#000;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:bold;white-space:nowrap">FRONT ${S.mapBuildingAngle}°</div>`});
  const frontLabel = L.marker(frontLabelPos,{icon:frontIcon,interactive:false}).addTo(leafletMap);
  frontLabel._isCW = true;

  // Wind arrow
  drawMapWindArrow();
  // Dimension labels
  drawMapDimensions(corners);

  // Refresh TC zones overlay if active
  if(S.overlayTC) drawTCZones();
  // Refresh Ms shielding overlay if active
  if(S.overlayMs) drawMsOverlay();
  // Refresh Mt topographic overlay if active
  if(S.overlayMt) drawMtOverlay();
}

function drawMapWindArrow(){
  if(!leafletMap) return;
  leafletMap.eachLayer(l=>{ if(l._isWindArrow) leafletMap.removeLayer(l); });

  const ang = (S.windAngle)*Math.PI/180;
  const mPerDegLat = 111320;
  const mPerDegLng = 111320 * Math.cos(S.lat*Math.PI/180);
  const dist = Math.max(S.width, S.depth)*1.5;

  const startLat = S.lat + Math.cos(ang)*dist/mPerDegLat;
  const startLng = S.lng + Math.sin(ang)*dist/mPerDegLng;

  const arrow = L.polyline([[startLat,startLng],[S.lat,S.lng]],
    {color:'#ff3333',weight:3,opacity:0.8,interactive:false}).addTo(leafletMap);
  arrow._isWindArrow = true;

  const headSize = dist*0.2;
  const ba1 = ang+Math.PI+0.4, ba2 = ang+Math.PI-0.4;
  const head = L.polygon([
    [S.lat,S.lng],
    [S.lat+Math.cos(ba1)*headSize/mPerDegLat, S.lng+Math.sin(ba1)*headSize/mPerDegLng],
    [S.lat+Math.cos(ba2)*headSize/mPerDegLat, S.lng+Math.sin(ba2)*headSize/mPerDegLng]
  ],{color:'#ff3333',fillColor:'#ff3333',fillOpacity:0.8,weight:1,interactive:false}).addTo(leafletMap);
  head._isWindArrow = true;

  // Place wind label upwind (~70% from site toward arrow start) so it clears the pin / FRONT label
  const lLat = S.lat + (startLat - S.lat) * 0.72;
  const lLng = S.lng + (startLng - S.lng) * 0.72;
  const windLabel = L.divIcon({className:'',
    html:`<div style="background:rgba(255,50,50,.85);color:#fff;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:bold;white-space:nowrap;transform:translate(-50%,-50%)">
      WIND ${S.windSpeed}m/s @ ${Math.round(S.windAngle)}°</div>`});
  const lbl = L.marker([lLat,lLng],{icon:windLabel,interactive:false}).addTo(leafletMap);
  lbl._isWindArrow = true;
}

function drawMapDimensions(corners){
  if(!leafletMap) return;
  leafletMap.eachLayer(l=>{ if(l._isDimLabel) leafletMap.removeLayer(l); });

  const midFront = [(corners[0][0]+corners[1][0])/2,(corners[0][1]+corners[1][1])/2];
  const wIcon = L.divIcon({className:'',
    html:`<div style="background:rgba(255,170,68,.9);color:#000;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:bold">${S.width}m</div>`});
  const wl = L.marker(midFront,{icon:wIcon,interactive:false}).addTo(leafletMap);
  wl._isDimLabel = true;

  const midRight = [(corners[1][0]+corners[2][0])/2,(corners[1][1]+corners[2][1])/2];
  const dIcon = L.divIcon({className:'',
    html:`<div style="background:rgba(68,170,255,.9);color:#000;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:bold">${S.depth}m</div>`});
  const dl = L.marker(midRight,{icon:dIcon,interactive:false}).addTo(leafletMap);
  dl._isDimLabel = true;
}

function reverseGeocode(lat,lng){
  fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}&zoom=16`)
    .then(r=>r.json()).then(data=>{
      if(data.display_name) S.address = data.display_name.split(',').slice(0,3).join(', ');
    }).catch(()=>{});
}

function onLocationInput(){
  if(S.analysisLocked && S.lockedSiteLat != null && S.lockedSiteLng != null){
    const latEl = document.getElementById('inp-lat');
    const lngEl = document.getElementById('inp-lng');
    if(latEl) latEl.value = S.lockedSiteLat.toFixed(6);
    if(lngEl) lngEl.value = S.lockedSiteLng.toFixed(6);
    toast('Unlock analysis to change site location');
    return;
  }
  S.lat = parseFloat(document.getElementById('inp-lat').value)||0;
  S.lng = parseFloat(document.getElementById('inp-lng').value)||0;
  resetSiteTerrainForNewPinLocation();
  if(leafletMap){
    leafletMap.setView([S.lat, S.lng], 17);
    updateMapBuilding();
  }
  scheduleSiteFollowUpAfterPinMove();
}

function searchLocation(){
  if(S.analysisLocked){
    toast('Unlock analysis to change site location');
    return;
  }
  const q = prompt('Enter address or place name:');
  if(!q) return;
  fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(q)}&limit=1`)
    .then(r=>r.json()).then(data=>{
      if(data.length){
        S.lat = parseFloat(data[0].lat);
        S.lng = parseFloat(data[0].lon);
        document.getElementById('inp-lat').value = S.lat.toFixed(6);
        document.getElementById('inp-lng').value = S.lng.toFixed(6);
        S.address = data[0].display_name.split(',').slice(0,3).join(', ');
        resetSiteTerrainForNewPinLocation();
        if(leafletMap){ leafletMap.setView([S.lat,S.lng],17); updateMapBuilding(); }
        scheduleSiteFollowUpAfterPinMove();
        toast('Location: '+S.address);
      } else toast('Location not found');
    }).catch(()=>toast('Search failed'));
}

function clearLocation(){
  if(S.analysisLocked){
    toast('Unlock analysis to change site location');
    return;
  }
  S.lat = 0; S.lng = 0;
  document.getElementById('inp-lat').value = '0';
  document.getElementById('inp-lng').value = '0';
  if(leafletMap) updateMapBuilding();
}

function toggleMapSatellite(){
  if(!leafletMap) return;
  if(mapCurrentLayer==='street'){
    leafletMap.removeLayer(mapStreetLayer);
    leafletMap.addLayer(mapSatelliteLayer);
    mapCurrentLayer='satellite';
    toast('Satellite view');
  } else {
    leafletMap.removeLayer(mapSatelliteLayer);
    leafletMap.addLayer(mapStreetLayer);
    mapCurrentLayer='street';
    toast('Street view');
  }
}

function centerMapOnBuilding(){
  if(leafletMap) leafletMap.setView([S.lat,S.lng],18);
}

function onBuildingAngleChange(val){
  S.mapBuildingAngle = parseInt(val)||0;
  const lbl = document.getElementById('building-angle-label');
  if(lbl) lbl.textContent = S.mapBuildingAngle + '°';
  updateMapBuilding();
  rebuildScene();
}

function toggleMapOverlay(type){
  if(type==='tc'){
    if(!mapOverlayPressAllowed('tc')) return;
    S.overlayTC = !S.overlayTC;
    const btn = document.getElementById('btn-ov-tc');
    if(btn) btn.classList.toggle('active', S.overlayTC);
    incrementMapOverlayPressCount('tc');
    if(S.overlayTC){
      // Same OSM + landuse pipeline as Ms — without Detect, TC_dir / terrainZones are empty and rings have nothing meaningful to show.
      if(!S.terrainRecalcCtx && !autoDetectInFlight){
        syncSiteCoordsFromMapPin();
        autoDetectAllMultipliers();
      }
      drawTCZones({ fitView: true });
    } else clearTCZones();
    return;
  }
  if(type==='ms'){
    if(!mapOverlayPressAllowed('ms')) return;
    S.overlayMs = !S.overlayMs;
    const btn = document.getElementById('btn-ov-ms');
    if(btn) btn.classList.toggle('active', S.overlayMs);
    const legend = document.getElementById('map-ms-legend');
    if(legend) legend.style.display = S.overlayMs ? 'block' : 'none';
    incrementMapOverlayPressCount('ms');
    if(S.overlayMs){
      // Ms overlay needs OSM building data from Overpass — same pipeline as Detect. Without this, list is empty and nothing draws.
      if(!S.terrainRecalcCtx && !autoDetectInFlight){
        syncSiteCoordsFromMapPin();
        autoDetectAllMultipliers();
      }
      drawMsOverlay({ fitView: true });
    } else clearMsOverlay();
    return;
  }
  if(type==='mt'){
    if(!mapOverlayPressAllowed('mt')) return;
    S.overlayMt = !S.overlayMt;
    const btn = document.getElementById('btn-ov-mt');
    if(btn) btn.classList.toggle('active', S.overlayMt);
    const legend = document.getElementById('map-mt-legend');
    if(legend) legend.style.display = S.overlayMt ? 'block' : 'none';
    incrementMapOverlayPressCount('mt');
    if(S.overlayMt){
      const noElev = !S.detectedElevations || S.detectedElevations.length === 0;
      if(noElev && !autoDetectInFlight){
        syncSiteCoordsFromMapPin();
        autoDetectAllMultipliers();
      }
      drawMtOverlay({ fitView: true });
    } else clearMtOverlay();
    return;
  }
  toast('Map overlay: '+type.toUpperCase()+' (coming soon)');
}

function clearTCZones(){
  tcOverlayLayers.forEach(l=>{ if(leafletMap) leafletMap.removeLayer(l); });
  tcOverlayLayers = [];
}

function drawTCZones(opts){
  opts = opts || {};
  const fitView = opts.fitView === true;
  clearTCZones();
  if(!leafletMap) return;
  alignMapMarkerWithSiteState();
  if(S.terrainRecalcCtx){
    const c = S.terrainRecalcCtx;
    if(Math.abs(c.lat - S.lat) > 1e-4 || Math.abs(c.lng - S.lng) > 1e-4){
      S.terrainZones = [];
    }
  }

  const lat=S.lat, lng=S.lng;
  const mPerDegLat=111320;
  const mPerDegLng=111320*Math.cos(lat*Math.PI/180);
  const h = S.R ? S.R.h : S.height;
  const zoneRadius = Math.max(500, 40 * h);

  function bPt(r, bearing){
    const bRad=bearing*Math.PI/180;
    return [lat + r*Math.cos(bRad)/mPerDegLat, lng + r*Math.sin(bRad)/mPerDegLng];
  }

  const bearings=[0,45,90,135,180,225,270,315];
  const dirNames=['N','NE','E','SE','S','SW','W','NW'];
  // Traffic-light style: TC2 green, TC3 orange, TC4 red; TC1 / TC2.5 distinct but subordinate
  const tcBorder={1:'#64748b',2:'#22c55e',2.5:'#eab308',3:'#f97316',4:'#ef4444'};
  function tcColor(tc){ return tcBorder[tc] || tcBorder[Math.round(tc)] || '#f97316'; }

  // Draw concentric ring guides
  const ringDistances = [125, 250, 375, 500];
  if(zoneRadius > 500) ringDistances.push(zoneRadius);
  ringDistances.forEach(r=>{
    const pts=[];
    for(let a=0;a<=360;a+=3) pts.push(bPt(r,a));
    const rl=L.polyline(pts,{color:'#6666cc',weight:1,opacity:0.35,dashArray:'4,4',interactive:false}).addTo(leafletMap);
    tcOverlayLayers.push(rl);
  });

  // Draw sector dividers
  for(let i=0;i<8;i++){
    const bEdge=bearings[i]-22.5;
    const ep=bPt(zoneRadius, bEdge);
    const sl=L.polyline([[lat,lng],ep],{color:'#6666cc',weight:1.5,opacity:0.5,interactive:false}).addTo(leafletMap);
    tcOverlayLayers.push(sl);
  }

  // Draw per-direction zones (multi-zone if terrain zones detected)
  for(let i=0;i<8;i++){
    const zones = (S.terrainZones && S.terrainZones[i] && S.terrainZones[i].length > 0)
      ? S.terrainZones[i] : [{ from: 0, to: zoneRadius, tc: S.TC_dir[i] }];
    const startB=bearings[i]-22.5;
    const endB=bearings[i]+22.5;

    // Merge consecutive zones with same TC for cleaner display
    const merged = [];
    for(const z of zones){
      if(merged.length > 0 && merged[merged.length-1].tc === z.tc){
        merged[merged.length-1].to = z.to;
      } else {
        merged.push({ from: z.from, to: z.to, tc: z.tc });
      }
    }
    // If merged zones don't extend to zoneRadius, add final zone
    if(merged.length > 0 && merged[merged.length-1].to < zoneRadius){
      merged.push({ from: merged[merged.length-1].to, to: zoneRadius, tc: merged[merged.length-1].tc });
    }

    // Draw each zone as a sector annulus
    merged.forEach((z, zIdx)=>{
      const pts=[];
      // Inner arc
      for(let a=startB;a<=endB;a+=2) pts.push(bPt(z.from,a));
      pts.push(bPt(z.from,endB));
      // Outer arc (reversed)
      for(let a=endB;a>=startB;a-=2) pts.push(bPt(z.to,a));
      pts.push(bPt(z.to,startB));

      const col=tcColor(z.tc);
      const sector=L.polygon(pts,{
        color:col, weight:2, opacity:0.6,
        fillColor:col, fillOpacity:0.12,
        interactive:false
      }).addTo(leafletMap);
      tcOverlayLayers.push(sector);

      // Zone label — centroid of this sector-band wedge: radial mid of [from,to], angular mid of 45° sector
      const bandW = Math.max(z.to - z.from, 1e-6);
      let labelR = (z.from + z.to) / 2;
      const inset = Math.min(bandW * 0.08, 4);
      labelR = Math.max(z.from + inset, Math.min(z.to - inset, labelR));
      if(!(labelR > z.from && labelR < z.to)) labelR = z.from + bandW / 2;
      const labelAngle = bearings[i];
      const lp = bPt(labelR, labelAngle);
      const tcDisplay = (z.tc == null || !Number.isFinite(Number(z.tc))) ? '—' : (Number.isInteger(z.tc) ? z.tc : z.tc.toFixed(1));
      const tipText = dirNames[i]+': TC '+tcDisplay+' → '+z.to.toFixed(0)+' m';
      // Invisible point at wedge centroid; permanent tooltip with direction:center so text is
      // actually centered on lat/lng (divIcon+iconAnchor[0,0]+transform drifted labels to sector edges).
      const lbl = L.circleMarker(lp,{
        radius: 1, stroke: false, fillOpacity: 0, interactive: false, pane: 'markerPane'
      }).addTo(leafletMap);
      lbl.bindTooltip(tipText,{
        permanent: true,
        direction: 'center',
        className: 'tc-zone-map-label',
        opacity: 1,
        offset: [0, 0]
      });
      tcOverlayLayers.push(lbl);
    });
  }

  if(fitView){
    const outerPts=bearings.map(b=>bPt(zoneRadius*1.1,b));
    try{ leafletMap.fitBounds(L.latLngBounds(outerPts),{padding:[20,20],maxZoom:17}); }catch(e){}
  }
}

// ═══════════════════════════════════════════════
//   Ms SHIELDING OVERLAY (Wind Rose on Map)
// ═══════════════════════════════════════════════
function clearMsOverlay(){
  msOverlayLayers.forEach(l=>{ if(leafletMap) leafletMap.removeLayer(l); });
  msOverlayLayers = [];
  msManualLayers.forEach(l=>{ if(leafletMap) leafletMap.removeLayer(l); });
  msManualLayers = [];
  const hintEl = document.getElementById('map-ms-shield-hint');
  if(hintEl){ hintEl.style.display = 'none'; hintEl.innerHTML = ''; }
}

/** @param {{ fitView?: boolean }} [opts] — fitView: zoom map to 20h (only when turning overlay on; omit on refresh to avoid lag) */
function drawMsOverlay(opts){
  opts = opts || {};
  const fitView = opts.fitView === true;
  clearMsOverlay();
  if(!leafletMap) return;
  alignMapMarkerWithSiteState();
  if(S.terrainRecalcCtx){
    const c = S.terrainRecalcCtx;
    if(Math.abs(c.lat - S.lat) > 1e-4 || Math.abs(c.lng - S.lng) > 1e-4){
      S.terrainZones = [];
    }
  }

  const lat=S.lat, lng=S.lng;
  // Mean roof height — must match calc() formula exactly
  const ridgeRise = (S.pitch||0) > 0 && S.roofType !== 'flat' ? Math.tan((S.pitch||0)*Math.PI/180) * (S.depth||10) / 2 : 0;
  const h = (S.height||6) + (S.parapet||0) + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);
  const mPerDegLat=111320;
  const mPerDegLng=111320*Math.cos(lat*Math.PI/180);
  const shieldDist = Math.min(20*h, 500);
  const maxRadius = shieldDist;

  function bPt(r, bearing){
    const bRad=bearing*Math.PI/180;
    return [lat + r*Math.cos(bRad)/mPerDegLat, lng + r*Math.sin(bRad)/mPerDegLng];
  }

  const bearings=[0,45,90,135,180,225,270,315];
  const dirNames=['N','NE','E','SE','S','SW','W','NW'];

  // ── Concentric distance rings (white, dashed) ──
  const ringDistances = [50,100,200,300,400,500].filter(d=> d<=maxRadius);
  ringDistances.forEach(r=>{
    const pts=[];
    for(let a=0;a<=360;a+=2) pts.push(bPt(r,a));
    const rl=L.polyline(pts,{color:'#ffffff',weight:1,opacity:0.5,dashArray:'6,4',interactive:false}).addTo(leafletMap);
    msOverlayLayers.push(rl);
    const labelPt = bPt(r, 90);
    const distIcon = L.divIcon({className:'',
      html:`<div style="transform:translate(-50%,-50%);color:#333;font-size:9px;font-weight:600;pointer-events:none;background:rgba(255,255,255,.8);padding:1px 4px;border-radius:3px">${r} m</div>`,
      iconAnchor:[0,0]});
    const distLbl = L.marker(labelPt,{icon:distIcon,interactive:false}).addTo(leafletMap);
    msOverlayLayers.push(distLbl);
  });

  // ── 8 radial sector boundary lines (white, solid) ──
  for(let i=0;i<8;i++){
    const bEdge = bearings[i] - 22.5;
    const ep = bPt(maxRadius, bEdge);
    const sl = L.polyline([[lat,lng],ep],{color:'#ffffff',weight:1.5,opacity:0.7,interactive:false}).addTo(leafletMap);
    msOverlayLayers.push(sl);
  }

  // ── Center pin (blue, Checkwind-style) ──
  const centerIcon = L.divIcon({className:'',
    html:'<div style="width:12px;height:12px;background:#2196F3;border:2px solid #fff;border-radius:50%;transform:translate(-50%,-50%);box-shadow:0 0 6px rgba(33,150,243,.8)"></div>',
    iconAnchor:[6,6]});
  const centerMarker = L.marker([lat,lng],{icon:centerIcon,interactive:false,zIndexOffset:1000}).addTo(leafletMap);
  msOverlayLayers.push(centerMarker);

  // ── Building footprint polygons (green=shielding h_s≥h, red=too short, grey=out of range) ──
  const buildings = S.detectedBuildingsList || [];
  buildings.forEach((b, idx)=>{
    if(!b.footprint || b.footprint.length < 3) return;
    const isEffective = b.height >= h;
    const inRange = b.distance <= shieldDist && b.distance > 5;
    // Only draw footprints inside 5m…20h (Ms search annulus). Skip grey “outside” buildings — huge Leaflet speed-up in cities.
    if(!inRange) return;
    let color, fillColor, fillOpacity;

    if(isEffective){
      color='#4CAF50'; fillColor='#4CAF50'; fillOpacity=0.35;
    } else {
      color='#F44336'; fillColor='#F44336'; fillOpacity=0.25;
    }

    const poly = L.polygon(b.footprint, {
      color, weight:1.5, opacity:0.9,
      fillColor, fillOpacity
    }).addTo(leafletMap);

    // Click popup with per-building properties (Checkwind-style)
    const popupHtml = `
      <div style="font-family:Arial,sans-serif;font-size:12px;min-width:200px">
        <div style="font-weight:bold;color:#333;border-bottom:2px solid ${isEffective?'#4CAF50':'#F44336'};padding-bottom:4px;margin-bottom:6px">
          SHIELDING STRUCTURE #${String(idx+1).padStart(3,'0')}
        </div>
        <table style="width:100%;font-size:11px;border-collapse:collapse">
          <tr><td style="color:#666;padding:2px 8px 2px 0">Height (h<sub>s</sub>)</td><td style="font-weight:600">${b.height.toFixed(1)} m</td></tr>
          <tr><td style="color:#666;padding:2px 8px 2px 0">h<sub>s</sub> ≥ h</td><td style="font-weight:600;color:${isEffective?'#4CAF50':'#F44336'}">${isEffective?'Yes':'No'} (h=${h.toFixed(1)}m)</td></tr>
          <tr><td style="color:#666;padding:2px 8px 2px 0">Distance</td><td style="font-weight:600">${b.distance.toFixed(1)} m</td></tr>
          <tr><td style="color:#666;padding:2px 8px 2px 0">Bearing</td><td style="font-weight:600">${b.bearing.toFixed(1)}° (${dirNames[b.sectorIdx]})</td></tr>
          <tr><td style="color:#666;padding:2px 8px 2px 0">Area</td><td style="font-weight:600">${b.area.toFixed(0)} m²</td></tr>
          <tr><td style="color:#666;padding:2px 8px 2px 0">Breadth b<sub>s</sub> (${dirNames[b.sectorIdx]})</td><td style="font-weight:600">${b.breadth.toFixed(1)} m</td></tr>
          ${b.elevation!==null ? `<tr><td style="color:#666;padding:2px 8px 2px 0">Elevation</td><td style="font-weight:600">${b.elevation.toFixed(2)} m</td></tr>` : ''}
          ${Math.abs(b.slope)>0.001 ? `<tr><td style="color:#666;padding:2px 8px 2px 0">Slope</td><td style="font-weight:600">${b.slope.toFixed(4)}</td></tr>` : ''}
        </table>
      </div>`;
    poly.bindPopup(popupHtml);
    msOverlayLayers.push(poly);
  });

  // ── Per-sector Ms labels with shielding details ──
  for(let i=0;i<8;i++){
    const msVal = S.Ms[i];
    const lp = bPt(maxRadius * 0.75, bearings[i]);
    const sd = S.shieldingDetails ? S.shieldingDetails[i] : null;
    // Count detected & qualifying buildings in this sector
    const sectorBldgs = buildings.filter(b=>b.sectorIdx===i && b.distance<=shieldDist && b.distance>5);
    const qualCount = sectorBldgs.filter(b=>b.height>=h).length;
    let detailLine;
    if(sd && sd.ns > 0){
      detailLine = `n<sub>s</sub>=${sd.ns}/${sectorBldgs.length}, s=${sd.s.toFixed(1)}`;
    } else {
      detailLine = `${qualCount}/${sectorBldgs.length} qualify`;
    }

    const icon = L.divIcon({className:'',
      html:`<div style="transform:translate(-50%,-50%);background:#fff;color:#333;padding:4px 10px;border-radius:4px;font-size:11px;font-weight:bold;text-align:center;white-space:nowrap;border:1px solid #ccc;line-height:1.5;pointer-events:none;box-shadow:0 1px 4px rgba(0,0,0,.15)">
        ${dirNames[i]}:<br>Ms = ${msVal.toFixed(msVal===1?1:2)}<br>
        <span style="font-size:9px;font-weight:normal;color:#888">${detailLine}</span></div>`,
      iconAnchor:[0,0]});
    const lbl = L.marker(lp,{icon:icon,interactive:false}).addTo(leafletMap);
    msOverlayLayers.push(lbl);
  }

  // ── Show mean roof height, 20h radius, and manual building instructions ──
  const manualCount = S.manualShieldBuildings.length;
  const clearBtn = manualCount > 0 ? ` · <a href="#" onclick="clearManualShieldBuildings();return false" style="color:#ff6b6b;text-decoration:underline;pointer-events:auto">Clear ${manualCount} manual</a>` : '';
  const bList = S.detectedBuildingsList;
  const nRaw = (bList && bList.length) || 0;
  let drawableInRing = 0;
  for(const b of bList || []){
    if(!b.footprint || b.footprint.length < 3) continue;
    if(b.distance <= shieldDist && b.distance > 5) drawableInRing++;
  }
  let hintExtra = '';
  if(!S.terrainRecalcCtx){
    hintExtra = ` <span style="font-weight:normal;color:#9ee">· Loading…</span>`;
  } else if(nRaw === 0){
    hintExtra = `<br><span style="font-weight:normal;color:#ffb74d;font-size:9px">No OSM buildings in radius — move pin, <b>Detect</b>, or right‑click.</span>`;
  } else if(drawableInRing === 0){
    hintExtra = `<br><span style="font-weight:normal;color:#ffb74d;font-size:9px">Buildings outside 5–${shieldDist.toFixed(0)} m ring or no footprints.</span>`;
  }
  const hintEl = document.getElementById('map-ms-shield-hint');
  if(hintEl){
    hintEl.innerHTML = `<div class="map-ms-shield-hint-card" title="Right‑click map: add manual shielding. Green dot on manual building: remove. While loading, you can still click the map to move the pin.">
      h=${h.toFixed(1)}m · 20h=${shieldDist.toFixed(0)}m${clearBtn}<br>
      <span class="map-ms-shield-hint-sub">Right‑click · green dot removes</span>${hintExtra}</div>`;
    hintEl.style.display = 'block';
  }

  // Draw existing manual shielding buildings
  drawManualShieldMarkers();

  if(fitView){
    const outerPts = bearings.map(b=>bPt(maxRadius*1.15,b));
    try{ leafletMap.fitBounds(L.latLngBounds(outerPts),{padding:[30,30],maxZoom:17}); }catch(e){}
  }
}

// ── Manual shielding building support ──
// Right-click on map to add a shielding building when Ms overlay is active
function initMsContextMenu(){
  if(!leafletMap) return;
  leafletMap.on('contextmenu', onMsRightClick);
}

function onMsRightClick(e){
  if(!S.overlayMs) return; // only when Ms overlay is active
  e.originalEvent.preventDefault();

  const lat = S.lat, lng = S.lng;
  const mPerDegLat = 111320;
  const mPerDegLng = 111320 * Math.cos(lat * Math.PI / 180);
  // Mean roof height — must match calc() formula exactly
  const ridgeRise = (S.pitch||0) > 0 && S.roofType !== 'flat' ? Math.tan((S.pitch||0)*Math.PI/180) * (S.depth||10) / 2 : 0;
  const h = (S.height||6) + (S.parapet||0) + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);
  const shieldDist = Math.min(20 * h, 500);

  const clickLat = e.latlng.lat;
  const clickLng = e.latlng.lng;
  const dLat = clickLat - lat;
  const dLng = clickLng - lng;
  const dist = Math.sqrt((dLat * mPerDegLat) ** 2 + (dLng * mPerDegLng) ** 2);

  if(dist > shieldDist){
    toast('⚠ Building outside 20h shielding radius');
    return;
  }
  if(dist < 5){
    toast('⚠ Too close to subject building');
    return;
  }

  // Add manual building with assumed height = h (qualifying) and 10m breadth
  S.manualShieldBuildings.push({
    lat: clickLat, lng: clickLng,
    height: Math.max(h, 7),  // at least h so it qualifies
    breadth: 10  // reasonable default for NZ house
  });

  recalcMsFromAll();
  drawMsOverlay();
  toast(`✓ Manual shielding building added (${S.manualShieldBuildings.length} total)`);
}

function clearManualShieldBuildings(){
  S.manualShieldBuildings = [];
  recalcMsFromAll();
  drawMsOverlay();
  toast('Manual shielding buildings cleared');
}

function drawManualShieldMarkers(){
  // Clear old manual markers
  msManualLayers.forEach(l => { if(leafletMap) leafletMap.removeLayer(l); });
  msManualLayers = [];

  if(!leafletMap || !S.manualShieldBuildings.length) return;

  S.manualShieldBuildings.forEach((b, idx) => {
    const icon = L.divIcon({
      className: '',
      html: `<div class="cw-ms-manual-marker" style="width:16px;height:16px;background:#4CAF50;border:2px solid #fff;border-radius:50%;transform:translate(-50%,-50%);box-shadow:0 0 4px rgba(0,0,0,.4);cursor:pointer" title="Manual building #${idx+1}"></div>`,
      iconAnchor: [8, 8]
    });
    const marker = L.marker([b.lat, b.lng], { icon, interactive: true }).addTo(leafletMap);
    marker.on('click', () => {
      // Click to remove
      S.manualShieldBuildings.splice(idx, 1);
      recalcMsFromAll();
      drawMsOverlay();
      toast(`Manual building removed (${S.manualShieldBuildings.length} remaining)`);
    });
    msManualLayers.push(marker);
  });
}

// Recalculate Ms from combined OSM + manual buildings
function recalcMsFromAll(){
  const lat = S.lat, lng = S.lng;
  const mPerDegLat = 111320;
  const mPerDegLng = 111320 * Math.cos(lat * Math.PI / 180);
  // Mean roof height — must match calc() formula exactly
  const ridgeRise = (S.pitch||0) > 0 && S.roofType !== 'flat' ? Math.tan((S.pitch||0)*Math.PI/180) * (S.depth||10) / 2 : 0;
  const h = (S.height||6) + (S.parapet||0) + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);
  const shieldDist = Math.min(20 * h, 500);
  const dirNames = ['N','NE','E','SE','S','SW','W','NW'];

  // Build combined list: OSM detected + manual
  const osmBuildings = S.detectedBuildingsList || [];

  // Convert manual buildings to same format
  const manualConverted = S.manualShieldBuildings.map(b => {
    const dLat = b.lat - lat;
    const dLng = b.lng - lng;
    let bearing = Math.atan2(dLng, dLat) * 180 / Math.PI;
    bearing = ((bearing % 360) + 360) % 360;
    const idx = Math.round(bearing / 45) % 8;
    const dist = Math.sqrt((dLat * mPerDegLat) ** 2 + (dLng * mPerDegLng) ** 2);
    return {
      height: b.height, breadth: b.breadth,
      distance: dist, bearing, sectorIdx: idx,
      manual: true
    };
  });

  const allBuildings = [...osmBuildings, ...manualConverted];

  // Table 4.2 interpolation
  const msTable = [[1.5, 0.7], [3.0, 0.8], [6.0, 0.9], [12.0, 1.0]];
  function tableLookupMs(sParam){
    if(sParam <= msTable[0][0]) return msTable[0][1];
    if(sParam >= msTable[msTable.length - 1][0]) return msTable[msTable.length - 1][1];
    for(let j = 0; j < msTable.length - 1; j++){
      if(sParam <= msTable[j + 1][0]){
        const t = (sParam - msTable[j][0]) / (msTable[j + 1][0] - msTable[j][0]);
        return msTable[j][1] + t * (msTable[j + 1][1] - msTable[j][1]);
      }
    }
    return 1.0;
  }

  S.shieldingDetails = [];
  if(h > 25){
    for(let i = 0; i < 8; i++){
      S.Ms[i] = 1.0;
      S.shieldingDetails[i] = { ns: 0, ls: 0, hs: 0, bs: 0, s: Infinity, Ms: 1.0, reason: 'h > 25m' };
    }
  } else {
    for(let i = 0; i < 8; i++){
      const sectorBldgs = allBuildings.filter(b => b.sectorIdx === i && b.distance <= shieldDist && b.distance > 5);
      const qualifying = sectorBldgs.filter(b => b.height >= h);
      const ns = qualifying.length;

      if(ns === 0){
        S.Ms[i] = 1.0;
        S.shieldingDetails[i] = { ns: 0, ls: 0, hs: 0, bs: 0, s: Infinity, Ms: 1.0 };
        continue;
      }

      const hs = qualifying.reduce((sum, b) => sum + b.height, 0) / ns;
      const bs = qualifying.reduce((sum, b) => sum + b.breadth, 0) / ns;
      const ls = h * (10 / ns + 5);
      const sParam = ls / Math.sqrt(Math.max(hs * bs, 0.01));
      const msVal = parseFloat(tableLookupMs(sParam).toFixed(2));
      S.Ms[i] = msVal;
      S.shieldingDetails[i] = { ns, ls: parseFloat(ls.toFixed(1)), hs: parseFloat(hs.toFixed(1)), bs: parseFloat(bs.toFixed(1)), s: parseFloat(sParam.toFixed(2)), Ms: msVal };
    }
  }

  // Recalc downstream
  calc();
  recalcPressures();
  refreshDirectionalWindUI();
}

// ═══════════════════════════════════════════════
//   Mt TOPOGRAPHIC OVERLAY
// ═══════════════════════════════════════════════
function clearMtOverlay(){
  mtOverlayLayers.forEach(l=>{ if(leafletMap) leafletMap.removeLayer(l); });
  mtOverlayLayers = [];
}

/** @param {{ fitView?: boolean }} [opts] — fitView: zoom map to overlay extent (use when user opens overlay; omit on refresh) */
function drawMtOverlay(opts){
  opts = opts || {};
  const fitView = opts.fitView === true;
  clearMtOverlay();
  if(!leafletMap) return;
  alignMapMarkerWithSiteState();

  const lat=S.lat, lng=S.lng;
  const mPerDegLat=111320;
  const mPerDegLng=111320*Math.cos(lat*Math.PI/180);
  const maxRadius = 2500; // display radius in metres

  function bPt(r, bearing){
    const bRad=bearing*Math.PI/180;
    return [lat + r*Math.cos(bRad)/mPerDegLat, lng + r*Math.sin(bRad)/mPerDegLng];
  }

  const bearings8=[0,45,90,135,180,225,270,315];
  const dir8Names=['North','North East','East','South East','South','South West','West','North West'];
  // 32-point compass names (every 11.25°)
  const dir32Names=['N','NbE','NNE','NEbN','NE','NEbE','ENE','EbN',
    'E','EbS','ESE','SEbE','SE','SEbS','SSE','SbE',
    'S','SbW','SSW','SWbS','SW','SWbW','WSW','WbS',
    'W','WbN','WNW','NWbW','NW','NWbN','NNW','NbW'];
  const dir32Full=['North','North by East','North North East','North East by North',
    'North East','North East by East','East North East','East by North',
    'East','East by South','East South East','South East by East',
    'South East','South East by South','South South East','South by East',
    'South','South by West','South South West','South West by South',
    'South West','South West by West','West South West','West by South',
    'West','West by North','West North West','North West by West',
    'North West','North West by North','North North West','North by West'];

  const nSub = S.nSubDirs || 32;
  const sampleDists = S.detectedSampleDistances || [];
  const elevs       = S.detectedElevations || [];
  const siteElev    = S.detectedSiteElev || 0;
  const mhSub       = S.mhSub || [];
  const nDist       = sampleDists.length;
  const stepDeg     = 360 / nSub;

  // ── Red dashed radial lines ──
  for(let si=0;si<nSub;si++){
    const bearing = si*stepDeg;
    const ep = bPt(maxRadius, bearing);
    const wt = (si%4===0) ? 1.5 : (si%2===0) ? 1.2 : 0.8;
    const op = (si%4===0) ? 0.7 : 0.4;
    const rl = L.polyline([[lat,lng],ep],{color:'#cc3333',weight:wt,opacity:op,dashArray:'8,6',interactive:false}).addTo(leafletMap);
    mtOverlayLayers.push(rl);
  }

  // ── Distance ring ──
  const ringR = Math.min(maxRadius, 1500);
  const ringPts=[];
  for(let a=0;a<=360;a+=2) ringPts.push(bPt(ringR,a));
  const rl=L.polyline(ringPts,{color:'#cc3333',weight:1.5,opacity:0.6,dashArray:'8,6',interactive:false}).addTo(leafletMap);
  mtOverlayLayers.push(rl);
  const distPt = bPt(ringR, 12);
  const distIcon = L.divIcon({className:'',
    html:`<div style="transform:translate(-50%,-50%);background:rgba(255,255,255,.85);color:#333;font-size:10px;font-weight:600;padding:2px 6px;border:1px solid #cc3333;border-radius:3px;pointer-events:none;white-space:nowrap">${ringR.toFixed(0)} m from site</div>`,
    iconAnchor:[0,0]});
  mtOverlayLayers.push(L.marker(distPt,{icon:distIcon,interactive:false}).addTo(leafletMap));

  // ── Coloured elevation sample points per direction ──
  // Use full adaptive profiles if available, otherwise fall back to flat grid
  const adaptiveProfiles = S.detectedProfiles || null;
  if(adaptiveProfiles && adaptiveProfiles.length >= nSub){
    for(let si=0;si<nSub;si++){
      const bearing = si*stepDeg;
      const prof = adaptiveProfiles[si];
      for(let k=1; k<prof.length; k++){  // skip k=0 (site itself)
        const dist = prof[k].dist;
        if(dist > maxRadius*1.1) continue;
        const elev = prof[k].elev;
        const diff = elev - siteElev;
        let col;
        if(diff < -20)      col = '#22aa44';
        else if(diff < -5)  col = '#66bb66';
        else if(diff < 5)   col = '#cccc44';
        else if(diff < 20)  col = '#ee8833';
        else                 col = '#dd3333';
        const pt = bPt(dist, bearing);
        const sz = 5;
        const dotIcon = L.divIcon({className:'',
          html:`<div style="width:${sz}px;height:${sz}px;background:${col};border:1px solid #fff;border-radius:50%;transform:translate(-50%,-50%);box-shadow:0 0 3px rgba(0,0,0,.3)"></div>`,
          iconAnchor:[sz/2,sz/2]});
        mtOverlayLayers.push(L.marker(pt,{icon:dotIcon,interactive:false}).addTo(leafletMap));
      }
    }
  } else if(elevs.length > 0 && nDist > 0){
    for(let si=0;si<nSub;si++){
      const bearing = si*stepDeg;
      const baseIdx = 1 + si*nDist;
      for(let d=0; d<nDist; d++){
        const dist = sampleDists[d];
        if(dist > maxRadius*1.1) continue;
        const elev = elevs[baseIdx+d];
        const diff = elev - siteElev;
        // Colour by elevation difference: green=lower, yellow=similar, red=higher
        let col;
        if(diff < -20)      col = '#22aa44';
        else if(diff < -5)  col = '#66bb66';
        else if(diff < 5)   col = '#cccc44';
        else if(diff < 20)  col = '#ee8833';
        else                 col = '#dd3333';
        const pt = bPt(dist, bearing);
        const sz = 6;
        const dotIcon = L.divIcon({className:'',
          html:`<div style="width:${sz}px;height:${sz}px;background:${col};border:1px solid #fff;border-radius:50%;transform:translate(-50%,-50%);box-shadow:0 0 3px rgba(0,0,0,.3)"></div>`,
          iconAnchor:[sz/2,sz/2]});
        mtOverlayLayers.push(L.marker(pt,{icon:dotIcon,interactive:false}).addTo(leafletMap));
      }
    }
  }

  // ── Centre marker (blue pin) ──
  const centerIcon = L.divIcon({className:'',
    html:'<div style="width:12px;height:12px;background:#4488ff;border:2px solid #fff;border-radius:50%;transform:translate(-50%,-50%);box-shadow:0 0 6px rgba(68,136,255,.6)"></div>',
    iconAnchor:[6,6]});
  mtOverlayLayers.push(L.marker([lat,lng],{icon:centerIcon,interactive:false}).addTo(leafletMap));

  // ── 8 Sector Mt labels (white boxes at ~40% radius) ──
  for(let i=0;i<8;i++){
    const mtVal = S.Mt[i];
    const lp = bPt(maxRadius*0.38, bearings8[i]);
    const borderCol = mtVal > 1.05 ? '#cc6600' : '#4488cc';
    const sIcon = L.divIcon({className:'',
      html:`<div style="transform:translate(-50%,-50%);background:rgba(255,255,255,.92);color:#333;padding:3px 8px;border-radius:3px;font-size:10px;font-weight:600;text-align:center;white-space:nowrap;border:1px solid ${borderCol};line-height:1.4;pointer-events:none;box-shadow:0 1px 4px rgba(0,0,0,.2)">
        ${dir8Names[i]} Sector:<br>M<sub>t</sub> = ${mtVal.toFixed(4)} <span style="font-size:8px;font-weight:500">(incl. lee)</span></div>`,
      iconAnchor:[0,0]});
    mtOverlayLayers.push(L.marker(lp,{icon:sIcon,interactive:false}).addTo(leafletMap));
  }

  // ── Sub-direction Mh labels (at outer edge) ──
  if(mhSub.length >= nSub){
    for(let si=0;si<nSub;si++){
      const mhVal = mhSub[si];
      const bearing = si*stepDeg;
      const pt = bPt(maxRadius*0.72, bearing);
      const dirName = (nSub===32) ? dir32Full[si] : (dir32Full[si*2] || ('Dir '+si));
      const mhIcon = L.divIcon({className:'',
        html:`<div style="transform:translate(-50%,-50%);text-align:center;pointer-events:none">
          <div style="background:rgba(255,255,255,.90);color:#333;font-size:9px;font-weight:600;padding:2px 6px;border-radius:3px;white-space:nowrap;border:1px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.2)">
            ${dirName}:<br>Mh = ${mhVal.toFixed(4)}</div>
        </div>`,
        iconAnchor:[0,0]});
      mtOverlayLayers.push(L.marker(pt,{icon:mhIcon,interactive:false}).addTo(leafletMap));
    }
  }

  if(fitView){
    const outerPts = bearings8.map(b=>bPt(maxRadius*1.1,b));
    try{ leafletMap.fitBounds(L.latLngBounds(outerPts),{padding:[30,30],maxZoom:16}); }catch(e){}
  }
}

// ═══════════════════════════════════════════════
//   WIND REGIONS & SPEED LOOKUP
// ═══════════════════════════════════════════════
// AS/NZS 1170.2:2021 Table 3.2(A) — Wind direction multipliers M_d — Australia
// AS/NZS 1170.2:2021 Table 3.2(B) — Wind direction multipliers M_d — New Zealand
// Direction order: [N, NE, E, SE, S, SW, W, NW]
const MD_TABLE = {
  // Australia
  'A0':[0.90,0.85,0.85,0.90,0.90,0.95,1.00,0.95],
  'A1':[0.90,0.85,0.85,0.80,0.80,0.95,1.00,0.95],
  'A2':[0.85,0.75,0.85,0.95,0.95,0.95,1.00,0.95],
  'A3':[0.90,0.75,0.75,0.90,0.90,0.95,1.00,0.95],
  'A4':[0.85,0.75,0.75,0.80,0.90,0.95,1.00,0.95],
  'A5':[0.95,0.80,0.80,0.80,0.90,0.95,1.00,0.95],
  'B1':[0.75,0.75,0.85,0.90,0.90,0.95,1.00,0.90],
  'B2':[0.90,0.90,0.90,0.90,0.90,0.90,0.90,0.90],
  'C': [0.90,0.90,0.90,0.90,0.90,0.90,0.90,0.90],
  'D': [0.90,0.90,0.90,0.90,0.90,0.90,0.90,0.90],
  // New Zealand
  'NZ1':[0.90,0.95,0.95,0.95,0.90,1.00,1.00,0.95],
  'NZ2':[0.95,0.90,0.80,0.90,0.95,1.00,1.00,1.00],
  'NZ3':[1.00,0.75,0.75,0.85,0.95,0.95,0.90,1.00],
  'NZ4':[0.95,0.75,0.75,0.75,0.85,0.95,1.00,1.00],
};

// AS/NZS 1170.2:2021 Table 3.1(A) — Regional wind speeds V_R (m/s) — Australia
// AS/NZS 1170.2:2021 Table 3.1(B) — Regional wind speeds V_R (m/s) — New Zealand
const REGION_VR = {
  // Australia — Non-cyclonic
  'A0':{100:41,200:43,500:45,1000:46},
  'A1':{100:41,200:43,500:45,1000:46},
  'A2':{100:41,200:43,500:45,1000:46},
  'A3':{100:41,200:43,500:45,1000:46},
  'A4':{100:41,200:43,500:45,1000:46},
  'A5':{100:41,200:43,500:45,1000:46},
  'B1':{100:48,200:52,500:57,1000:60},
  'B2':{100:48,200:52,500:57,1000:60},
  // Australia — Cyclonic (maximum values)
  'C' :{100:56,200:61,500:66,1000:70},
  'D' :{100:66,200:72,500:80,1000:85},
  // New Zealand
  'NZ1':{100:42,200:43,500:45,1000:46},
  'NZ2':{100:42,200:43,500:45,1000:46},
  'NZ3':{100:50,200:51,500:53,1000:54},
  'NZ4':{100:47,200:48,500:50,1000:50}
};

// AS/NZS 1170.0 — Table 3.3 annual probability of exceedance for wind (ULS) → return period R = 1/P (years).
// Rows: design working life. IL4 @ 100+ years: * (site-specific); R = 2500 used as default lower bound.
const WIND_ULS_RETURN_PERIOD = {
  25:  { 1: 50,   2: 250,  3: 500,  4: 1000 },
  50:  { 1: 100,  2: 500,  3: 1000, 4: 2500 },
  100: { 1: 250,  2: 1000, 3: 2500, 4: 2500 },
};

function windUlsReturnPeriodYears(designLifeYears, importanceLevel){
  const life = [25, 50, 100].includes(designLifeYears) ? designLifeYears : 50;
  const il = Math.min(4, Math.max(1, parseInt(importanceLevel, 10) || 2));
  const row = WIND_ULS_RETURN_PERIOD[life] || WIND_ULS_RETURN_PERIOD[50];
  return row[il] || 500;
}

/** Interpolate/extrapolate V_R (m/s) from Table 3.1 tabulated recurrence intervals in REGION_VR. */
function regionalWindSpeedFromReturnPeriod(region, R){
  const rv = REGION_VR[region];
  if(!rv || !Number.isFinite(R)) return 45;
  R = Math.max(1, R);
  if(rv[R] != null) return rv[R];
  const keys = Object.keys(rv).map(Number).sort((a, b) => a - b);
  const kMin = keys[0], kMax = keys[keys.length - 1];
  if(R <= kMin){
    const k1 = keys[1] != null ? keys[1] : kMin;
    const v0 = rv[kMin], v1 = rv[k1];
    const slope = (v1 - v0) / (k1 - kMin);
    return Math.max(0, v0 + slope * (R - kMin));
  }
  if(R >= kMax){
    const k0 = keys[keys.length - 2], k1 = keys[keys.length - 1];
    const v0 = rv[k0], v1 = rv[k1];
    const slope = (v1 - v0) / (k1 - k0);
    return Math.max(0, v1 + slope * (R - k1));
  }
  for(let i = 0; i < keys.length - 1; i++){
    if(R >= keys[i] && R <= keys[i + 1]){
      const k0 = keys[i], k1 = keys[i + 1];
      const t = (R - k0) / (k1 - k0);
      return rv[k0] + t * (rv[k1] - rv[k0]);
    }
  }
  return rv[kMax];
}

function updateDesignAriDisplay(){
  const el = document.getElementById('val-design-ari');
  if(!el) return;
  const R = S.ari;
  const P = R ? (1 / R) : 0;
  el.textContent = R
    ? `${R} yr  (P ≈ ${P.toExponential(2)} /yr, 1/${R})`
    : '—';
}

/** When false, V_R is taken from AS/NZS 1170.2 Table 3.1 using ARI from 1170.0 Table 3.3. */
function releaseVrManual(){
  S.vrManual = false;
}

function onWindSpeedManualEdit(){
  S.vrManual = true;
}

// ═══════════════════════════════════════════════
//   WIND REGION AUTO-DETECTION FROM LAT/LNG
// ═══════════════════════════════════════════════
// Point-in-polygon ray-casting test
function pointInPoly(lat,lng,poly){
  let inside=false;
  for(let i=0,j=poly.length-1;i<poly.length;j=i++){
    const yi=poly[i][0],xi=poly[i][1],yj=poly[j][0],xj=poly[j][1];
    if(((yi>lat)!==(yj>lat))&&(lng<(xj-xi)*(lat-yi)/(yj-yi)+xi)) inside=!inside;
  }
  return inside;
}

// New Zealand region polygons (from Figure 3.1B)
// NZ1: Upper North Island (north of New Plymouth – Rotorua line)
const NZ1_POLY = [
  [-34.0,172.0],[-34.0,179.0],[-38.0,179.0],[-38.4,177.5],
  [-38.7,177.0],[-38.9,176.6],[-39.1,174.0],[-38.5,174.0],
  [-38.0,174.5],[-37.8,174.0],[-36.5,173.5],[-34.5,172.0],[-34.0,172.0]
];
// NZ3: Eastern South Island (Marlborough/Canterbury/Otago east coast)
const NZ3_POLY = [
  [-41.3,174.0],[-41.3,175.0],[-41.5,175.5],[-41.7,175.0],
  [-42.0,174.5],[-42.5,173.5],[-43.0,172.5],[-43.5,172.0],
  [-44.0,171.5],[-44.5,170.8],[-45.0,170.5],[-45.5,170.0],
  [-46.0,169.5],[-46.5,169.0],[-46.5,172.0],[-46.0,172.5],
  [-45.0,172.0],[-44.0,173.0],[-43.0,174.0],[-42.0,175.0],
  [-41.3,174.0]
];
// NZ4: Far south South Island + offshore islands (Stewart Island, Chathams, Auckland Islands)
const NZ4_POLY = [
  [-46.0,165.0],[-46.0,180.0],[-56.0,180.0],[-56.0,165.0],[-46.0,165.0]
];
// NZ2: Lower North Island + western/southern South Island (everything NZ not NZ1/NZ3/NZ4)
// No explicit polygon needed — NZ2 is the default for NZ coordinates

// Australian region polygons (simplified from Figure 3.1A)
// Regions tested in order: D, C, B2, B1, then A0-A5 by geography
// D: Northwest WA coast (Carnarvon to Broome area, within ~50km of coast)
const AU_D_POLY = [
  [-15.0,122.0],[-15.0,129.0],[-20.0,129.0],[-20.0,119.0],
  [-24.0,113.0],[-25.5,113.0],[-25.5,114.5],[-23.5,114.5],
  [-20.5,118.5],[-19.0,121.0],[-16.0,122.5],[-15.0,122.0]
];
// C: Northern tropical coast (WA Kimberley, NT, QLD north)
const AU_C_POLY = [
  [-10.0,114.0],[-10.0,145.5],[-16.5,145.5],[-19.0,147.0],
  [-20.5,149.0],[-21.0,150.0],[-22.0,150.5],[-23.5,151.5],
  [-23.5,150.0],[-21.0,148.0],[-19.5,146.5],[-17.0,145.0],
  [-16.0,140.0],[-15.5,135.0],[-14.0,130.0],[-15.0,129.0],
  [-20.0,129.0],[-20.0,119.0],[-24.0,113.0],[-22.0,113.0],
  [-15.0,114.0],[-10.0,114.0]
];
// B2: Isolated coastal areas (Norfolk Island, Christmas Island, etc.)
// Simplified to key areas — Norfolk Island + parts of QLD coast
const AU_B2_POLY_NORFOLK = [
  [-28.5,167.5],[-29.5,167.5],[-29.5,168.5],[-28.5,168.5],[-28.5,167.5]
];
// B1: Southeast QLD coast (roughly Bundaberg to Brisbane)
const AU_B1_POLY = [
  [-23.5,151.5],[-23.5,153.5],[-28.5,154.5],[-28.5,153.0],
  [-27.0,153.5],[-24.5,152.0],[-23.5,151.5]
];
// A5: Southeast QLD / Northern NSW (inland/southern of B1)
const AU_A5_POLY = [
  [-25.0,148.0],[-25.0,153.5],[-28.5,154.5],[-33.0,152.5],
  [-33.0,148.0],[-25.0,148.0]
];
// A2: Greater Sydney / NSW coast (south of A5)
const AU_A2_POLY = [
  [-33.0,148.0],[-33.0,152.5],[-37.5,150.5],[-37.5,148.0],[-33.0,148.0]
];
// A1: Tasmania + southern Victoria
const AU_A1_POLY = [
  [-37.5,143.5],[-37.5,150.5],[-44.0,150.5],[-44.0,143.5],[-37.5,143.5]
];
// A4: Southwest WA coast (Perth and south)
const AU_A4_POLY = [
  [-30.0,114.5],[-30.0,119.0],[-36.0,119.0],[-36.0,114.5],[-30.0,114.5]
];
// A3: SA coast (Adelaide region + southern coast)
const AU_A3_POLY = [
  [-31.0,133.0],[-31.0,143.5],[-38.0,143.5],[-38.0,133.0],[-31.0,133.0]
];
// A0: Central/inland Australia (default for AU coordinates not matching others)

function autoDetectRegion(lat, lng){
  // ── New Zealand (lat roughly -34 to -53, lng roughly 165-180) ──
  if(lat < -33.5 && lat > -53 && lng > 165 && lng < 180){
    if(pointInPoly(lat,lng,NZ4_POLY)) return 'NZ4';
    if(pointInPoly(lat,lng,NZ1_POLY)) return 'NZ1';
    if(pointInPoly(lat,lng,NZ3_POLY)) return 'NZ3';
    return 'NZ2'; // default NZ region
  }
  // ── Chatham Islands (NZ4) ──
  if(lat < -43 && lat > -45 && lng > -177.5 && lng < -175.5) return 'NZ4';

  // ── Australia (lat roughly -10 to -44, lng roughly 113-155) ──
  if(lat < -9 && lat > -45 && lng > 112 && lng < 155){
    // Check cyclonic regions first (coast-specific, higher priority)
    if(pointInPoly(lat,lng,AU_D_POLY)) return 'D';
    if(pointInPoly(lat,lng,AU_C_POLY)) return 'C';
    if(pointInPoly(lat,lng,AU_B1_POLY)) return 'B1';
    // Non-cyclonic regions by geography
    if(pointInPoly(lat,lng,AU_A1_POLY)) return 'A1';
    if(pointInPoly(lat,lng,AU_A2_POLY)) return 'A2';
    if(pointInPoly(lat,lng,AU_A3_POLY)) return 'A3';
    if(pointInPoly(lat,lng,AU_A4_POLY)) return 'A4';
    if(pointInPoly(lat,lng,AU_A5_POLY)) return 'A5';
    return 'A0'; // Central/inland Australia default
  }
  // Norfolk Island (B2)
  if(pointInPoly(lat,lng,AU_B2_POLY_NORFOLK)) return 'B2';

  // Fallback — outside AU/NZ, return null (keep current region)
  return null;
}

function applyRegion(region){
  if(!region || !REGION_VR[region]) return;
  S.region = region;
  const dd = document.getElementById('inp-region');
  if(dd) dd.value = region;
  const rd = document.getElementById('region-display');
  if(rd) rd.textContent = region;
  if(!S.vrManual){
    const R = windUlsReturnPeriodYears(S.life, S.importance);
    S.ari = R;
    const vr = regionalWindSpeedFromReturnPeriod(region, R);
    S.windSpeed = vr;
    const wEl = document.getElementById('inp-windspeed');
    if(wEl) wEl.value = vr;
    const svcVr = Math.round(vr * 0.71);
    S.svcVr = svcVr;
    const svcEl = document.getElementById('inp-svc-vr');
    if(svcEl) svcEl.value = svcVr;
    updateDesignAriDisplay();
  }
  // Update Md per Table 3.2
  const mdVals = MD_TABLE[region] || [1,1,1,1,1,1,1,1];
  for(let i=0;i<8;i++) S.Md[i] = mdVals[i];
}

function onRegionChange(){
  const region = document.getElementById('inp-region').value;
  S.region = region;
  if(!S.vrManual){
    const R = windUlsReturnPeriodYears(S.life, S.importance);
    S.ari = R;
    const vr = regionalWindSpeedFromReturnPeriod(region, R);
    S.windSpeed = vr;
    const wEl = document.getElementById('inp-windspeed');
    if(wEl) wEl.value = vr;
    const svcVr = Math.round(vr * 0.71);
    S.svcVr = svcVr;
    const svcEl = document.getElementById('inp-svc-vr');
    if(svcEl) svcEl.value = svcVr;
    updateDesignAriDisplay();
  }
  // Update M_d per Table 3.2 for the selected region
  const mdVals = MD_TABLE[region] || [1,1,1,1,1,1,1,1];
  for(let i=0;i<8;i++) S.Md[i] = mdVals[i];
  onInput();
}

function syncWindSpeed(val){
  S.windSpeed = parseFloat(val);
  document.getElementById('inp-windspeed').value = val;
}

function syncImportanceSegButtons(){
  const h = document.getElementById('inp-importance');
  const wrap = document.getElementById('seg-importance');
  if(!h || !wrap) return;
  const v = String(h.value);
  wrap.querySelectorAll('.seg-btn').forEach(btn=>{
    const on = btn.getAttribute('data-il') === v;
    btn.classList.toggle('active', on);
    btn.setAttribute('aria-pressed', on);
  });
}

function syncDesignLifeSegButtons(){
  const h = document.getElementById('inp-life');
  const wrap = document.getElementById('seg-life');
  if(!h || !wrap) return;
  const v = String(h.value);
  wrap.querySelectorAll('.seg-btn').forEach(btn=>{
    const on = btn.getAttribute('data-life') === v;
    btn.classList.toggle('active', on);
    btn.setAttribute('aria-pressed', on);
  });
}

function setImportanceLevel(n){
  releaseVrManual();
  const il = Math.min(4, Math.max(1, parseInt(n, 10) || 2));
  const h = document.getElementById('inp-importance');
  if(h) h.value = String(il);
  syncImportanceSegButtons();
  onInput();
}

function setDesignLifeYears(years){
  releaseVrManual();
  const y = [25, 50, 100].includes(years) ? years : 50;
  const h = document.getElementById('inp-life');
  if(h) h.value = String(y);
  syncDesignLifeSegButtons();
  onInput();
}

function syncImportance(val){
  S.importance = parseInt(val, 10);
  const siteEl = document.getElementById('inp-importance');
  if(siteEl) siteEl.value = val;
  syncImportanceSegButtons();
  S.vrManual = false;
  applyRegion(S.region);
}

// ═══════════════════════════════════════════════
//   Kr — PARAPET REDUCTION FACTOR (Table 5.7)
// ═══════════════════════════════════════════════
// Kr reduces roof suction (negative Cpe) when parapets are present.
// For H ≤ 25m: ratio = hp/h; For H > 25m: ratio = hp/w
// hp = parapet height, h = reference height, w = min(width, depth)
function calcKr(hp, h, w){
  if(!hp || hp <= 0) return 1.0;
  let ratio, breakpoints;
  if(h <= 25){
    ratio = hp / h;
    // Table 5.7: hp/h thresholds
    breakpoints = [[0.07, 1.0], [0.10, 0.8], [0.20, 0.5]];
  } else {
    ratio = hp / w;
    // Table 5.7: hp/w thresholds
    breakpoints = [[0.02, 1.0], [0.03, 0.8], [0.05, 0.5]];
  }
  if(ratio <= breakpoints[0][0]) return 1.0;
  if(ratio >= breakpoints[2][0]) return 0.5;
  // Linear interpolation
  for(let i = 1; i < breakpoints.length; i++){
    if(ratio <= breakpoints[i][0]){
      const [r0, kr0] = breakpoints[i-1];
      const [r1, kr1] = breakpoints[i];
      return kr0 + (kr1 - kr0) * (ratio - r0) / (r1 - r0);
    }
  }
  return 0.5;
}

// ═══════════════════════════════════════════════
//   WIND LOAD CALCULATION  (AS/NZS 1170.2:2021)
// ═══════════════════════════════════════════════

/** Combined Mt per direction: topography (Mh) and lee (Mlee) per regional rules — used once in Vsit. */
function combineMtMlee(mh, mlee, region, siteElevM){
  const m = mlee || 1;
  const el = siteElevM || 0;
  if(region === 'A0'){
    return Math.max(0.5 + 0.5 * mh, 1);
  }
  if(['A4','NZ1','NZ2','NZ3','NZ4'].includes(region) && el > 500){
    return Math.max(mh * m * (1 + 0.00015 * el), 1);
  }
  return Math.max(Math.max(mh, m), 1);
}

/** Refresh S.Mt[i] from S.Mt_hill and S.Mlee (call after elevation, lee check, or manual Mlee edit). */
function recalcMtCombined(){
  const el = S.detectedSiteElev;
  const siteElev = (el !== undefined && el !== null && !isNaN(el)) ? el : 0;
  for(let i=0;i<8;i++){
    const mh = S.Mt_hill[i] || 1;
    S.Mt[i] = parseFloat(combineMtMlee(mh, S.Mlee[i]||1, S.region, siteElev).toFixed(4));
  }
}

/**
 * Design q_z with max-of-three-direction envelope — must match recalcPressures().
 * Requires S.Mzcat[i] set for all i (terrain multiplier M_z,cat per sector).
 * Tall windward walls (h > 25 m) use envelopeQzWindPressureAtZ(z) per height band instead.
 */
function envelopeQzWindPressure(VRdir, dirIdx){
  const Md = effectiveMd(dirIdx);
  const Mzcat = effectiveMzcat(dirIdx);
  const Ms = S.Ms[dirIdx];
  const Mt = S.Mt[dirIdx];
  const v0 = VRdir * Md * Mzcat * Ms * Mt;
  let qz = 0.5 * 1.2 * v0 * v0 / 1000;
  const adjL = (dirIdx + 7) % 8, adjR = (dirIdx + 1) % 8;
  const qzL = (()=>{ const v=VRdir*effectiveMd(adjL)*effectiveMzcat(adjL)*S.Ms[adjL]*S.Mt[adjL]; return 0.5*1.2*v*v/1000; })();
  const qzR = (()=>{ const v=VRdir*effectiveMd(adjR)*effectiveMzcat(adjR)*S.Ms[adjR]*S.Mt[adjR]; return 0.5*1.2*v*v/1000; })();
  return Math.max(qz, qzL, qzR);
}

/** M_z,cat at height z (m) for sector dirIndex — same TC resolution as calc() S.Mzcat. */
function mzAtHeightForDir(z, dirIndex){
  const tc = S.terrainCat;
  const tcDir = S.TC_dir[dirIndex];
  const tcForMz = (tcDir != null && Number.isFinite(Number(tcDir))) ? tcDir : tc;
  return mzCat(z, tcForMz);
}

/**
 * Design q_z at reference height z with max-of-three-direction envelope — mirrors envelopeQzWindPressure
 * but uses M_z,cat(z) per sector (for windward wall strips when h > 25 m).
 */
function envelopeQzWindPressureAtZ(VRdir, dirIdx, z){
  const Md = effectiveMd(dirIdx);
  const Ms = S.Ms[dirIdx];
  const Mt = S.Mt[dirIdx];
  const Mzcat = mzAtHeightForDir(z, dirIdx);
  const v0 = VRdir * Md * Mzcat * Ms * Mt;
  let qz = 0.5 * 1.2 * v0 * v0 / 1000;
  const adjL = (dirIdx + 7) % 8, adjR = (dirIdx + 1) % 8;
  const qzL = (()=>{ const v=VRdir*effectiveMd(adjL)*mzAtHeightForDir(z, adjL)*S.Ms[adjL]*S.Mt[adjL]; return 0.5*1.2*v*v/1000; })();
  const qzR = (()=>{ const v=VRdir*effectiveMd(adjR)*mzAtHeightForDir(z, adjR)*S.Ms[adjR]*S.Mt[adjR]; return 0.5*1.2*v*v/1000; })();
  return Math.max(qz, qzL, qzR);
}

/** Vertical bands aligned to Table 4.1 height rows (3,5,10,…) for windward q_z(z) integration. */
function windwardWallHeightBands(wallH){
  if(wallH <= 0) return [];
  const hs = [3,5,10,15,20,30,40,50,75,100,150,200];
  const edges = new Set([0, wallH]);
  for(const x of hs){
    if(x > 0 && x < wallH) edges.add(x);
  }
  const sorted = [...edges].sort((a,b)=>a-b);
  const bands = [];
  for(let i=0;i<sorted.length-1;i++){
    const zLo = sorted[i], zHi = sorted[i+1];
    if(zHi - zLo < 1e-9) continue;
    bands.push({zLo, zHi});
  }
  return bands;
}

/** Height-weighted mean q_z for windward (matches strip integration; used where one q_z per face is needed). */
function windwardHeightMeanQz(VRdir, dirIdx, wallH){
  const bands = windwardWallHeightBands(wallH);
  if(!bands.length) return envelopeQzWindPressure(VRdir, dirIdx);
  let s = 0, t = 0;
  for(const {zLo, zHi} of bands){
    const dz = zHi - zLo;
    const zMid = (zLo + zHi) / 2;
    s += envelopeQzWindPressureAtZ(VRdir, dirIdx, zMid) * dz;
    t += dz;
  }
  return t > 0 ? s / t : envelopeQzWindPressure(VRdir, dirIdx);
}

/** Net pressure case 1 — Table 5.5: p_net = p_e K_c,e − p_i (same as tabulated pressures). */
function pNetTable55Case1(pe, kce, pi1){
  return pe * kce - pi1;
}

/** User edits combined Mt in dir table → back-solve Mh for current Mlee/region. */
function combinedMtTargetToMhill(T, mlee, region, siteElevM){
  T = Math.max(1, T);
  const m = mlee || 1;
  const el = siteElevM || 0;
  if(region === 'A0'){
    return Math.max(1, Math.min(4, 2 * T - 1));
  }
  if(['A4','NZ1','NZ2','NZ3','NZ4'].includes(region) && el > 500){
    const k = 1 + 0.00015 * el;
    return Math.max(0.5, T / (m * k));
  }
  return T >= m ? T : T;
}

function calc(){
  if(!S.Mt_hill || S.Mt_hill.length !== 8){
    S.Mt_hill = (S.Mt && S.Mt.length === 8) ? S.Mt.slice() : [1,1,1,1,1,1,1,1];
  }
  recalcMtCombined();
  const activeLimitBtn = document.querySelector('.limitstate-btn.active');
  const isUlt = !activeLimitBtn || activeLimitBtn.dataset.limit === 'ultimate';
  const VR = isUlt ? S.windSpeed : S.svcVr;
  const VRdir = VR;
  const rt0 = S.roofType || 'gable';
  const geomW = S.width;
  const geomD = S.depth;
  const roofD = rt0 === 'monoslope' ? geomD : S.depth;
  // Reference height h = average roof height per Clause 4.2.1
  const ridgeRise = S.pitch>0 && S.roofType!=='flat' ? Math.tan(S.pitch*Math.PI/180)*S.depth/2 : 0;
  const h = S.height + S.parapet + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);
  const tc = S.terrainCat;
  const dirIdx = Math.round(((S.windAngle % 360 + 360) % 360) / 45) % 8;
  // Per-direction TC from map detect; until then use global terrain category (inp-terrain) so Mz,cat is not nulled by mzCat(h, null) → effectiveMzcat 1.0
  for(let i=0;i<8;i++){
    const tcDir = S.TC_dir[i];
    const tcForMz = (tcDir != null && Number.isFinite(Number(tcDir))) ? tcDir : tc;
    S.Mzcat[i] = mzCat(h, tcForMz);
  }
  const qz = envelopeQzWindPressure(VRdir, dirIdx);
  const Md = effectiveMd(dirIdx);
  const Mz = effectiveMzcat(dirIdx);
  const Ms = S.Ms[dirIdx];
  const Mt = S.Mt[dirIdx];
  const Mlee = S.Mlee[dirIdx] || 1;
  const Vsit = VRdir * Md * Mz * Ms * Mt;
  const impF = 1;

  let effW=geomW, effD=geomD, angleOff=0;
  if(S.loadCase==='B'){angleOff=90;effW=geomD;effD=geomW}
  else if(S.loadCase==='C') angleOff=45;

  // Swap effective dimensions when wind hits the depth face (left/right wall)
  const relAng = ((S.windAngle - S.mapBuildingAngle) % 360 + 360) % 360;
  const hitsSide = (relAng >= 45 && relAng < 135) || (relAng >= 225 && relAng < 315);
  if(hitsSide){ const tmp=effW; effW=effD; effD=tmp; }

  // h/d ratio for roof tables, d/b ratio for leeward table
  const rHD = h / effD;
  const db = effD / effW;
  // Table 5.2(A): Cp,e = 0.7 for h ≤ 25m (wind speed at z=h), 0.8 for h > 25m
  let CpWW = h > 25 ? 0.8 : 0.7;
  let CpLW = leewardCp(db, S.pitch);
  if(S.loadCase==='D') CpLW=-0.5;
  if(S.loadCase==='C'){CpWW*=0.85;CpLW*=1.1}

  // Roof pressure coefficients — Tables 5.3(A)/(B)/(C)
  // AS/NZS 1170.2: U/D for α < 10° (and monoslope/flat rules) → 5.3(A) strips; α ≥ 10° (hip/gable) → 5.3(B) U, 5.3(C) D.
  // Crosswind slope R: gable → 5.3(A) strips for all α; hip with α ≥ 10° → 5.3(C) R; monoslope crosswind etc. per §5.3.
  const alphaLt10 = S.pitch < 10;
  const useTableA_UD = alphaLt10 || rt0 === 'monoslope' || rt0 === 'flat';
  const useTableBC_UD = S.pitch >= 10 && rt0 !== 'monoslope' && rt0 !== 'flat';
  // Table 5.3(C): crosswind R for hip roofs only when α ≥ 10°. Gable crosswind R uses Table 5.3(A) at all α (per standard).
  const rCrosswindSlopeR_TableC = rt0 === 'hip' && S.pitch >= 10;
  let CpRW, CpRL_D, CpRL_R;
  if(useTableA_UD){
    // Table 5.3(A): distance-based upwind (α < 10°), or monoslope/flat (all α)
    if(rHD <= 0.5){ CpRW=-0.9; }
    else if(rHD <= 1.0){ CpRW=-0.9-(rHD-0.5)/0.5*0.4; } // -0.9 → -1.3
    else { CpRW=-1.3; }
    CpRL_D=-0.5;
    CpRL_R=-0.5;
  } else {
    CpRW = roofUpwindCp(S.pitch, rHD);
    CpRL_D = roofDownwindSlopeD(S.pitch, rHD);
    CpRL_R = roofCrosswindHipR(S.pitch, rHD);
  }

  const Aww=effW*S.height, Alw=effW*S.height, Asw=effD*S.height;
  const Aroof=(effW*effD/2)/(Math.cos(S.pitch*Math.PI/180)||1);

  // Clause 5.3.4 — Kv open area/volume factor
  const { vol: volInit } = prepareVolumeForKv();
  const aWWi = S.openWW / 100 * S.width * S.height;
  const aLWi = S.openLW / 100 * S.width * S.height;
  const aSWi = S.openSW / 100 * S.depth * S.height;
  const Ai = Math.max(aWWi, aLWi, aSWi);
  const kvAutoCalc = computeKvAuto(volInit, Ai);
  let KvInit = kvAutoCalc.Kv;

  const cpiResult = getCpiCasesForDesign();
  const cpi1 = cpiResult.cpi1;
  const cpi2 = cpiResult.cpi2;
  const kce1 = parseFloat(document.getElementById('kce1-val')?.value) || 0.8;
  const kci1 = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;

  // Kr parapet reduction factor (Table 5.7) — reduces roof suction
  const wMin = Math.min(effW, effD);
  const Kr = calcKr(S.parapet, h, wMin);

  const mk=(name,Cpe,area,clause)=>{
    const Ka = 1.0, Kp = 1.0;
    const pe = qz * Cpe * Ka * Kp;
    const pi1 = qz * cpi1 * kci1 * KvInit;
    const pi2 = qz * cpi2 * kci1 * KvInit;
    const p = pNetTable55Case1(pe, kce1, pi1);
    const pAlt = pNetTable55Case1(pe, kce1, pi2);
    return{
      name,Cp_e:Cpe,Cp_i:cpi1,Cp_i_alt:cpi2,p,p_case1:p,p_case2:pAlt,
      area,force:Math.abs(p*area),clause
    };
  };
  const mkWindward = ()=>{
    if(h <= 25) return mk('Windward Wall', CpWW, Aww,'Table 5.2(A)');
    const wallH = S.height;
    const bands = windwardWallHeightBands(wallH);
    if(!bands.length) return mk('Windward Wall', CpWW, Aww,'Table 5.2(A)');
    const Ka = 1.0, Kp = 1.0;
    let totalForce = 0;
    const zones = bands.map(({zLo, zHi})=>{
      const zMid = (zLo + zHi) / 2;
      const qzLocal = envelopeQzWindPressureAtZ(VRdir, dirIdx, zMid);
      const area = effW * (zHi - zLo);
      const pe = qzLocal * CpWW * Ka * Kp;
      const pi1 = qz * cpi1 * kci1 * KvInit;
      const pi2 = qz * cpi2 * kci1 * KvInit;
      const p = pNetTable55Case1(pe, kce1, pi1);
      const pAlt = pNetTable55Case1(pe, kce1, pi2);
      const force = Math.abs(p * area);
      totalForce += force;
      return {
        zLo, zHi, zMid, qz: qzLocal, Cpe: CpWW, area, p, p_case1:p, p_case2:pAlt, force,
        dist: zLo.toFixed(1)+'–'+zHi.toFixed(1)+' m AGL'
      };
    });
    const sumPA = zones.reduce((s,zn)=> s + zn.p * zn.area, 0);
    const sumPA2 = zones.reduce((s,zn)=> s + zn.p_case2 * zn.area, 0);
    const pAvg = Aww > 0 ? sumPA / Aww : 0;
    const pAvg2 = Aww > 0 ? sumPA2 / Aww : 0;
    const peExternalAvg = Aww > 0
      ? zones.reduce((s,zn)=> s + zn.qz * CpWW * zn.area, 0) / Aww
      : qz * CpWW;
    return{
      name: 'Windward Wall',
      Cp_e: CpWW,
      zones,
      Cp_i: cpi1,
      Cp_i_alt: cpi2,
      p: pAvg,
      p_case1: pAvg,
      p_case2: pAvg2,
      area: Aww,
      force: totalForce,
      peExternalAvg,
      clause: 'Table 5.2(A) — q_z varies with height (h > 25 m)'
    };
  };
  // Table 5.2(C) — sidewall Cp,e varies with distance from windward edge
  const mkSidewall=(name)=>{
    const Ka = 1.0;
    const KpUniform = getKpValue();
    const zones=sidewallCpZones(effD,h,S.height);
    let totalForce=0;
    const zoneData=zones.map(z=>{
      const Kp = (z.Cpe < 0) ? KpUniform : 1.0;
      const pe = qz * z.Cpe * Ka * Kp;
      const pi1 = qz * cpi1 * kci1 * KvInit;
      const pi2 = qz * cpi2 * kci1 * KvInit;
      const p = pNetTable55Case1(pe, kce1, pi1);
      const pAlt = pNetTable55Case1(pe, kce1, pi2);
      const force=Math.abs(p*z.area);
      totalForce+=force;
      return{dist:z.dist,Cpe:z.Cpe,area:z.area,width:z.width,p,p_case1:p,p_case2:pAlt,force};
    });
    const Asw=effD*S.height;
    const sumPA = zoneData.reduce((s,z)=> s + z.p * z.area, 0);
    const sumPA2 = zoneData.reduce((s,z)=> s + z.p_case2 * z.area, 0);
    const pAvg = Asw > 0 ? sumPA / Asw : 0;
    const pAvg2 = Asw > 0 ? sumPA2 / Asw : 0;
    return{
      name,Cp_e:-0.65,zones:zoneData,Cp_i:cpi1,Cp_i_alt:cpi2,p:pAvg,p_case1:pAvg,p_case2:pAvg2,
      area:Asw,force:totalForce,clause:'Table 5.2(C)'
    };
  };
  // Apply Kr to negative roof Cpe (suction reduction due to parapet)
  const CpRW_kr = CpRW < 0 ? CpRW * Kr : CpRW;
  const CpRL_kr_D = CpRL_D < 0 ? CpRL_D * Kr : CpRL_D;
  const CpRL_kr_R = CpRL_R < 0 ? CpRL_R * Kr : CpRL_R;
  const roofWWClause = useTableBC_UD ? 'Table 5.3(B)' : 'Table 5.3(A)';
  const roofLWClause = useTableBC_UD ? 'Table 5.3(C)' : 'Table 5.3(A)';
  const clauseRCrosswindC = 'Table 5.3(C) — crosswind slope R (α ≥ 10°)';
  // Monoslope/gable: ridge ∥ width (X). Wind on L/R → parallel to ridge → Fig 5.2 **R**. Wind on low eave (front) → **U**; wind on high wall (back) → **R** (Fig 5.2).
  const alongRidge = hitsSide;
  const ov = S.overhang || 0;
  const rw = effW + 2*ov;
  const clauseR = 'Table 5.3(A) Fig 5.2 R — crosswind slope R (strip)';
  const mkRoofWW = ()=>{
    // Wind ∥ ridge (gable/hip end): both slopes are R (crosswind), not U/D — Figure 5.2
    if(alongRidge){
      if(rCrosswindSlopeR_TableC){
        return mk('Roof slope R (crosswind)', CpRL_kr_R, Aroof, clauseRCrosswindC);
      }
      const uLee = effD + ov;
      const rwAlong = (S.roofType==='gable'||S.roofType==='hip') ? (S.depth/2 + ov) : (roofD + 2*ov);
      const zraw = roofCrosswindCpZonesWallSegment(rHD, h, -ov, uLee, rwAlong);
      return packZonedRoofFace('Roof slope R (crosswind)', zraw, CpRW_kr, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, clauseR);
    }
    if(useTableBC_UD) return mk('Roof Windward', CpRW_kr, Aroof, roofWWClause);
    const d0 = roofD;
    const uRidge = d0 / 2;
    const uLeeEave = d0 + ov;
    if(S.roofType==='gable'||S.roofType==='hip'){
      const zraw = roofCrosswindCpZonesWallSegment(rHD, h, -ov, uRidge, rw);
      return packZonedRoofFace('Roof Windward', zraw, CpRW_kr, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, roofWWClause);
    }
    const zraw = roofCrosswindCpZonesWallSegment(rHD, h, -ov, uLeeEave, rw);
    const fmRoof = getWindFaceMap();
    const wwName = rt0 === 'monoslope'
      ? (fmRoof.back === 'windward' ? 'Roof slope R (crosswind)' : 'Roof slope U (upwind)')
      : 'Roof Windward';
    return packZonedRoofFace(wwName, zraw, CpRW_kr, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, roofWWClause);
  };
  const mkRoofLW = ()=>{
    if(alongRidge){
      if(rCrosswindSlopeR_TableC){
        return mk('Roof slope R (crosswind)', CpRL_kr_R, Aroof, clauseRCrosswindC);
      }
      const uLee = effD + ov;
      const rwAlong = (S.roofType==='gable'||S.roofType==='hip') ? (S.depth/2 + ov) : (roofD + 2*ov);
      const zraw = roofCrosswindCpZonesWallSegment(rHD, h, -ov, uLee, rwAlong);
      return packZonedRoofFace('Roof slope R (crosswind)', zraw, CpRL_kr_R, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, clauseR);
    }
    if(useTableBC_UD) return mk('Roof Leeward', CpRL_kr_D, Aroof, roofLWClause);
    if(S.roofType==='gable'||S.roofType==='hip'){
      const d0 = S.depth;
      const zraw = roofCrosswindCpZonesWallSegment(rHD, h, d0/2, d0+ov, rw);
      return packZonedRoofFace('Roof Leeward', zraw, CpRL_kr_D, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, roofLWClause);
    }
    const lwName = rt0 === 'monoslope' ? 'Roof slope D (downwind)' : 'Roof Leeward';
    return mk(lwName, CpRL_kr_D, Aroof, roofLWClause);
  };
  // Hip roof: triangular “hip end” facets are roof slopes U, D, or R (Fig 5.2) — never wall S/W/L
  const ovHip = S.overhang || 0;
  const hipTriA = hipTriangleSlopeArea(S.width, S.depth, ovHip, S.pitch);
  const rightWallWW = relAng >= 45 && relAng < 135;
  const leftWallWW = relAng >= 225 && relAng < 315;
  let roof_hip_l, roof_hip_r;
  if(S.roofType === 'hip'){
    if(!alongRidge){
      const CpeR = S.pitch >= 10 ? CpRL_kr_R : hipEndCrosswindRCpeAlphaLt10(rHD, Kr);
      const clauseHipR = S.pitch >= 10
        ? 'Table 5.3(C) — crosswind slope R (hip)'
        : 'Table 5.3(A) — Fig 5.2 R (crosswind)';
      roof_hip_l = mk('Roof slope R (hip end)', CpeR, hipTriA, clauseHipR);
      roof_hip_r = mk('Roof slope R (hip end)', CpeR, hipTriA, clauseHipR);
    } else {
      let uZones = null, dZones = null;
      if(alphaLt10){
        const rlH = Math.max(S.width - S.depth, 2);
        const horizH = (S.width + 2 * ovHip - rlH) / 2;
        const baseE = S.depth + 2 * ovHip;
        uZones = hipEndUpwindCpZonesTableA(rHD, h, baseE, horizH, S.pitch);
        dZones = hipEndDownwindCpZonesTableA(rHD, h, S.depth, ovHip, S.pitch);
      }
      if(rightWallWW){
        if(alphaLt10 && uZones && uZones.length && dZones && dZones.length){
          roof_hip_l = packZonedRoofFace('Roof slope D (hip end)', dZones, CpRL_kr_D, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, roofLWClause);
          roof_hip_r = packZonedRoofFace('Roof slope U (hip end)', uZones, CpRW_kr, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, roofWWClause);
        } else {
          roof_hip_l = mk('Roof slope D (hip end)', CpRL_kr_D, hipTriA, roofLWClause);
          roof_hip_r = mk('Roof slope U (hip end)', CpRW_kr, hipTriA, roofWWClause);
        }
      } else {
        // alongRidge && leftWallWW (wind on +X / −X end walls per getWindFaceMap)
        if(alphaLt10 && uZones && uZones.length && dZones && dZones.length){
          roof_hip_l = packZonedRoofFace('Roof slope U (hip end)', uZones, CpRW_kr, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, roofWWClause);
          roof_hip_r = packZonedRoofFace('Roof slope D (hip end)', dZones, CpRL_kr_D, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, roofLWClause);
        } else {
          roof_hip_l = mk('Roof slope U (hip end)', CpRW_kr, hipTriA, roofWWClause);
          roof_hip_r = mk('Roof slope D (hip end)', CpRL_kr_D, hipTriA, roofLWClause);
        }
      }
    }
  }
  // Crosswind roof slope (Figure 5.2 R). Cpe from Table 5.3(C) — same magnitude as the
  // hip-end crosswind slope. For uploaded IFC models whose slope normal is perpendicular
  // to the wind, semanticToFaceKey emits 'roof_cw' and this entry drives the UI & pressure.
  const mkRoofCW = ()=> mk('Roof slope R (crosswind)', CpRL_kr_R, Aroof, clauseRCrosswindC);
  const faces={
    windward : mkWindward(),
    leeward  : mk('Leeward Wall',  CpLW, Alw,'Table 5.2(B)'),
    sidewall1: mkSidewall('Side Wall L'),
    sidewall2: mkSidewall('Side Wall R'),
    roof_ww  : mkRoofWW(),
    roof_lw  : mkRoofLW(),
    roof_cw  : mkRoofCW(),
    ...(S.roofType === 'hip' ? { roof_hip_l, roof_hip_r } : {})
  };

  const CpRW_max = useTableBC_UD ? roofUpwindCpMax(S.pitch, rHD) : CpRW;
  S.R = {VR,Vsit,Mz,Ms,Mt,Md,Mlee,qz,Cpi:cpi1,Kv:KvInit,impF,h,faces,angleOff,effW,effD,tc,Kr,CpRW_max};

  for(let i=0;i<8;i++){
    S.Vsit_dir[i] = VR * effectiveMd(i) * effectiveMzcat(i) * S.Ms[i] * S.Mt[i];
  }
}

/** When terrain category is unset, M_z,cat is null in state; use 1.0 for V Sit / q_z so results stay finite. */
function effectiveMzcat(i){
  const v = S.Mzcat[i];
  return (v != null && Number.isFinite(v)) ? v : 1;
}

function mzCat(h,tc){
  if(tc == null || tc === '' || !Number.isFinite(Number(tc))) return null;
  // AS/NZS 1170.2:2021 Table 4.1 — All regions except A0
  const T={
    1:  {3:0.97,5:1.01,10:1.08,15:1.12,20:1.14,30:1.18,40:1.21,50:1.23,75:1.27,100:1.31,150:1.36,200:1.39},
    2:  {3:0.91,5:0.91,10:1.00,15:1.05,20:1.08,30:1.12,40:1.16,50:1.18,75:1.22,100:1.24,150:1.27,200:1.29},
    2.5:{3:0.87,5:0.87,10:0.92,15:0.97,20:1.01,30:1.06,40:1.10,50:1.13,75:1.17,100:1.20,150:1.24,200:1.27},
    3:  {3:0.83,5:0.83,10:0.83,15:0.89,20:0.94,30:1.00,40:1.04,50:1.07,75:1.12,100:1.16,150:1.21,200:1.24},
    4:  {3:0.75,5:0.75,10:0.75,15:0.75,20:0.75,30:0.80,40:0.85,50:0.90,75:0.98,100:1.03,150:1.11,200:1.16}
  };
  const hs=[3,5,10,15,20,30,40,50,75,100,150,200];
  // For fractional TC (e.g. 2.67 from terrain averaging), interpolate between categories
  const tcs=[1,2,2.5,3,4];
  let t;
  if(T[tc]){
    t=T[tc];
  } else {
    // Linear interpolation between adjacent terrain categories
    let lo=1, hi=4;
    for(let i=0;i<tcs.length-1;i++){
      if(tc>=tcs[i]&&tc<=tcs[i+1]){ lo=tcs[i]; hi=tcs[i+1]; break; }
    }
    const frac=(tc-lo)/(hi-lo);
    t={};
    for(const h2 of hs) t[h2]=T[lo][h2]+frac*(T[hi][h2]-T[lo][h2]);
  }
  if(h<=3)return t[3]; if(h>=200)return t[200];
  for(let i=0;i<hs.length-1;i++) if(h>=hs[i]&&h<=hs[i+1]){
    const f=(h-hs[i])/(hs[i+1]-hs[i]); return t[hs[i]]+f*(t[hs[i+1]]-t[hs[i]]);
  }
  return t[10];
}
// Table 5.2(C) — Side wall Cp,e zones by distance from windward edge
// Returns zones with tributary width, area, Cp,e for proper force summation
function sidewallCpZones(effD, h, wallHeight){
  const z1w = Math.min(h, effD);
  const z2w = Math.max(0, Math.min(2*h, effD) - h);
  const z3w = Math.max(0, Math.min(3*h, effD) - 2*h);
  const z4w = Math.max(0, effD - 3*h);
  const zones = [];
  if(z1w > 0.001) zones.push({dist:'0 to 1h', Cpe:-0.65, width:z1w, area:z1w*wallHeight});
  if(z2w > 0.001) zones.push({dist:'1h to 2h', Cpe:-0.5, width:z2w, area:z2w*wallHeight});
  if(z3w > 0.001) zones.push({dist:'2h to 3h', Cpe:-0.3, width:z3w, area:z3w*wallHeight});
  if(z4w > 0.001) zones.push({dist:'> 3h', Cpe:-0.2, width:z4w, area:z4w*wallHeight});
  return zones.length ? zones : [{dist:'0 to 1h', Cpe:-0.65, width:effD, area:effD*wallHeight}];
}

// Table 5.3(A) — Upwind roof slope Cp,e zones for α < 10°
// roofDepth = horizontal distance from windward edge to ridge (effD/2 for gable/hip, effD for mono/flat)
// Returns zones with width, area, Cpe. h/d interpolates between columns when 0.5 < rHD < 1.0
function roofUpwindCpZones(rHD, h, roofDepth, roofWidth){
  const getCpe=(col)=>{ // col: 0=h/d≤0.5, 1=h/d≥1.0
    const c05=[-0.9,-0.9,-0.5,-0.3,-0.2], c10=[-1.3,-0.7,-0.5,-0.3,-0.2];
    if(col===0) return c05;
    if(col===1) return c10;
    const f=Math.max(0,Math.min(1,(rHD-0.5)/0.5));
    return c05.map((v,i)=>v+f*(c10[i]-v));
  };
  const cp=getCpe(rHD<=0.5?0:rHD>=1?1:0.5);
  const z1w=Math.min(0.5*h,roofDepth);
  const z2w=Math.max(0,Math.min(h,roofDepth)-0.5*h);
  const z3w=Math.max(0,Math.min(2*h,roofDepth)-h);
  const z4w=Math.max(0,Math.min(3*h,roofDepth)-2*h);
  const z5w=Math.max(0,roofDepth-3*h);
  const zones=[];
  const cosP=Math.cos(S.pitch*Math.PI/180)||1;
  if(z1w>0.001) zones.push({dist:'0 to 0.5h',Cpe:cp[0],width:z1w,area:z1w*roofWidth/cosP});
  if(z2w>0.001) zones.push({dist:'0.5h to 1h',Cpe:cp[1],width:z2w,area:z2w*roofWidth/cosP});
  if(z3w>0.001) zones.push({dist:'1h to 2h',Cpe:cp[2],width:z3w,area:z3w*roofWidth/cosP});
  if(z4w>0.001) zones.push({dist:'2h to 3h',Cpe:cp[3],width:z4w,area:z4w*roofWidth/cosP});
  if(z5w>0.001) zones.push({dist:'> 3h',Cpe:cp[4],width:z5w,area:z5w*roofWidth/cosP});
  return zones.length?zones:[{dist:'0 to 0.5h',Cpe:cp[0],width:roofDepth,area:roofDepth*roofWidth/cosP}];
}

// Hip end upwind (U), α < 10° — Table 5.3(A) bands along horizontal distance from windward eave; triangular plan area per band
function hipEndUpwindCpZonesTableA(rHD, hRef, B, horiz, pitchDeg){
  const cosP = Math.cos(pitchDeg * Math.PI / 180) || 1;
  const getCpe = col=>{
    const c05 = [-0.9, -0.9, -0.5, -0.3, -0.2], c10 = [-1.3, -0.7, -0.5, -0.3, -0.2];
    if(col === 0) return c05;
    if(col === 1) return c10;
    const f = Math.max(0, Math.min(1, (rHD - 0.5) / 0.5));
    return c05.map((v, i) => v + f * (c10[i] - v));
  };
  const cp = getCpe(rHD <= 0.5 ? 0 : rHD >= 1 ? 1 : 0.5);
  const chord = u=> (horiz < 1e-9 ? 0 : B * Math.max(0, 1 - u / horiz));
  const planStripArea = (u0, u1)=>{
    if(u1 <= u0) return 0;
    return 0.5 * (chord(u0) + chord(u1)) * (u1 - u0);
  };
  const bounds = [0, 0.5 * hRef, hRef, 2 * hRef, 3 * hRef, Infinity];
  const labels = ['0 to 0.5h', '0.5h to 1h', '1h to 2h', '2h to 3h', '> 3h'];
  const zones = [];
  for(let i = 0; i < 5; i++){
    const uLo = Math.max(bounds[i], 0);
    const uHi = Math.min(bounds[i + 1], horiz);
    if(uHi - uLo < 1e-6) continue;
    const pa = planStripArea(uLo, uHi);
    if(pa < 1e-12) continue;
    zones.push({dist: labels[i], Cpe: cp[i], width: uHi - uLo, area: pa / cosP});
  }
  if(!zones.length && horiz > 1e-6){
    const pa = planStripArea(0, horiz);
    zones.push({dist: labels[0], Cpe: cp[0], width: horiz, area: pa / cosP});
  }
  return zones;
}

// Hip end downwind (D), α < 10° — same strip pairs as leeward half of Table 5.3(A); u from ridge (d/2) to leeward eave (d+ov)
function hipEndDownwindCpZonesTableA(rHD, hRef, d, ov, pitchDeg){
  const uRidge = d / 2;
  const uLee = d + ov;
  const B = d + 2 * ov;
  const horizSpan = uLee - uRidge;
  const cosP = Math.cos(pitchDeg * Math.PI / 180) || 1;
  const chord = u=>{
    if(u <= uRidge || u >= uLee) return 0;
    return B * (u - uRidge) / horizSpan;
  };
  const planStripArea = (u0, u1)=>{
    if(u1 <= u0) return 0;
    return 0.5 * (chord(u0) + chord(u1)) * (u1 - u0);
  };
  const row05 = [[-0.9, -0.4], [-0.9, -0.4], [-0.5, 0], [-0.3, 0.1], [-0.2, 0.2]];
  const row10 = [[-1.3, -0.6], [-0.7, -0.3], [-0.5, 0], [-0.3, 0.1], [-0.2, 0.2]];
  const pairAt = i=>{
    if(rHD <= 0.5) return row05[i];
    if(rHD >= 1) return row10[i];
    const f = (rHD - 0.5) / 0.5;
    const a = row05[i], b = row10[i];
    return [a[0] + f * (b[0] - a[0]), a[1] + f * (b[1] - a[1])];
  };
  const bounds = [0, 0.5 * hRef, hRef, 2 * hRef, 3 * hRef, Infinity];
  const labels = ['0 to 0.5h', '0.5h to 1h', '1h to 2h', '2h to 3h', '> 3h'];
  const zones = [];
  for(let i = 0; i < 5; i++){
    const segLo = Math.max(uRidge, bounds[i]);
    const segHi = Math.min(uLee, bounds[i + 1]);
    if(segHi - segLo < 1e-6) continue;
    const pr = pairAt(i);
    const Cpe = Math.min(pr[0], pr[1]);
    const pa = planStripArea(segLo, segHi);
    if(pa < 1e-12) continue;
    zones.push({dist: labels[i], Cpe, width: segHi - segLo, area: pa / cosP});
  }
  if(!zones.length && horizSpan > 1e-6){
    const pr = pairAt(0);
    const Cpe = Math.min(pr[0], pr[1]);
    const pa = 0.5 * B * horizSpan;
    zones.push({dist: labels[0], Cpe, width: horizSpan, area: pa / cosP});
  }
  return zones;
}

// Table 5.3(A) crosswind: horizontal distance u from windward WALL plane (u=0 at wall; eave overhang has u<0, same Cp as 0–0.5h band)
function roofCrosswindCpZonesWallSegment(rHD, hRef, uMin, uMax, roofWidth){
  const bounds = [0, 0.5*hRef, hRef, 2*hRef, 3*hRef, Infinity];
  const labels = ['0 to 0.5h','0.5h to 1h','1h to 2h','2h to 3h','> 3h'];
  const row05 = [[-0.9,-0.4],[-0.9,-0.4],[-0.5,0],[-0.3,0.1],[-0.2,0.2]];
  const row10 = [[-1.3,-0.6],[-0.7,-0.3],[-0.5,0],[-0.3,0.1],[-0.2,0.2]];
  const pairAt = i=>{
    if(rHD <= 0.5) return row05[i];
    if(rHD >= 1) return row10[i];
    const f = (rHD - 0.5) / 0.5;
    const a = row05[i], b = row10[i];
    return [a[0]+f*(b[0]-a[0]), a[1]+f*(b[1]-a[1])];
  };
  const cosP = Math.cos(S.pitch*Math.PI/180) || 1;
  const zones = [];
  const loClamp = (i)=> (i === 0 ? -Infinity : bounds[i]);
  for(let i = 0; i < 5; i++){
    const segLo = Math.max(uMin, loClamp(i));
    const segHi = Math.min(uMax, bounds[i+1]);
    if(segHi - segLo < 0.001) continue;
    const pr = pairAt(i);
    const w = segHi - segLo;
    const uDispLo = segLo < 0 ? 0 : segLo;
    zones.push({
      dist: labels[i],
      uLo: segLo, uHi: segHi, uDispLo, uDispHi: segHi,
      Cpe_pair: pr,
      Cpe: Math.min(pr[0], pr[1]),
      width: w,
      area: w * roofWidth / cosP
    });
  }
  if(zones.length) return zones;
  const w = Math.max(uMax - uMin, 0.001);
  const pr = pairAt(0);
  return [{dist: labels[0], uLo: uMin, uHi: uMax, uDispLo: Math.max(0,uMin), uDispHi: uMax, Cpe_pair: pr, Cpe: Math.min(pr[0], pr[1]), width: w, area: w * roofWidth / cosP}];
}

function packZonedRoofFace(name, zonesRaw, fallbackCpe, qz, cpi1, cpi2, KvInit, kce1, kci1, Kr, clause){
  const Ka = 1.0;
  const KpUniform = getKpValue();
  let totalForce = 0, totalArea = 0;
  const zoneData = zonesRaw.map(z=>{
    const CpeKr = z.Cpe < 0 ? z.Cpe * Kr : z.Cpe;
    const Kp = (CpeKr < 0) ? KpUniform : 1.0;
    const pe = qz * CpeKr * Ka * Kp;
    const pi1 = qz * cpi1 * kci1 * KvInit;
    const pi2 = qz * cpi2 * kci1 * KvInit;
    const p = pNetTable55Case1(pe, kce1, pi1);
    const pAlt = pNetTable55Case1(pe, kce1, pi2);
    const force = Math.abs(p * z.area);
    totalForce += force;
    totalArea += z.area;
    return {dist: z.dist, Cpe: z.Cpe, area: z.area, width: z.width, p, p_case1:p, p_case2:pAlt, force};
  });
  let Cp_e = fallbackCpe;
  if(zoneData.length > 1 && totalArea > 0){
    const sumCpeA = zoneData.reduce((s, z)=> s + z.Cpe * z.area, 0);
    Cp_e = sumCpeA / totalArea;
  } else if(zoneData.length === 1){
    Cp_e = zoneData[0].Cpe;
  }
  const sumPA = zoneData.reduce((s, z)=> s + z.p * z.area, 0);
  const sumPA2 = zoneData.reduce((s, z)=> s + z.p_case2 * z.area, 0);
  const pAvg = totalArea > 0 ? sumPA / totalArea : 0;
  const pAvg2 = totalArea > 0 ? sumPA2 / totalArea : 0;
  return {
    name, Cp_e, zones: zoneData, Cp_i: cpi1, Cp_i_alt:cpi2,
    p: pAvg, p_case1:pAvg, p_case2:pAvg2, area: totalArea, force: totalForce, clause
  };
}

// Hip-end triangle slope area (one facet); ridge length rl = max(w−d, 2) matches buildHip geometry
function hipTriangleSlopeArea(w, d, ov, pitchDeg){
  const rl = Math.max(w - d, 2);
  const base = d + 2 * (ov || 0);
  const horiz = (w + 2 * (ov || 0) - rl) / 2;
  const plan = 0.5 * base * horiz;
  const c = Math.cos((pitchDeg || 0) * Math.PI / 180) || 1;
  return plan / c;
}

// Table 5.3(A) Fig 5.2 R — first distance band (0–0.5h), suction design value; Kr applied to negative Cp,e
function hipEndCrosswindRCpeAlphaLt10(rHD, Kr){
  const row05 = [-0.9, -0.4], row10 = [-1.3, -0.6];
  const f = rHD <= 0.5 ? 0 : rHD >= 1 ? 1 : (rHD - 0.5) / 0.5;
  const pr0 = row05[0] + f * (row10[0] - row05[0]);
  const pr1 = row05[1] + f * (row10[1] - row05[1]);
  const Cpe = Math.min(pr0, pr1);
  return Cpe < 0 ? Cpe * Kr : Cpe;
}

// Table 5.2(B) — Leeward wall Cp,e
// θ=0°: depends on roof pitch (α) and d/b ratio
// d = along-wind depth, b = crosswind breadth
function leewardCp(db, pitch){
  pitch = pitch || 0;
  if(pitch < 10){
    // α < 10°: d/b ≤1 → -0.5, d/b=2 → -0.3, d/b≥4 → -0.2
    if(db <= 1) return -0.5;
    if(db <= 2) return -0.5 + (db - 1) * 0.2;      // -0.5 → -0.3
    if(db <= 4) return -0.3 + (db - 2) / 2 * 0.1;   // -0.3 → -0.2
    return -0.2;
  }
  if(pitch < 15){
    // α=10°: d/b ≤1 → -0.5, d/b=2 → -0.3, d/b≥4 → -0.2 (same as α<10)
    // interpolate between α<10 row and α=15 row
    const cp10 = leewardCp(db, 9);
    const cp15 = -0.3; // α=15, all d/b → -0.3
    return cp10 + (pitch - 10) / 5 * (cp15 - cp10);
  }
  if(pitch < 20){
    // α=15°: all d/b → -0.3
    // α=20°: all d/b → -0.4
    return -0.3 + (pitch - 15) / 5 * (-0.1); // -0.3 → -0.4
  }
  if(pitch < 25){
    // α=20°: -0.4; α≥25°: depends on d/b
    const cp20 = -0.4;
    const cp25 = leewardCp25(db);
    return cp20 + (pitch - 20) / 5 * (cp25 - cp20);
  }
  // α ≥ 25°: d/b ≤0.1 → -0.75, d/b 0.3 → -0.5, d/b≥0.3 → -0.5
  return leewardCp25(db);
}
function leewardCp25(db){
  // Table 5.2(B) α≥25°: ≤0.3 → -0.5, ≤0.1 → -0.75  
  if(db <= 0.1) return -0.75;
  if(db <= 0.3) return -0.75 + (db - 0.1) / 0.2 * 0.25; // -0.75 → -0.5
  return -0.5;
}

// Table 5.3(B) — Roof upwind slope Cp,e for α ≥ 10°
// Uses the more negative (design critical) value from each row
// Interpolates between pitch angles and h/d ratios
function roofUpwindCp(pitch, hd){
  // Data: [pitch, [[hd, Cpe], ...]] — more negative value from each cell
  const data = [
    [10, [[0.25,-0.7],[0.5,-0.9],[1.0,-1.3]]],
    [15, [[0.25,-0.5],[0.5,-0.7],[1.0,-1.0]]],
    [20, [[0.25,-0.3],[0.5,-0.4],[1.0,-0.7]]],
    [25, [[0.25,-0.2],[0.5,-0.3],[1.0,-0.5]]],
    [30, [[0.25,-0.2],[0.5,-0.2],[1.0,-0.3]]],
    [35, [[0.25, 0.0],[0.5,-0.2],[1.0,-0.3]]],
    [45, [[0.25, 0.0],[0.5, 0.0],[1.0, 0.0]]]
  ];
  return interpRoofTable(data, pitch, hd);
}

// Table 5.3(B) — Roof upwind slope Cp,e (MAXIMUM / less negative) for α ≥ 10°
function roofUpwindCpMax(pitch, hd){
  const data = [
    [10, [[0.25, 0.0],[0.5, 0.0],[1.0, 0.0]]],
    [15, [[0.25, 0.0],[0.5, 0.0],[1.0, 0.0]]],
    [20, [[0.25, 0.3],[0.5, 0.2],[1.0, 0.0]]],
    [25, [[0.25, 0.4],[0.5, 0.3],[1.0, 0.1]]],
    [30, [[0.25, 0.5],[0.5, 0.3],[1.0, 0.2]]],
    [35, [[0.25, 0.5],[0.5, 0.3],[1.0, 0.2]]],
    [45, [[0.25, 0.5],[0.5, 0.5],[1.0, 0.5]]]
  ];
  return interpRoofTable(data, pitch, hd);
}

// Table 5.3(C) — Downwind slope (D): three h/d rows (≤0.25, 0.5, ≥1.0). Values per AS/NZS 1170.2:2021.
function roofDownwindSlopeD(pitch, hd){
  const data = [
    [10, [[0.25,-0.3],[0.5,-0.3],[1.0,-0.7]]],
    [15, [[0.25,-0.5],[0.5,-0.5],[1.0,-0.6]]],
    [20, [[0.25,-0.6],[0.5,-0.6],[1.0,-0.6]]],
    [25, [[0.25,-0.6],[0.5,-0.6],[1.0,-0.6]]],
    [30, [[0.25,-0.6],[0.5,-0.6],[1.0,-0.6]]],
    [35, [[0.25,-0.6],[0.5,-0.6],[1.0,-0.6]]],
    [45, [[0.25,-0.6],[0.5,-0.6],[1.0,-0.6]]]
  ];
  return interpRoofTable(data, pitch, hd);
}

// Table 5.3(C) — Crosswind slope (R) for hip roofs: two h/d rows (≤0.25, ≥1.0); interpolate between.
function roofCrosswindHipR(pitch, hd){
  const data = [
    [10, [[0.25,-0.3],[1.0,-0.5]]],
    [15, [[0.25,-0.5],[1.0,-0.5]]],
    [20, [[0.25,-0.6],[1.0,-0.6]]],
    [25, [[0.25,-0.6],[1.0,-0.6]]],
    [30, [[0.25,-0.6],[1.0,-0.6]]],
    [35, [[0.25,-0.6],[1.0,-0.6]]],
    [45, [[0.25,-0.6],[1.0,-0.6]]]
  ];
  return interpRoofTable(data, pitch, hd);
}

// Bilinear interpolation helper for roof tables
function interpRoofTable(data, pitch, hd){
  // Clamp pitch and h/d to table range
  pitch = Math.max(data[0][0], Math.min(data[data.length-1][0], pitch));
  hd = Math.max(0.25, Math.min(1.0, hd));

  // Find pitch bracket
  let pi = 0;
  for(let i = 0; i < data.length - 1; i++){
    if(pitch >= data[i][0] && pitch <= data[i+1][0]){ pi = i; break; }
    if(i === data.length - 2) pi = i;
  }
  const pFrac = data[pi+1][0] === data[pi][0] ? 0 :
    (pitch - data[pi][0]) / (data[pi+1][0] - data[pi][0]);

  // Interpolate h/d for lower pitch
  const cpLow = interpHd(data[pi][1], hd);
  // Interpolate h/d for upper pitch
  const cpHigh = interpHd(data[pi+1][1], hd);

  return cpLow + pFrac * (cpHigh - cpLow);
}

function interpHd(hdData, hd){
  hd = Math.max(hdData[0][0], Math.min(hdData[hdData.length-1][0], hd));
  for(let i = 0; i < hdData.length - 1; i++){
    if(hd >= hdData[i][0] && hd <= hdData[i+1][0]){
      const f = (hd - hdData[i][0]) / (hdData[i+1][0] - hdData[i][0]);
      return hdData[i][1] + f * (hdData[i+1][1] - hdData[i][1]);
    }
  }
  return hdData[hdData.length-1][1];
}

// ═══════════ INTERNAL PRESSURE — AS/NZS 1170.2 Table 5.1 ═══════════

function setPermCondition(cond){
  permCondition = cond;
  document.querySelectorAll('.perm-btn').forEach(b=>{
    b.classList.toggle('active', b.dataset.perm===cond);
    if(cond!=='auto') b.classList.remove('detected');
  });
  recalcPressures();
}

function toggleCpiManual(){
  const on = document.getElementById('cpi-manual-toggle')?.checked;
  const fields = document.getElementById('cpi-manual-fields');
  if(fields) fields.style.display = on ? '' : 'none';
  document.querySelectorAll('.perm-btn').forEach(b=>b.disabled=on);
  document.querySelectorAll('.cpi-auto-val').forEach(el=>el.style.opacity=on?'0.4':'1');
  recalcPressures();
}

const KL_DESIGN_MD_REGIONS = new Set(['B2','C','D']);

/** Table 5.6 Kl mesh, local pressure tab, Cl 3.3(b) M_d, and default Cp,i — tied to Local Pressure Zones toolbar. */
function isKlDesignMode(){
  return !!S.showPressureMap;
}

/** AS/NZS 1170.2:2021 Cl 3.3(b): cladding design in B2, C, D uses M_d = 1.0 (not Table 3.2). */
function mdOverrideKlDesign(){
  return isKlDesignMode() && KL_DESIGN_MD_REGIONS.has(S.region);
}

function effectiveMd(i){
  return mdOverrideKlDesign() ? 1 : S.Md[i];
}

/** Same M<sub>d</sub> as when Local Pressure Zones overlay is on (Cl 3.3(b) in B2/C/D). Used for Local Pressure Zones tab numbers. */
function effectiveMdKlDesignAlways(i){
  return KL_DESIGN_MD_REGIONS.has(S.region) ? 1 : S.Md[i];
}

/** Manual Cp,i > Local Pressure Zones defaults (−0.3 / +0.2) > Table 5.1 auto. */
function getCpiCasesForDesign(){
  if(document.getElementById('cpi-manual-toggle')?.checked){
    return {
      cpi1: parseFloat(document.getElementById('cpi1-manual')?.value)||0,
      cpi2: parseFloat(document.getElementById('cpi2-manual')?.value)||0,
      clause: 'Manual override',
      detected: null
    };
  }
  if(isKlDesignMode()){
    return {
      cpi1: -0.3,
      cpi2: 0.2,
      clause: 'Local Pressure Zones — default Cp,i = −0.3 / +0.2 (Table 5.1 envelope)',
      detected: 'kl-map'
    };
  }
  return calcCpiCases(permCondition);
}

// Return {cpi1, cpi2, clause, detected} based on permeability condition
// Convention: cpi1 = negative/suction (more severe for windward inward), cpi2 = neutral/positive
// Table 5.1(A): walls without openings > 0.5% and impermeable roof
// Table 5.1(B): walls with openings > 0.5% of surface area
function calcCpiCases(cond){
  const wwWallArea = S.width * S.height;
  const lwWallArea = S.width * S.height;
  const swWallArea = S.depth * S.height;
  const wwOpen = (S.openWW / 100) * wwWallArea;
  const lwOpen = (S.openLW / 100) * lwWallArea;
  const swOpen = (S.openSW / 100) * swWallArea;
  const totalOpen = wwOpen + lwOpen + swOpen * 2;
  const hasLargeOpenings = S.openWW > 0.5 || S.openLW > 0.5 || S.openSW > 0.5;

  if(cond==='auto'){
    if(!hasLargeOpenings || totalOpen < 0.01){
      const wwPerm = S.openWW >= 0.1;
      const lwPerm = S.openLW >= 0.1;
      const swPerm = S.openSW >= 0.1;
      const numPerm = (wwPerm?1:0) + (lwPerm?1:0) + (swPerm?2:0);
      if(numPerm===0){
        cond = 'sealed';
      } else if(numPerm===4){
        cond = 'permeable';
      } else if(numPerm===1){
        if(wwPerm) return {cpi1:0.7, cpi2:0.7, clause:'Table 5.1(A) — One wall permeable: windward (Cp,i = Cp,e = 0.7)', detected:'dom-ww'};
        return {cpi1:-0.3, cpi2:-0.3, clause:'Table 5.1(A) — One wall permeable: windward impermeable', detected:'dom-lw'};
      } else {
        if(wwPerm) return {cpi1:-0.1, cpi2:0.2, clause:'Table 5.1(A) — 2–3 walls permeable, windward permeable', detected:'dom-ww'};
        return {cpi1:-0.3, cpi2:-0.3, clause:'Table 5.1(A) — 2–3 walls permeable, windward impermeable', detected:'dom-lw'};
      }
    } else {
      const maxOpen = Math.max(wwOpen, lwOpen, swOpen);
      const otherOpen = totalOpen - maxOpen;
      const ratio = otherOpen > 0 ? maxOpen / otherOpen : 6;
      let face;
      if(maxOpen===wwOpen) face='windward';
      else if(maxOpen===lwOpen) face='leeward';
      else face='sidewall';
      return calcCpiTable5_1B(ratio, face);
    }
  }

  switch(cond){
    case 'sealed':
      return {cpi1:-0.2, cpi2:0.0, clause:'Table 5.1(A) — Effectively sealed & non-opening windows', detected:'sealed'};
    case 'permeable':
      return {cpi1:-0.3, cpi2:0.0, clause:'Table 5.1(A) — All walls equally permeable', detected:'permeable'};
    case 'dom-ww':{
      if(hasLargeOpenings){
        const otherOpen = totalOpen - wwOpen;
        const ratio = otherOpen > 0 ? wwOpen / otherOpen : 6;
        return calcCpiTable5_1B(ratio, 'windward');
      }
      return {cpi1:0.7, cpi2:0.7, clause:'Table 5.1(A) — One wall permeable, windward: Cp,i = Cp,e', detected:'dom-ww'};
    }
    case 'dom-lw':{
      if(hasLargeOpenings){
        const maxLW = Math.max(lwOpen, swOpen);
        const otherOpen = totalOpen - maxLW;
        const ratio = otherOpen > 0 ? maxLW / otherOpen : 6;
        return calcCpiTable5_1B(ratio, maxLW===lwOpen?'leeward':'sidewall');
      }
      return {cpi1:-0.3, cpi2:-0.3, clause:'Table 5.1(A) — Windward wall impermeable', detected:'dom-lw'};
    }
    default:
      return {cpi1:-0.3, cpi2:0.0, clause:'Table 5.1(A)', detected:'permeable'};
  }
}

// Table 5.1(B) — Cp,i from opening ratio and dominant face
function calcCpiTable5_1B(ratio, dominantFace){
  // External Cp,e at dominant opening location
  let CpeAtOpening;
  if(dominantFace==='windward'){
    CpeAtOpening = 0.7; // Table 5.2(A)
  } else if(dominantFace==='leeward'){
    const effW=S.width, effD=S.depth;
    const db = effD / effW;
    CpeAtOpening = leewardCp(db, S.pitch);
  } else {
    CpeAtOpening = -0.65; // Table 5.2(C)
  }
  const Ka=1.0, KrOpen=1.0; // Conservative defaults (Ka per Cl.5.4.2, Kr per Cl.5.4.4)
  const cpe = Ka * KrOpen * CpeAtOpening;
  const det = dominantFace==='windward' ? 'dom-ww' : 'dom-lw';
  const fLabel = dominantFace==='windward' ? 'windward' : dominantFace==='leeward' ? 'leeward' : 'side wall';

  if(ratio <= 0.5){
    return {cpi1:-0.3, cpi2:0.0, clause:`Table 5.1(B) — Ratio ≤ 0.5, no dominant opening`, detected:det};
  }
  if(ratio <= 1){
    if(dominantFace==='windward'){
      const f=(ratio-0.5)/0.5;
      return {cpi1:-0.3+f*0.2, cpi2:0.0+f*0.2, clause:`Table 5.1(B) — Ratio ${ratio.toFixed(1)}, largest on ${fLabel}`, detected:det};
    }
    return {cpi1:-0.3, cpi2:0.0, clause:`Table 5.1(B) — Ratio ${ratio.toFixed(1)}, largest on ${fLabel}`, detected:det};
  }
  if(ratio <= 2){
    const f=(ratio-1)/1;
    if(dominantFace==='windward'){
      const cpiAt2 = 0.7 * cpe;
      return {cpi1:-0.1+f*(cpiAt2-(-0.1)), cpi2:0.2+f*(cpiAt2-0.2), clause:`Table 5.1(B) — Ratio ${ratio.toFixed(1)}, ${fLabel} (interp → 0.7·Cp,e)`, detected:det};
    }
    return {cpi1:-0.3+f*(cpe-(-0.3)), cpi2:0.0+f*(cpe-0.0), clause:`Table 5.1(B) — Ratio ${ratio.toFixed(1)}, ${fLabel} (interp → Ka·Kr·Cp,e)`, detected:det};
  }
  // Ratio > 2: single Cp,i value from Cp,e formula
  let factor;
  if(ratio <= 3){
    factor = dominantFace==='windward' ? 0.7+(ratio-2)/1*0.15 : 1.0; // 0.7→0.85
  } else if(ratio <= 6){
    factor = dominantFace==='windward' ? 0.85+(ratio-3)/3*0.15 : 1.0; // 0.85→1.0
  } else {
    factor = 1.0;
  }
  const cpiVal = factor * cpe;
  const fStr = dominantFace==='windward' && factor<1 ? `${(factor*100).toFixed(0)}%·` : '';
  return {cpi1:cpiVal, cpi2:cpiVal, clause:`Table 5.1(B) — Ratio ${ratio>=6?'≥ 6':ratio.toFixed(1)}, ${fLabel} (${fStr}Ka·Kr·Cp,e = ${cpiVal.toFixed(2)})`, detected:det};
}

// Legacy single-value Cpi for 3D model display
function calcCpi(){
  return getCpiCasesForDesign().cpi1;
}

// ═══════════ COMBINATION FACTOR — AS/NZS 1170.2 Table 5.5 ═══════════

const KC_DESIGN_CASES = {
  a: { kce: 0.8, kci: 1.0, label: '3 surfaces (WW+LW+roof), internal not effective' },
  b: { kce: 0.8, kci: 0.8, label: '4 surfaces (WW+LW+roof+internal)' },
  c: { kce: 0.8, kci: 1.0, label: '3 surfaces (SW+roof), internal not effective' },
  d: { kce: 0.8, kci: 0.8, label: '4 surfaces (SW+roof+internal)' },
  e: { kce: 1.0, kci: 1.0, label: '1 surface (roof alone)' },
  f: { kce: 0.9, kci: 0.9, label: '2 surfaces (roof+internal)' },
  g: { kce: 0.9, kci: 1.0, label: '2 surfaces (lateral WW+LW), internal not effective' },
  h: { kce: 0.9, kci: 0.9, label: '2 surfaces (lateral walls+internal)' }
};

function setKcDesignCase(caseId){
  const dc = KC_DESIGN_CASES[caseId];
  if(dc){
    const ei = document.getElementById('kci1-val');
    const ee = document.getElementById('kce1-val');
    if(ei) ei.value = String(dc.kci);
    if(ee) ee.value = String(dc.kce);
  }
  renderKcDiagram(caseId);
  validateKcProduct();
  calc();
  recalcPressures();
}

function onKcManualChange(){
  const sel = document.getElementById('kc-design-case');
  if(!sel) return;
  const kci = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;
  const kce = parseFloat(document.getElementById('kce1-val')?.value) || 0.8;
  const match = Object.entries(KC_DESIGN_CASES).find(
    ([, dc]) => dc.kce === kce && dc.kci === kci
  );
  sel.value = match ? match[0] : 'custom';
  renderKcDiagram(sel.value);
  validateKcProduct();
  calc();
  recalcPressures();
}

// ── Kc design case isometric diagram ──

const KC_CASE_DEFS = {
  a:      { ww:1, lw:1, sw:0, roof:1, int:0, desc:'WW + LW + Roof' },
  b:      { ww:1, lw:1, sw:0, roof:1, int:1, desc:'WW + LW + Roof + Internal' },
  c:      { ww:0, lw:0, sw:1, roof:1, int:0, desc:'SW + Roof' },
  d:      { ww:0, lw:0, sw:1, roof:1, int:1, desc:'SW + Roof + Internal' },
  e:      { ww:0, lw:0, sw:0, roof:1, int:0, desc:'Roof alone' },
  f:      { ww:0, lw:0, sw:0, roof:1, int:1, desc:'Roof + Internal' },
  g:      { ww:1, lw:1, sw:0, roof:0, int:0, desc:'Lateral WW + LW' },
  h:      { ww:1, lw:1, sw:0, roof:0, int:1, desc:'Lateral walls + Internal' },
  custom: { ww:0, lw:0, sw:0, roof:0, int:0, desc:'Custom Kc values' }
};

function renderKcDiagram(caseId){
  const el = document.getElementById('kc-diagram-svg');
  if(!el) return;
  const c = KC_CASE_DEFS[caseId] || KC_CASE_DEFS.custom;

  const fbl='50,98', fbr='111,132', ftl='50,43', ftr='111,77';
  const bbl='89,75', bbr='150,110', btl='89,20', btr='150,55';

  const hi  = 'rgba(0,210,255,0.35)';
  const dim = 'rgba(200,220,255,0.06)';
  const stk = 'rgba(200,220,255,0.4)';
  const hiStk = 'rgba(0,210,255,0.7)';

  const roofF = c.roof ? hi : dim, wwF = c.ww ? hi : dim, swF = c.sw ? hi : dim;
  const roofS = c.roof ? hiStk : stk, wwS = c.ww ? hiStk : stk, swS = c.sw ? hiStk : stk;

  let s = `<defs><marker id="arrKc" markerWidth="6" markerHeight="4" refX="5" refY="2" orient="auto">` +
    `<path d="M0,0 L6,2 L0,4" fill="rgba(255,255,255,0.5)"/></marker></defs>`;

  s += `<polygon points="${btl} ${btr} ${ftr} ${ftl}" fill="${roofF}" stroke="${roofS}" stroke-width="1.2"/>`;
  s += `<polygon points="${bbr} ${fbr} ${ftr} ${btr}" fill="${swF}" stroke="${swS}" stroke-width="1.2"/>`;
  s += `<polygon points="${fbl} ${fbr} ${ftr} ${ftl}" fill="${wwF}" stroke="${wwS}" stroke-width="1.2"/>`;

  s += `<line x1="89" y1="75" x2="150" y2="110" stroke="${stk}" stroke-width="1"/>`;
  s += `<line x1="89" y1="75" x2="89" y2="20" stroke="${stk}" stroke-width="1"/>`;
  s += `<line x1="89" y1="75" x2="50" y2="98" stroke="${stk}" stroke-width="0.8" stroke-dasharray="3,2"/>`;

  s += `<text x="80" y="94" fill="white" font-size="9" text-anchor="middle" opacity="0.65" font-weight="600">WW</text>`;
  s += `<text x="133" y="97" fill="white" font-size="9" text-anchor="middle" opacity="0.65" font-weight="600">SW</text>`;
  s += `<text x="100" y="52" fill="white" font-size="9" text-anchor="middle" opacity="0.65" font-weight="600">Roof</text>`;

  if(c.lw){
    s += `<line x1="89" y1="75" x2="89" y2="20" stroke="${hi}" stroke-width="2.5" opacity="0.8"/>`;
    s += `<line x1="89" y1="75" x2="50" y2="98" stroke="${hi}" stroke-width="2.5" opacity="0.8" stroke-dasharray="4,2"/>`;
    s += `<text x="62" y="70" fill="rgba(0,210,255,0.9)" font-size="8" text-anchor="end" font-weight="600">LW</text>`;
  }

  if(c.int){
    s += `<rect x="70" y="96" width="18" height="14" rx="3" fill="rgba(255,171,0,0.2)" stroke="rgba(255,171,0,0.55)" stroke-width="1"/>`;
    s += `<text x="79" y="106" fill="rgba(255,171,0,0.9)" font-size="7" text-anchor="middle" font-weight="700">p\u1d62</text>`;
  }

  s += `<line x1="22" y1="128" x2="38" y2="113" stroke="rgba(255,255,255,0.45)" stroke-width="1.5" marker-end="url(#arrKc)"/>`;
  s += `<text x="14" y="136" fill="rgba(255,255,255,0.35)" font-size="7" font-style="italic">wind</text>`;

  s += `<text x="100" y="150" fill="rgba(255,255,255,0.55)" font-size="9" text-anchor="middle">(${caseId}) ${c.desc}</text>`;

  el.innerHTML = s;
}

function validateKcProduct(){
  const warn = document.getElementById('kc-warn');
  if(!warn) return;
  const kce = parseFloat(document.getElementById('kce1-val')?.value) || 0.8;
  const kci = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;
  const prod = kce * kci;
  if(prod < 0.8 - 1e-9){
    warn.textContent = `Kc,e × Kc,i = ${prod.toFixed(2)} < 0.8 — Table 5.5 requires product ≥ 0.8`;
    warn.style.display = '';
  } else {
    warn.style.display = 'none';
  }
}

function suggestKcFromPermeability(){
  const hint = document.getElementById('kc-hint');
  const sel = document.getElementById('kc-design-case');
  if(!hint) return;
  const cpiResult = getCpiCasesForDesign();
  const maxAbsCpi = Math.max(Math.abs(cpiResult.cpi1), Math.abs(cpiResult.cpi2));
  if(maxAbsCpi <= 0.21){
    hint.innerHTML = 'Permeability → small C<sub>p,i</sub> — internal likely <b>not effective</b> (K<sub>c,i</sub> = 1.0)';
    hint.style.display = '';
  } else if(maxAbsCpi >= 0.5){
    hint.innerHTML = 'Permeability → significant C<sub>p,i</sub> — internal likely <b>effective</b> (K<sub>c,i</sub> = 0.8–0.9)';
    hint.style.display = '';
  } else {
    hint.style.display = 'none';
  }
}

// ═══════════════════════════════════════════════
//   BUILD 3D SCENE
// ═══════════════════════════════════════════════
function onInput(){
  readUI();
  if(S.analysisLocked){
    // When locked, don't recalc — keep frozen results
    // But still read UI so sliders update labels
    return;
  }
  calc();
  rebuild();
  refreshDirectionalWindUI();
  syncWindDirSlider();
  recalcPressures();
  // Update ref height display on structure panel
  const rhEl = document.getElementById('val-ref-height');
  if(rhEl && S.R && S.R.h) rhEl.textContent = S.R.h.toFixed(1) + ' m';
}

function syncWindDirSlider(){
  const el = document.getElementById('inp-winddir');
  const lbl = document.getElementById('val-winddir');
  if(el) el.value = Math.round(S.windAngle / 5) * 5;
  if(lbl) lbl.textContent = Math.round(S.windAngle);
}

function onWindDirSlider(val){
  S.windAngle = parseFloat(val) || 0;
  document.getElementById('val-winddir').textContent = Math.round(S.windAngle);
  if(S.analysisLocked) return;
  calc();
  rebuild();
  drawCompass();
  drawMapWindArrow();
  updateMapBuilding();
  refreshDirectionalWindUI();
  recalcPressures();
}

/** Label next to opening % — extra decimals when value is small. */
function formatOpeningPctLabel(n){
  n = Number(n);
  if(!Number.isFinite(n)) return '0';
  if(n < 1) return n.toFixed(2);
  if(Math.abs(n - Math.round(n)) < 1e-9) return String(Math.round(n));
  return (Math.round(n * 10) / 10).toFixed(1);
}

/** face: 'ww' | 'lw' | 'sw' — clamp, sync number + range + label, recalc. */
function syncOpeningPct(face, rawValue){
  let n = parseFloat(rawValue);
  if(!Number.isFinite(n)) n = 0;
  n = Math.min(60, Math.max(0, n));
  const num = document.getElementById('inp-open-'+face);
  const rng = document.getElementById('inp-open-'+face+'-range');
  const val = document.getElementById('val-open-'+face);
  if(num) num.value = n;
  if(rng) rng.value = n;
  if(val) val.textContent = formatOpeningPctLabel(n);
  onInput();
}

function syncRoofTypeButtons(rt){
  document.querySelectorAll('.rooftype-btn').forEach(btn=>{
    btn.classList.toggle('active', btn.getAttribute('data-roof') === rt);
  });
}

function setRoofType(rt){
  const h = document.getElementById('inp-rooftype');
  if(h) h.value = rt;
  syncRoofTypeButtons(rt);
  onInput();
}

function readUI(){
  const v=id=>{const el=document.getElementById(id);return el?parseFloat(el.value):0};
  const s=id=>{const el=document.getElementById(id);return el?el.value:''};
  S.width=v('inp-width');S.depth=v('inp-depth');S.height=v('inp-height');
  S.pitch=v('inp-pitch');
  S.roofType=s('inp-rooftype');
  S.parapet=v('inp-parapet');S.overhang=v('inp-overhang');
  S.terrainCat=parseFloat(s('inp-terrain')||'3');
  S.importance=parseInt(s('inp-importance')||'2',10);
  S.life=parseInt(s('inp-life')||'50',10)||50;
  S.region=s('inp-region')||'NZ1';
  S.ari=windUlsReturnPeriodYears(S.life,S.importance);
  if(!S.vrManual){
    const vr=regionalWindSpeedFromReturnPeriod(S.region,S.ari);
    S.windSpeed=vr;
    S.svcVr=Math.round(vr*0.71);
    const wEl=document.getElementById('inp-windspeed');
    if(wEl) wEl.value=String(S.windSpeed);
    const svcEl=document.getElementById('inp-svc-vr');
    if(svcEl) svcEl.value=String(S.svcVr);
  }else{
    S.windSpeed=parseFloat(document.getElementById('inp-windspeed')?.value)||66;
    const svcRaw=document.getElementById('inp-svc-vr')?.value;
    S.svcVr=svcRaw!=null&&svcRaw!==''?parseInt(svcRaw,10):Math.round(S.windSpeed*0.71);
  }
  updateDesignAriDisplay();
  S.loadCase=s('inp-loadcase')||'A';
  S.openWW=v('inp-open-ww');S.openLW=v('inp-open-lw');S.openSW=v('inp-open-sw');
  S.Kp=s('inp-kp')||'1.0';
  // Show/hide Kp info text
  const kpInfo=document.getElementById('kp-info');
  if(kpInfo) kpInfo.style.display = S.Kp!=='1.0' ? 'block' : 'none';

  ['width','depth','height','pitch','parapet','overhang',
   'open-ww','open-lw','open-sw'].forEach(k=>{
    const el=document.getElementById('val-'+k);
    const inp=document.getElementById('inp-'+k);
    if(!el||!inp) return;
    if(k.startsWith('open-')){
      const n = parseFloat(inp.value);
      el.textContent = formatOpeningPctLabel(Number.isFinite(n) ? n : 0);
    }else{
      el.textContent = inp.value;
    }
  });
  readKvOverrideFromUi();
}

function rebuild(){
  [grpBuild,grpDim,grpLabel,grpArrows,grpInternal,grpWind].forEach(clearGrp);
  faceMap.clear();

  buildWalls();
  buildRoof();

  if(S.showDimensions) buildDims();
  if(S.showLabels) buildLabels();
  if(S.showPressureArrows) buildPArrows();
  if(S.showInternalPressure) buildInternal();
  if(S.showParticles) buildParticles();
  buildWindIndicator();
  buildNorthArrow();

  // Apply building orientation rotation to all building groups (wind must match calc() / Fig 5.2 — rel. to building)
  const bRad = -S.mapBuildingAngle * Math.PI / 180;
  [grpBuild,grpDim,grpLabel,grpArrows,grpInternal,grpWind].forEach(g=>{ g.rotation.y = bRad; });

  // If an uploaded model is active and visible, keep parametric hidden
  if(uploadedModelGroup && uploadedModelVisible && !parametricVisible){
    grpBuild.visible = false;
    grpDim.visible = false;
    grpLabel.visible = false;
    grpArrows.visible = false;
    grpInternal.visible = false;

    // Apply wind analysis to uploaded model
    void applyWindToUploadedModel();
  } else {
    // No uploaded model active — hide upload overlay if it exists
    if(grpUploadOverlay) grpUploadOverlay.visible = false;
  }

  updateLegend();
  drawCompass();

  if(leafletMap) updateMapBuilding();
}

function rebuildScene(){
  rebuild();
}

function clearGrp(g){
  while(g.children.length){
    const o=g.children[0];
    if(o.geometry)o.geometry.dispose();
    if(o.material){
      if(Array.isArray(o.material))o.material.forEach(m=>m.dispose()); else o.material.dispose();
    }
    g.remove(o);
  }
}

// ── Material helpers ──
function wallMat(fk,defCol){
  const F=S.R.faces, isTr=S.viewMode==='transparent', isWf=S.viewMode==='wireframe';
  const col=S.showHeatmap?heatCol(F[fk].p):defCol;
  const o={color:col,roughness:.6,metalness:.1,
    transparent:isTr||S.showHeatmap,opacity:isTr?(S.showHeatmap?.75:.4):1,
    wireframe:isWf,side:THREE.DoubleSide};
  return new THREE.MeshStandardMaterial(o);
}
function roofMat(fk){
  const F=S.R.faces;
  const fp = (F[fk] && Number.isFinite(F[fk].p)) ? F[fk].p : 0;
  const col=S.showHeatmap?heatCol(fp):0xcc6622;
  const isTr=S.viewMode==='transparent';
  const o={color:col,roughness:.5,side:THREE.DoubleSide,
    transparent:isTr||S.showHeatmap,
    opacity:isTr?(S.showHeatmap?.75:.4):1,wireframe:S.viewMode==='wireframe'};
  return new THREE.MeshStandardMaterial(o);
}

function addFace(geo,fk,defCol,pos,rot,grp){
  grp=grp||grpBuild;
  const m=new THREE.Mesh(geo,wallMat(fk,defCol));
  if(pos)m.position.copy(pos);if(rot)m.rotation.copy(rot);
  m.castShadow=true;m.receiveShadow=true;
  faceMap.set(m.uuid,{key:fk,...S.R.faces[fk]});
  grp.add(m);
  if(S.viewMode!=='wireframe'){
    const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25}));
    if(pos)e.position.copy(pos);if(rot)e.rotation.copy(rot);
    grp.add(e);
  }
  return m;
}

// Table 5.2(C) — Sidewall Cp,e zones in 3D: subdivides sidewall into strips by distance from windward edge
// faceKey: sidewall1 or sidewall2; wallLength: along-wind dimension (d or w); wallH: height
// centerPos: face center; varAxis: 'x'|'z' for strip offset axis; windwardAtHigh: windward end is +half
function addSidewallCpZones(faceKey, wallLength, wallH, centerPos, varAxis, windwardAtHigh){
  if(!S.R?.faces?.[faceKey]) return;
  const faceData = S.R.faces[faceKey];
  const zones = faceData.zones && faceData.zones.length ? faceData.zones
    : sidewallCpZones(wallLength, S.R.h, wallH);
  let cum = 0;
  const half = wallLength / 2;
  zones.forEach(z => {
    const stripW = z.width;
    if(stripW < 0.01) return;
    const geo = new THREE.PlaneGeometry(stripW, wallH);
    const offset = windwardAtHigh ? (half - cum - stripW/2) : (-half + cum + stripW/2);
    const pos = centerPos.clone();
    if(varAxis === 'z') pos.z += offset; else pos.x += offset;
    const yRot = varAxis === 'z' ? -Math.PI/2 : 0;
    const rot = new THREE.Euler(0, yRot, 0);
    const zoneData = {key:faceKey, name:faceData.name+' ('+z.dist+')', Cp_e:z.Cpe, Cp_i:faceData.Cp_i, p:z.p, area:z.area, force:z.force, clause:faceData.clause};
    const col = S.showHeatmap ? heatCol(z.p) : 0xcc8844;
    const mat = new THREE.MeshStandardMaterial({
      color:col, roughness:.6, metalness:.1, transparent:S.viewMode==='transparent', opacity:S.viewMode==='transparent'?.75:1,
      wireframe:S.viewMode==='wireframe', side:THREE.DoubleSide
    });
    const m = new THREE.Mesh(geo, mat);
    m.position.copy(pos); m.rotation.copy(rot);
    m.castShadow = true; m.receiveShadow = true;
    faceMap.set(m.uuid, zoneData);
    grpBuild.add(m);
    if(S.viewMode !== 'wireframe'){
      const e = new THREE.LineSegments(new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({color:0x000000, transparent:true, opacity:.25}));
      e.position.copy(pos); e.rotation.copy(rot);
      grpBuild.add(e);
    }
    cum += stripW;
  });
}

// Table 5.2(C) — Sidewall Cp,e zones in 3D: creates zone strips with zone-specific face data
// centerPos: face center; axis: 'x'|'z'; windwardEnd: +1 = high end, -1 = low end
function addSidewallCpZones(faceKey, wallLength, wallH, centerPos, yRot, axis, windwardEnd, grp){
  grp=grp||grpBuild;
  const fd=S.R.faces[faceKey];
  if(!fd) return;
  const zones=fd.zones||sidewallCpZones(wallLength,S.R.h,wallH);
  let cumWidth=0;
  const half=wallLength/2;
  zones.forEach(z=>{
    const zoneW=Math.min(z.width,wallLength-cumWidth);
    if(zoneW<0.01) return;
    const stripCenterOffset=windwardEnd*(half-cumWidth-zoneW/2);
    const pos=centerPos.clone();
    if(axis==='z') pos.z+=stripCenterOffset; else pos.x+=stripCenterOffset;
    const geo=new THREE.PlaneGeometry(zoneW,wallH);
    const zoneData={key:faceKey,name:fd.name+' ('+z.dist+')',Cp_e:z.Cpe,Cp_i:fd.Cp_i,p:z.p,area:z.area,force:z.force,clause:fd.clause};
    const col=S.showHeatmap?heatCol(z.p):0xcc8844;
    const mat=new THREE.MeshStandardMaterial({color:col,roughness:.6,metalness:.1,
      transparent:S.viewMode==='transparent'||S.showHeatmap,opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe',side:THREE.DoubleSide});
    const m=new THREE.Mesh(geo,mat);
    m.position.copy(pos); m.rotation.set(0,yRot,0);
    m.castShadow=true; m.receiveShadow=true;
    faceMap.set(m.uuid,zoneData);
    grp.add(m);
    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25}));
      e.position.copy(pos); e.rotation.set(0,yRot,0);
      grp.add(e);
    }
    cumWidth+=zoneW;
  });
}

// Table 5.2(C) — Sidewall Cp,e zones in 3D: subdivides sidewall into strips by distance from windward edge
// faceKey: sidewall1 or sidewall2; wallLength: along-wind dimension; axis: 'x'|'z'; windwardSign: +1 = windward at +axis
function addSidewallCpZones(faceKey, wallLength, wallH, centerPos, axis, windwardSign, yRot){
  const F=S.R?.faces;
  if(!F||!F[faceKey]) return;
  const faceData=F[faceKey];
  let zones=faceData.zones;
  if(!zones||!zones.length) zones=sidewallCpZones(wallLength,S.R?.h||S.height,S.height).map(z=>({...z,p:0,force:0}));
  let cum=0;
  zones.forEach((z,i)=>{
    const stripW=z.width||z.area/h;
    if(stripW<0.01) return;
    const geo=new THREE.PlaneGeometry(stripW,wallH);
    const stripCenter=windwardSign*(wallLength/2-cum-stripW/2);
    const pos=centerPos.clone();
    if(axis==='z') pos.z+=stripCenter; else pos.x+=stripCenter;
    const col=S.showHeatmap?heatCol(z.p):0xcc8844;
    const mat=new THREE.MeshStandardMaterial({color:col,roughness:.6,metalness:.1,
      transparent:S.viewMode==='transparent'||S.showHeatmap,opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe',side:THREE.DoubleSide});
    const m=new THREE.Mesh(geo,mat);
    m.position.copy(pos); m.rotation.set(0,yRot,0);
    m.castShadow=true; m.receiveShadow=true;
    faceMap.set(m.uuid,{key:faceKey,name:faceData.name+' ('+z.dist+')',Cp_e:z.Cpe,Cp_i:faceData.Cp_i,p:z.p,area:z.area,force:z.force,clause:faceData.clause,Kl:1.0});
    grpBuild.add(m);
    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25}));
      e.position.copy(pos); e.rotation.set(0,yRot,0);
      grpBuild.add(e);
    }
    cum+=stripW;
  });
}

// Table 5.2(C) — Sidewall Cp,e zones in 3D: create zone strips for hover/heatmap
// wallLength = along-wind extent, windwardEnd = +1 (pos end) or -1 (neg end)
// axis: 'z' for left/right (extent along Z), 'x' for front/back (extent along X)
function addSidewallCpZones(faceKey, wallLength, wallH, centerPos, yRot, axis, windwardSign, grp){
  grp=grp||grpBuild;
  const faceData=S.R?.faces?.[faceKey];
  if(!faceData) return;
  const zones=faceData.zones||[];
  if(zones.length===0) return;

  const half=wallLength/2;
  let cum=0;
  zones.forEach((z,i)=>{
    const zw=z.width;
    if(zw<0.01) return;
    const stripCenter=windwardSign*(half-cum-zw/2);
    const pos=new THREE.Vector3().copy(centerPos);
    if(axis==='z') pos.z+=stripCenter; else pos.x+=stripCenter;

    const zoneData={key:faceKey,name:faceData.name+' ('+z.dist+')',Cp_e:z.Cpe,Cp_i:faceData.Cp_i,p:z.p,area:z.area,force:z.force,clause:faceData.clause};
    const col=S.showHeatmap?heatCol(z.p):0xcc8844;
    const mat=new THREE.MeshStandardMaterial({
      color:col,roughness:.6,metalness:.1,transparent:S.viewMode==='transparent'||S.showHeatmap,opacity:(S.viewMode==='transparent'||S.showHeatmap)?0.75:1,wireframe:S.viewMode==='wireframe',side:THREE.DoubleSide
    });
    const geo=new THREE.PlaneGeometry(zw,wallH);
    const m=new THREE.Mesh(geo,mat);
    m.position.copy(pos); m.rotation.set(0,yRot,0);
    m.castShadow=true; m.receiveShadow=true;
    faceMap.set(m.uuid,zoneData);
    grp.add(m);
    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25}));
      e.position.copy(pos); e.rotation.set(0,yRot,0);
      grp.add(e);
    }
    cum+=zw;
  });
}

// Table 5.2(C) — Sidewall Cp,e zones in 3D: creates zone strips so hover shows zone-specific values
function addSidewallCpZones(faceKey, wallLength, wallH, centerPos, axis, windwardSign, yRot, grp){
  grp=grp||grpBuild;
  const faceData = S.R?.faces?.[faceKey];
  const zones = faceData?.zones;
  if(!zones || zones.length===0){
    const effD = S.R?.effD ?? wallLength;
    const hRef = S.R?.h ?? S.height;
    const swZones = sidewallCpZones(effD, hRef, wallH);
    if(!swZones.length) return addFace(new THREE.PlaneGeometry(wallLength,wallH), faceKey, 0xcc8844, centerPos, new THREE.Euler(0,yRot,0), grp);
    return swZones.forEach((z,i)=>{
      const cum = swZones.slice(0,i).reduce((s,zz)=>s+zz.width,0);
      const cx = axis==='x' ? (windwardSign>0 ? wallLength/2 - cum - z.width/2 : -wallLength/2 + cum + z.width/2) : 0;
      const cz = axis==='z' ? (windwardSign>0 ? wallLength/2 - cum - z.width/2 : -wallLength/2 + cum + z.width/2) : 0;
      const pos = new THREE.Vector3(centerPos.x+(axis==='x'?cx:0), centerPos.y, centerPos.z+(axis==='z'?cz:0));
      const Kp = (z.Cpe < 0) ? getKpValue() : 1.0;
      const kce1 = parseFloat(document.getElementById('kce1-val')?.value) || 0.8;
      const kci1 = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;
      const Kv = S.R.Kv || 1;
      const pe = S.R.qz * z.Cpe * Kp;
      const pi1 = S.R.qz * faceData.Cp_i * kci1 * Kv;
      const p = pNetTable55Case1(pe, kce1, pi1);
      const zoneData = {key:faceKey, name:faceData.name+' ('+z.dist+')', Cp_e:z.Cpe, Cp_i:faceData.Cp_i, p: -Math.abs(p), area:z.area, force:Math.abs(p*z.area), clause:faceData.clause};
      const m = new THREE.Mesh(new THREE.PlaneGeometry(z.width, wallH), wallMat(faceKey, 0xcc8844));
      m.position.copy(pos); m.rotation.y = yRot;
      m.castShadow=true; m.receiveShadow=true;
      faceMap.set(m.uuid, zoneData);
      grp.add(m);
    });
  }
  let cum = 0;
  zones.forEach((z,i)=>{
    const stripW = z.width;
    const stripCenter = (windwardSign > 0 ? 1 : -1) * (wallLength/2 - cum - stripW/2);
    const pos = new THREE.Vector3(centerPos.x, centerPos.y, centerPos.z);
    if(axis==='x') pos.x += stripCenter; else pos.z += stripCenter;
    const zoneData = {key:faceKey, name:faceData.name+' ('+z.dist+')', Cp_e:z.Cpe, Cp_i:faceData.Cp_i, p:z.p, area:z.area, force:z.force, clause:faceData.clause};
    const col = S.showHeatmap ? heatCol(z.p) : 0xcc8844;
    const mat = new THREE.MeshStandardMaterial({color:col, roughness:.6, metalness:.1, transparent:S.viewMode==='transparent'||S.showHeatmap, opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1, wireframe:S.viewMode==='wireframe', side:THREE.DoubleSide});
    const m = new THREE.Mesh(new THREE.PlaneGeometry(stripW, wallH), mat);
    m.position.copy(pos); m.rotation.y = yRot;
    m.castShadow=true; m.receiveShadow=true;
    faceMap.set(m.uuid, zoneData);
    grp.add(m);
    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(m.geometry),new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25}));
      e.position.copy(pos); e.rotation.y = yRot; grp.add(e);
    }
    cum += stripW;
  });
}

// Table 5.2(C) — Sidewall Cp,e zones in 3D: subdivide sidewall into zones by distance from windward edge
// centerPos: face center; alongAxis: 'x'|'z'; alongSign: +1 = windward at positive end
function addSidewallCpZones(faceKey, wallLength, wallH, centerPos, alongAxis, alongSign, yRot, grp){
  grp=grp||grpBuild;
  const faceData = S.R?.faces?.[faceKey];
  if(!faceData) return;
  let zones = faceData.zones;
  if(!zones || zones.length===0){
    zones = sidewallCpZones(wallLength, S.R?.h||S.height, wallH);
  }
  const half = wallLength / 2;
  let cum = 0;
  zones.forEach(z => {
    const stripW = Math.max(0.001, z.width);
    const stripCenter = half - cum - stripW/2;
    cum += stripW;
    const pos = centerPos.clone();
    if(alongAxis==='z') pos.z += alongSign * stripCenter;
    else pos.x += alongSign * stripCenter;
    const geo = new THREE.PlaneGeometry(stripW, wallH);
    const zoneFaceData = {name:faceData.name+' ('+z.dist+')', key:faceKey, Cp_e:z.Cpe, Cp_i:faceData.Cp_i,
      p:z.p, area:z.area, force:z.force, clause:faceData.clause, Kl:1.0};
    const col = S.showHeatmap ? heatCol(z.p) : 0xcc8844;
    const mat = new THREE.MeshStandardMaterial({
      color:col, roughness:.6, metalness:.1,
      transparent:S.viewMode==='transparent'||S.showHeatmap, opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe', side:THREE.DoubleSide
    });
    const m = new THREE.Mesh(geo, mat);
    m.position.copy(pos);
    m.rotation.set(0, yRot, 0);
    m.castShadow=true; m.receiveShadow=true;
    faceMap.set(m.uuid, zoneFaceData);
    grp.add(m);
    if(S.viewMode!=='wireframe'){
      const e = new THREE.LineSegments(new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25}));
      e.position.copy(pos); e.rotation.set(0,yRot,0);
      grp.add(e);
    }
  });
}

// Monoslope: roof pitch spans depth pd; sidewalls at ±pw/2 are one trapezoid each (top follows roof line).
function getMonoWallParams(){
  const w = S.width, d = S.depth;
  const pRad = (S.pitch || 0) * Math.PI / 180;
  const monoRise = S.roofType === 'monoslope' ? Math.tan(pRad) * d : 0;
  const ov = S.overhang || 0;
  return { w, d, pw: w, pd: d, monoRise, ov };
}

/**
 * Top of ±X monoslope wall at z — same sloping plane as buildMono / buildKlMonoRoof roof quad
 * (outer eaves at z = ±(d/2+ov); z is still the wall footprint in [−d/2, d/2]).
 * When ov=0 this matches h + monoRise*(d/2−z)/d.
 */
function monoSidewallTopY(z, h, d, monoRise, ov){
  ov = ov || 0;
  const span = d + 2 * ov;
  if(span < 1e-6) return h;
  return h + monoRise * (d / 2 + ov - z) / span;
}

/** Single planar quad at x=fixedX from z=zLo to z=zHi, bottom y=0, top follows monoslope. isLeft: outward normal −X. */
function monoSidewallStripGeometry(fixedX, zLo, zHi, h, d, monoRise, isLeft, ov){
  const yLo = monoSidewallTopY(zLo, h, d, monoRise, ov);
  const yHi = monoSidewallTopY(zHi, h, d, monoRise, ov);
  const positions = isLeft
    ? new Float32Array([fixedX,0,zLo, fixedX,0,zHi, fixedX,yHi,zHi, fixedX,yLo,zLo])
    : new Float32Array([fixedX,0,zHi, fixedX,0,zLo, fixedX,yLo,zLo, fixedX,yHi,zHi]);
  const g = new THREE.BufferGeometry();
  g.setAttribute('position', new THREE.Float32BufferAttribute(positions, 3));
  g.setIndex([0,1,2, 0,2,3]);
  g.computeVertexNormals();
  return g;
}

// Monoslope ±X end wall when that face is windward or leeward: one trapezoid from y=0 to roof line (not d×h + gap).
function addMonoEndWallWindLeeward(faceKey, isLeft, h, monoRise){
  const fd = S.R?.faces?.[faceKey];
  if(!fd) return;
  const { pd, ov } = getMonoWallParams();
  if(monoRise < 0.001){
    const yRot = isLeft ? -Math.PI/2 : Math.PI/2;
    return addFace(new THREE.PlaneGeometry(S.depth, h), faceKey, 0xcc8844, new THREE.Vector3(isLeft ? -S.width/2 : S.width/2, h/2, 0), new THREE.Euler(0, yRot, 0));
  }
  const halfPlaneX = S.width / 2;
  const fixedX = isLeft ? -halfPlaneX : halfPlaneX;
  const geo = monoSidewallStripGeometry(fixedX, -pd / 2, pd / 2, h, pd, monoRise, isLeft, ov);
  const m = new THREE.Mesh(geo, wallMat(faceKey, 0xcc8844));
  m.castShadow = true; m.receiveShadow = true;
  faceMap.set(m.uuid, { key: faceKey, ...fd });
  grpBuild.add(m);
  if(S.viewMode !== 'wireframe'){
    grpBuild.add(new THREE.LineSegments(new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({ color: 0, transparent: true, opacity: 0.25 })));
  }
}

// Table 5.2(C) — monoslope sidewalls: trapezoidal strips (same zone layout as rectangular case).
// wallLengthAlongZ = physical span along building Z (S.depth) for ±X walls. Calc() zone widths use effD which
// swaps with effW when wind hits the side — scale strips to wallLengthAlongZ or geometry leaves huge gaps.
function addMonoSidewallCpZonesImpl(faceKey, wallH, centerPos, windwardSign, yRot, grp, wallLengthAlongZ){
  grp = grp || grpBuild;
  const { pd, monoRise, ov } = getMonoWallParams();
  const zSpan = (wallLengthAlongZ != null && wallLengthAlongZ > 0) ? wallLengthAlongZ : pd;
  const faceData = S.R?.faces?.[faceKey];
  if(!faceData) return;
  const isLeft = Math.abs(yRot + Math.PI/2) < 0.01;
  // ±X walls must align with slab / front & back faces (S.width), not pw — monoRidgeAlongDepth swaps pw and would inset walls inside the footprint.
  const halfPlaneX = S.width / 2;
  const fixedX = isLeft ? -halfPlaneX : halfPlaneX;
  const zones = (faceData.zones && faceData.zones.length > 0) ? faceData.zones : null;
  if(!zones || zones.length === 0){
    const geo = monoSidewallStripGeometry(fixedX, -zSpan/2, zSpan/2, wallH, pd, monoRise, isLeft, ov);
    const m = new THREE.Mesh(geo, wallMat(faceKey, 0xcc8844));
    m.castShadow = true; m.receiveShadow = true;
    faceMap.set(m.uuid, { key: faceKey, ...faceData });
    grp.add(m);
    if(S.viewMode !== 'wireframe'){
      grp.add(new THREE.LineSegments(new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({ color: 0, transparent: true, opacity: 0.25 })));
    }
    return;
  }
  let sumRaw = 0;
  zones.forEach(z=>{
    const w = z.width ?? (z.area / wallH) ?? 0;
    if(w > 0) sumRaw += w;
  });
  const scale = sumRaw > 1e-6 ? zSpan / sumRaw : 1;
  const wB = S.width, dB = S.depth, hB = S.height;
  const aWall = Math.min(0.2 * Math.min(wB, dB), hB);
  const rB = hB / Math.min(wB, dB);
  const isClad = S.showPressureMap;
  const isGlaz = false;
  let cum = 0;
  zones.forEach(z=>{
    const rawW = z.width ?? (z.area / wallH) ?? 0;
    const stripW = rawW * scale;
    if(!stripW || stripW < 0.01) return;
    const stripCenterFromWindward = zSpan/2 - cum - stripW/2;
    const zC = windwardSign * stripCenterFromWindward;
    const zLo = zC - stripW/2;
    const zHi = zC + stripW/2;
    const geo = monoSidewallStripGeometry(fixedX, zLo, zHi, wallH, pd, monoRise, isLeft, ov);
    const distMid = cum + stripW/2;
    const { Kl: klStrip, klZone: klZ } = sidewallKlTable56(distMid, aWall, rB, isClad, isGlaz);
    const zoneFace = { Cp_e: z.Cpe, Cp_i: faceData.Cp_i };
    const pMod = (isClad || isGlaz) ? localPnet(zoneFace, klStrip, true) : null;
    const zoneData = { key: faceKey, name: faceData.name+' ('+z.dist+')', Cp_e: z.Cpe, Cp_i: faceData.Cp_i, p: z.p, area: z.area, force: z.force, clause: faceData.clause, Kl: klStrip, klZone: klZ, ...(pMod != null && Number.isFinite(pMod) ? { pMod } : {}) };
    const col = S.showHeatmap ? heatCol(z.p) : 0xcc8844;
    const isTr = S.viewMode === 'transparent', isWf = S.viewMode === 'wireframe';
    const mat = new THREE.MeshStandardMaterial({ color: col, roughness: 0.6, metalness: 0.1, transparent: isTr || S.showHeatmap, opacity: isTr ? (S.showHeatmap ? 0.75 : 0.4) : 1, wireframe: isWf, side: THREE.DoubleSide });
    const m = new THREE.Mesh(geo, mat);
    m.castShadow = true; m.receiveShadow = true;
    faceMap.set(m.uuid, zoneData);
    grp.add(m);
    if(S.viewMode !== 'wireframe'){
      grp.add(new THREE.LineSegments(new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({ color: 0, transparent: true, opacity: 0.25 })));
    }
    cum += stripW;
  });
  if(cum < zSpan - 0.02 && monoRise > 0.001){
    const rem = zSpan - cum;
    const stripCenterFromWindward = zSpan/2 - cum - rem/2;
    const zC = windwardSign * stripCenterFromWindward;
    const zLo = zC - rem/2;
    const zHi = zC + rem/2;
    const geo = monoSidewallStripGeometry(fixedX, zLo, zHi, wallH, pd, monoRise, isLeft, ov);
    const lz = zones[zones.length - 1];
    const zoneData = { key: faceKey, name: faceData.name+' (remainder)', Cp_e: lz.Cpe, Cp_i: faceData.Cp_i, p: faceData.p, area: faceData.area, force: faceData.force, clause: faceData.clause };
    const m = new THREE.Mesh(geo, wallMat(faceKey, 0xcc8844));
    m.castShadow = true; m.receiveShadow = true;
    faceMap.set(m.uuid, zoneData);
    grp.add(m);
    if(S.viewMode !== 'wireframe'){
      grp.add(new THREE.LineSegments(new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({ color: 0, transparent: true, opacity: 0.25 })));
    }
  }
}

// Table 5.2(C) — Add sidewall as zone strips for 3D model (Cp,e varies along wall).
// When Local Pressure Zones are on, also attach Table 5.6 Kl + local net p (Kp) per strip — same as buildKlWalls.
// faceKey, wallLength (along-wind), wallH, centerPos, axis 'x'|'z', windwardSign +1|−1, yRot (optional), grp (optional)
function addSidewallCpZones(faceKey,wallLength,wallH,centerPos,axis,windwardSign,yRot,grp){
  grp=grp||grpBuild;
  if(typeof yRot!=='number'){ grp=yRot; yRot=axis==='z'?-Math.PI/2:0; }
  if(S.roofType === 'monoslope' && axis === 'z'){
    return addMonoSidewallCpZonesImpl(faceKey, wallH, centerPos, windwardSign, yRot, grp, wallLength);
  }
  const faceData=S.R?.faces?.[faceKey];
  const zones=(faceData?.zones&&faceData.zones.length>0)?faceData.zones:null;
  if(!zones || zones.length===0){
    // Fallback: single face if no zones
    const geo=new THREE.PlaneGeometry(wallLength,wallH);
    return addFace(geo,faceKey,0xcc8844,centerPos,new THREE.Euler(0,yRot,0),grp);
  }
  const wB=S.width, dB=S.depth, hB=S.height;
  const aWall = Math.min(0.2 * Math.min(wB, dB), hB);
  const rB = hB / Math.min(wB, dB);
  const isClad = S.showPressureMap;
  const isGlaz = false;
  let cum=0;
  zones.forEach(z=>{
    const stripW=z.width ?? (z.area/wallH) ?? wallLength;
    if(!stripW||stripW<0.01)return;
    const geo=new THREE.PlaneGeometry(stripW,wallH);
    const stripCenterFromWindward=wallLength/2-cum-stripW/2;
    const pos=centerPos.clone();
    if(axis==='z')pos.z=centerPos.z+windwardSign*stripCenterFromWindward;
    else pos.x=centerPos.x+windwardSign*stripCenterFromWindward;
    const rot=new THREE.Euler(0,yRot,0);
    const distMid = cum + stripW / 2;
    const {Kl: klStrip, klZone: klZ} = sidewallKlTable56(distMid, aWall, rB, isClad, isGlaz);
    const zoneFace = {Cp_e: z.Cpe, Cp_i: faceData.Cp_i};
    const pMod = (isClad || isGlaz) ? localPnet(zoneFace, klStrip, true) : null;
    const zoneData={key:faceKey,name:faceData.name+' ('+z.dist+')',Cp_e:z.Cpe,Cp_i:faceData.Cp_i,p:z.p,area:z.area,force:z.force,clause:faceData.clause,Kl:klStrip,klZone:klZ,...(pMod!=null&&Number.isFinite(pMod)?{pMod}:{})};
    const col=S.showHeatmap?heatCol(z.p):0xcc8844;
    const isTr=S.viewMode==='transparent', isWf=S.viewMode==='wireframe';
    const mat=new THREE.MeshStandardMaterial({color:col,roughness:.6,metalness:.1,transparent:isTr||S.showHeatmap,opacity:isTr?(S.showHeatmap?.75:.4):1,wireframe:isWf,side:THREE.DoubleSide});
    const m=new THREE.Mesh(geo,mat);
    m.position.copy(pos);m.rotation.copy(rot);
    m.castShadow=true;m.receiveShadow=true;
    faceMap.set(m.uuid,zoneData);
    grp.add(m);
    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25}));
      e.position.copy(pos);e.rotation.copy(rot);
      grp.add(e);
    }
    cum+=stripW;
  });
}

function heatCol(p){
  if(!Number.isFinite(p)) return 0x8899aa;
  // Independent normalization: positive vs negative face pressures
  // Include face-level AND zone-level pressures so sidewall zone strips color correctly
  const faces = S.R?.faces || {};
  let maxPos = 0.3, maxNeg = 0.3; // minimum ranges to avoid division by zero
  for(const k in faces){
    const f = faces[k];
    if(!f) continue;
    if(f.p !== undefined){
      if(f.p > maxPos) maxPos = f.p;
      if(f.p < -maxNeg) maxNeg = -f.p;
    }
    if(f.zones && Array.isArray(f.zones)){
      f.zones.forEach(z=>{
        if(z.p !== undefined){
          if(z.p > maxPos) maxPos = z.p;
          if(z.p < -maxNeg) maxNeg = -z.p;
        }
      });
    }
  }
  let t;
  if(p >= 0){ t = Math.min(1, p / maxPos); }
  else { t = Math.max(-1, p / maxNeg); }

  // Smooth vivid gradient: deep-blue → blue → cyan → white → yellow → orange → red
  let r,g,b;
  if(t < -0.6){
    const s = (-t - 0.6) / 0.4; // 0→1 as t goes -0.6→-1
    r = Math.round(30 - s*10);   g = Math.round(60 - s*30);   b = Math.round(200 + s*20);
  } else if(t < -0.25){
    const s = (-t - 0.25) / 0.35;
    r = Math.round(40 - s*10);   g = Math.round(140 - s*80);  b = Math.round(230 - s*30);
  } else if(t < 0){
    const s = -t / 0.25;
    r = Math.round(235 - s*195); g = Math.round(240 - s*100); b = Math.round(245 - s*15);
  } else if(t < 0.25){
    const s = t / 0.25;
    r = Math.round(235 + s*20);  g = Math.round(240 - s*50);  b = Math.round(245 - s*205);
  } else if(t < 0.6){
    const s = (t - 0.25) / 0.35;
    r = 255;                      g = Math.round(190 - s*100); b = Math.round(40 - s*30);
  } else {
    const s = (t - 0.6) / 0.4;
    r = Math.round(255 - s*50);  g = Math.round(90 - s*65);   b = Math.round(10 - s*10);
  }
  return(r<<16)|(g<<8)|b;
}

function spriteText(txt,pos,col,sc){
  const lines = String(txt).split('\n');
  const lineCount = lines.length;
  const c=document.createElement('canvas'),x=c.getContext('2d');
  const fontSize = 48;
  const lineH = fontSize * 1.3;
  c.width=512;c.height=Math.round(lineH * lineCount + 24);
  x.fillStyle='rgba(0,0,0,.7)';x.beginPath();
  if(x.roundRect)x.roundRect(4,4,c.width-8,c.height-8,10); else x.fillRect(4,4,c.width-8,c.height-8);
  x.fill();
  x.font='bold '+fontSize+'px Segoe UI,system-ui,sans-serif';x.textAlign='center';x.textBaseline='middle';
  x.fillStyle='#'+col.toString(16).padStart(6,'0');
  lines.forEach((ln,i)=>{ x.fillText(ln, c.width/2, lineH*(i+0.5)+12); });
  const t=new THREE.CanvasTexture(c);
  const aspect = c.width / c.height;
  const sp=new THREE.Sprite(new THREE.SpriteMaterial({map:t,transparent:true,depthTest:false}));
  sp.position.copy(pos);sp.scale.set(sc*5, sc*5/aspect, 1);
  return sp;
}

// ── WALLS ──
// Compute which physical face the wind hits based on building orientation vs wind direction
function getWindFaceMap(){
  // relWind: compass bearing wind comes FROM, relative to building front
  const relWind = ((S.windAngle - S.mapBuildingAngle) % 360 + 360) % 360;
  // Determine which quadrant the wind blow direction falls in
  // front(+z), right(+x), back(-z), left(-x) in building local space
  if(relWind >= 315 || relWind < 45){
    // Wind blows at front face
    return {front:'windward', back:'leeward', left:'sidewall1', right:'sidewall2'};
  } else if(relWind >= 45 && relWind < 135){
    // Wind blows at right face
    return {right:'windward', left:'leeward', front:'sidewall2', back:'sidewall1'};
  } else if(relWind >= 135 && relWind < 225){
    // Wind blows at back face
    return {back:'windward', front:'leeward', right:'sidewall1', left:'sidewall2'};
  } else {
    // Wind blows at left face (225-315)
    return {left:'windward', right:'leeward', back:'sidewall2', front:'sidewall1'};
  }
}

function buildWalls(){
  const w=S.width,d=S.depth,h=S.height;
  const fm = getWindFaceMap();
  const { monoRise: monoBackRise } = getMonoWallParams();
  const backWallH = S.roofType === 'monoslope' ? h + monoBackRise : h;
  const backWallCy = backWallH / 2;

  const slab=new THREE.Mesh(new THREE.BoxGeometry(w+.4,.3,d+.4),new THREE.MeshStandardMaterial({color:0x555555,roughness:.8}));
  slab.position.y=-.15;slab.castShadow=true;slab.receiveShadow=true;grpBuild.add(slab);

  if(S.showPressureMap){
    buildKlWalls(w,d,h);
  } else {
    // Windward/leeward: single face; Sidewalls: Table 5.2(C) zone strips (Cp,e varies 0→1h: -0.65, 1h→2h: -0.5, 2h→3h: -0.3, >3h: -0.2)
    // Gable walls (±X): strips run along Z. First strip must start at the vertical edge shared with the WINDWARD face.
    // Front (+Z) windward → that edge is at z = +d/2  → offset strips from +Z toward −Z  → windwardSign = +1
    // Back (−Z) windward → that edge is at z = −d/2 → offset strips from −Z toward +Z  → windwardSign = −1
    const swWindwardZ = fm.front === 'windward' ? 1 : -1;
    // Table 5.2(C): first strip must start at the windward horizontal edge of the wall.
    // Front/back walls run along ±X; windward corner is +X if right is windward, −X if left is windward.
    const swWindwardX = fm.right === 'windward' ? 1 : -1;
    if(fm.front==='sidewall1'||fm.front==='sidewall2'){
      addSidewallCpZones(fm.front, w, h, new THREE.Vector3(0,h/2,d/2), 'x', swWindwardX, 0);
    } else { addFace(new THREE.PlaneGeometry(w,h), fm.front, 0x4488cc, new THREE.Vector3(0,h/2,d/2), new THREE.Euler(0,0,0)); }
    if(fm.back==='sidewall1'||fm.back==='sidewall2'){
      addSidewallCpZones(fm.back, w, backWallH, new THREE.Vector3(0,backWallCy,-d/2), 'x', swWindwardX, Math.PI);
    } else { addFace(new THREE.PlaneGeometry(w, backWallH), fm.back, 0x44aa66, new THREE.Vector3(0,backWallCy,-d/2), new THREE.Euler(0,Math.PI,0)); }
    if(fm.left==='sidewall1'||fm.left==='sidewall2'){
      addSidewallCpZones(fm.left, d, h, new THREE.Vector3(-w/2,h/2,0), 'z', swWindwardZ, -Math.PI/2);
    } else if(S.roofType === 'monoslope' && (fm.left === 'windward' || fm.left === 'leeward')){
      addMonoEndWallWindLeeward(fm.left, true, h, monoBackRise);
    } else { addFace(new THREE.PlaneGeometry(d,h), fm.left, 0xcc8844, new THREE.Vector3(-w/2,h/2,0), new THREE.Euler(0,-Math.PI/2,0)); }
    if(fm.right==='sidewall1'||fm.right==='sidewall2'){
      addSidewallCpZones(fm.right, d, h, new THREE.Vector3(w/2,h/2,0), 'z', swWindwardZ, Math.PI/2);
    } else if(S.roofType === 'monoslope' && (fm.right === 'windward' || fm.right === 'leeward')){
      addMonoEndWallWindLeeward(fm.right, false, h, monoBackRise);
    } else { addFace(new THREE.PlaneGeometry(d,h), fm.right, 0xcc8844, new THREE.Vector3(w/2,h/2,0), new THREE.Euler(0,Math.PI/2,0)); }
  }

  addOpenings(w,d,h);
}

// Local net pressure (Kl mesh): same Table 5.5 form as tabulated pressures — p_net = p_e K_c,e − p_i
// p_e,local = q_z × C_p,e × K_l × K_p; p_i = q_z × C_p,i × K_c,i × K_v
function localPnet(faceData, Kl, applyKp){
  const qz = S.R.qz, Kv = S.R.Kv || 1;
  const kce1 = parseFloat(document.getElementById('kce1-val')?.value) || 0.8;
  const kci1 = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;
  const cpi = faceData && Number.isFinite(faceData.Cp_i) ? faceData.Cp_i : 0;
  const terms = localPressureTerms(faceData, Kl, applyKp, qz, cpi);
  return pNetTable55Case1(terms.pe, kce1, terms.pi);
}

function localPressureTerms(faceData, Kl, applyKp, qz, cpi){
  const kci1 = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;
  const Kv = S.R.Kv || 1;
  let Kp = 1.0;
  if(applyKp && faceData.Cp_e < 0){
    Kp = getKpValue();
  }
  const peUniform = (faceData.peExternalAvg != null && Number.isFinite(faceData.peExternalAvg))
    ? faceData.peExternalAvg
    : qz * faceData.Cp_e;
  const pe = peUniform * Kl * Kp;
  const pi = qz * cpi * kci1 * Kv;
  return { pe, pi, Kp };
}

// Tooltip: infer Kl from Table 5.6 zone label when mesh has klZone but no numeric Kl
function klFromKlZoneName(zoneName){
  if(!zoneName || typeof zoneName !== 'string') return null;
  const z = zoneName.trim();
  if(z.startsWith('RC')) return 3.0;
  if(z.startsWith('RA4')) return 2.0;
  if(z.startsWith('RA3')) return 1.5;
  if(z.startsWith('RA2')) return 2.0;
  if(z.startsWith('RA1')) return 1.5;
  if(z.startsWith('MR')) return 1.0;
  if(z.startsWith('WA')) return 1.5;
  if(z.startsWith('SA5')) return 3.0;
  if(z.startsWith('SA4')) return 2.0;
  if(z.startsWith('SA3')) return 1.5;
  if(z.startsWith('SA2')) return 2.0;
  if(z.startsWith('SA1')) return 1.5;
  return null;
}

// Table 5.6 — sidewall Kl vs distance from windward corner along the wall (same breakpoints as buildKlWalls swZones)
function sidewallKlTable56(distFromWindwardCorner, a, r, isCladding, isGlazing){
  if(!isCladding && !isGlazing){
    return {Kl: 1.0, klZone: 'Other'};
  }
  const d = Math.max(0, distFromWindwardCorner);
  if(r <= 1){
    if(d < 0.5 * a) return {Kl: 2.0, klZone: 'SA2'};
    if(d < a) return {Kl: 1.5, klZone: 'SA1'};
    return {Kl: 1.0, klZone: 'Other'};
  }
  if(d < 0.5 * a) return {Kl: 3.0, klZone: 'SA5'};
  if(d < a) return {Kl: 2.0, klZone: 'SA4'};
  return {Kl: 1.5, klZone: 'SA3'};
}

// Get the uniform Kp value (1.0 if solid/auto, else user-selected)
function getKpValue(){
  if(S.Kp === '1.0') return 1.0;
  if(S.Kp === 'auto') return 1.0; // auto uses zone-based Kp, handled separately
  return parseFloat(S.Kp) || 1.0;
}

// Table 5.8: Kp zones based on distance from windward edge
// da = along-wind depth of the surface
function getKpZones(da){
  if(S.Kp === '1.0') return [{dist:Infinity, Kp:1.0}];
  if(S.Kp !== 'auto') return [{dist:Infinity, Kp:parseFloat(S.Kp)||1.0}];
  // Auto: Table 5.8 zones
  return [
    {dist:0.2*da, Kp:0.9},
    {dist:0.4*da, Kp:0.8},
    {dist:0.8*da, Kp:0.7},
    {dist:1.0*da, Kp:0.8},
    {dist:Infinity, Kp:1.0}  // beyond da — no permeable surface
  ];
}

// Get Kp at a specific fractional distance (0=windward edge, 1=leeward edge) along da
function getKpAtFraction(frac){
  if(S.Kp === '1.0') return 1.0;
  if(S.Kp !== 'auto') return parseFloat(S.Kp) || 1.0;
  // Table 5.8
  if(frac <= 0.2) return 0.9;
  if(frac <= 0.4) return 0.8;
  if(frac <= 0.8) return 0.7;
  if(frac <= 1.0) return 0.8;
  return 1.0;
}

// Zone colour based on Kl + Table 5.6 zone name (Kl=1.5: WA windward wall vs RA roof vs SA sidewall).
function klZoneColor(Kl, zoneName){
  const zn = zoneName && typeof zoneName === 'string' ? zoneName : '';
  if(Kl >= 2.5) return 0xDD2222; // red — RC1 corners (Kl=3.0)
  if(Kl >= 1.8) return 0xEECC00; // yellow — RA2/SA zones (Kl=2.0)
  if(Kl >= 1.3){
    if(zn.startsWith('WA')) return 0x22BB44; // windward wall WA1 — bright green
    if(zn.startsWith('RA')) return 0x22BB44; // roof RA* (e.g. RA1 Kl=1.5) — green (same family as WA1)
    if(/^RC/.test(zn)) return 0x2a9d8f; // roof RC* — teal (distinct from RA when Kl in this band)
    if(zn.startsWith('SA') || zn === 'gap-fill') return 0x4caf50; // sidewall SA* — lime green
    return 0x22BB44;
  }
  return 0x2288DD; // interior/MR (Kl=1.0)
}

// NZS 1170.2 Figure 5.3 / Table 5.6 zone reference name
// Prefer explicit strip.name when available; this is the fallback
function klZoneName(faceKey, Kl){
  const isRoof = faceKey.startsWith('roof');
  if(isRoof){
    if(Kl >= 2.5) return 'RC';   // caller should override with RC1 or RC2
    if(Kl >= 1.8) return 'RA2';
    if(Kl >= 1.3) return 'RA1';
    return 'MR';
  }
  if(faceKey === 'windward') return Kl >= 1.3 ? 'WA1' : 'Other';
  if(faceKey === 'leeward') return 'Other';
  return 'Other';
}

// ── Local Pressure Zones: subdivide walls into Kl zones ──
// Direction-aware: uses getWindFaceMap() to assign windward/leeward/sidewall
// roles to the correct physical faces based on the current wind angle.
function buildKlWalls(w,d,h){
  const fm = getWindFaceMap();

  // Clause 5.4.4 — dimension a (different for walls vs roofs)
  const a_wall = Math.min(0.2 * Math.min(w, d), h);
  const a_roof = (h/w >= 0.2 || h/d >= 0.2) ? 0.2 * Math.min(w, d) : 2 * h;
  const a = a_wall;
  // Building aspect ratio r = h / min(b, d) per Table 5.6 Note 2
  const r = h / Math.min(w, d);
  const isCladding = S.showPressureMap;
  const isGlazing  = false;

  // ─── Windward wall zones (WA1) ───
  const wwZones = (isCladding || isGlazing)
    ? [{dist:Infinity, Kl:1.5, name:'WA1'}]
    : [{dist:Infinity, Kl:1.0, name:'Other'}];

  // ─── Leeward wall — no local Kl amplification ───
  const lwZones = [{dist:Infinity, Kl:1.0, name:'Other'}];

  // ─── Side wall zones — depend on building aspect ratio r ───
  let swZones;
  if(r <= 1){
    if(isGlazing || isCladding){
      swZones = [
        {dist:0.5*a, Kl:2.0, name:'SA2'},
        {dist:a,     Kl:1.5, name:'SA1'},
        {dist:Infinity, Kl:1.0, name:'Other'}
      ];
    } else {
      swZones = [{dist:Infinity, Kl:1.0, name:'Other'}];
    }
  } else {
    if(isGlazing || isCladding){
      swZones = [
        {dist:0.5*a, Kl:3.0, name:'SA5'},
        {dist:a,     Kl:2.0, name:'SA4'},
        {dist:Infinity, Kl:1.5, name:'SA3'}
      ];
    } else {
      swZones = [{dist:Infinity, Kl:1.0, name:'Other'}];
    }
  }

  const wwFaceData = S.R.faces['windward'];
  const wa1Kl = (isCladding || isGlazing) ? 1.5 : 1.0;

  // Physical face definitions:
  // front(+z): width=w, strips along X   back(-z): width=w, strips along X
  // left(-x):  width=d, strips along Z   right(+x): width=d, strips along Z
  // highNeighbor = which adjacent face is at the "high" end of strip axis
  const physFaces = [
    {name:'front', wallW:w, fixed: d/2,  yRot:0,           zStrip:false, highNeighbor:'right'},
    {name:'back',  wallW:w, fixed:-d/2,  yRot:Math.PI,     zStrip:false, highNeighbor:'right'},
    {name:'left',  wallW:d, fixed:-w/2,  yRot:-Math.PI/2,  zStrip:true,  highNeighbor:'front'},
    {name:'right', wallW:d, fixed: w/2,  yRot:Math.PI/2,   zStrip:true,  highNeighbor:'front'},
  ];

  const { monoRise: monoBackRiseKl } = getMonoWallParams();
  physFaces.forEach(pf => {
    const role = fm[pf.name]; // 'windward','leeward','sidewall1','sidewall2'
    const wallHMono = S.roofType === 'monoslope' && pf.name === 'back' ? h + monoBackRiseKl : h;
    if(role === 'windward' || role === 'leeward'){
      const zones = role === 'windward' ? wwZones : lwZones;
      if(S.roofType === 'monoslope' && (pf.name === 'left' || pf.name === 'right')){
        const { pd, monoRise: mr, ov } = getMonoWallParams();
        if(mr > 0.001){
          buildKlMonoEndWallTrapezoid(pf, role, zones, h, mr, pd, ov);
        } else {
          buildKlFace(pf.wallW, wallHMono, pf.fixed, pf.yRot, role, zones, a,
                      false, true, pf.zStrip, null, 0);
        }
      } else {
        buildKlFace(pf.wallW, wallHMono, pf.fixed, pf.yRot, role, zones, a,
                    false, true, pf.zStrip, null, 0);
      }
    } else {
      // Side wall — windward corner is the end adjacent to the windward face
      const fromHighEnd = (fm[pf.highNeighbor] === 'windward');
      let wallHKl = h;
      if(S.roofType === 'monoslope' && pf.name === 'back'){
        wallHKl = h + monoBackRiseKl;
      }
      if(S.roofType === 'monoslope' && (pf.name === 'left' || pf.name === 'right')){
        const { pd, monoRise, ov } = getMonoWallParams();
        buildKlMonoSidewallFace(pf, fm, swZones, fromHighEnd, wwFaceData, wa1Kl, pd, h, monoRise, ov);
      } else {
        buildKlFace(pf.wallW, wallHKl, pf.fixed, pf.yRot, role, swZones, a,
                    true, fromHighEnd, pf.zStrip, wwFaceData, wa1Kl);
      }
    }
  });
}

// Unified Kl face builder — handles both windward/leeward (symmetric) and
// side walls (one-sided from windward corner).
// isZStrip: true for left/right faces (strips vary along Z), false for front/back (X)
// fromHighEnd: for one-sided strips, whether windward corner is at high end
function buildKlFace(wallW, wallH, fixedCoord, yRot, faceKey, zones, a,
                     isOneSided, fromHighEnd, isZStrip, wwFaceData, wa1Kl){
  const F = S.R.faces;
  const faceData = F[faceKey];
  if(!faceData) return;

  const strips = isOneSided
    ? getKlStripsOneSided(wallW, zones, fromHighEnd)
    : getKlStrips(wallW, zones);

  strips.forEach(strip => {
    const stripW = strip.x2 - strip.x1;
    if(stripW < 0.01) return;
    const geo = new THREE.PlaneGeometry(stripW, wallH);
    const stripCenter = (strip.x1 + strip.x2) / 2 - wallW/2;

    const klP = localPnet(faceData, strip.Kl, isOneSided);
    const zoneName = isOneSided
      ? (strip.name || 'Other')
      : (strip.name || klZoneName(faceKey, strip.Kl));
    const col = klZoneColor(strip.Kl, zoneName);
    const isTr = S.viewMode==='transparent', isWf = S.viewMode==='wireframe';
    const mat = new THREE.MeshStandardMaterial({
      color: col, roughness:.6, metalness:.1,
      transparent: isTr, opacity: isTr?.75:1,
      wireframe: isWf, side: THREE.DoubleSide,
      polygonOffset: true, polygonOffsetFactor: 1, polygonOffsetUnits: 1
    });

    const mesh = new THREE.Mesh(geo, mat);
    const faceNorm = new THREE.Vector3(0,0,1).applyEuler(new THREE.Euler(0, yRot, 0));
    // Position: strip offset along its axis, fixed on the perpendicular axis
    const pos = isZStrip
      ? new THREE.Vector3(fixedCoord, wallH/2, stripCenter)
      : new THREE.Vector3(stripCenter, wallH/2, fixedCoord);
    pos.add(faceNorm.clone().multiplyScalar(0.005));
    const rot = new THREE.Euler(0, yRot, 0);
    mesh.position.copy(pos); mesh.rotation.copy(rot);
    mesh.castShadow=true; mesh.receiveShadow=true;
    mesh.renderOrder = 1;
    faceMap.set(mesh.uuid, {key:faceKey, ...faceData, Kl:strip.Kl, klZone:zoneName, pMod:klP});
    grpBuild.add(mesh);

    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),
        new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25,depthWrite:false}));
      e.position.copy(pos); e.rotation.copy(rot);
      e.renderOrder = 2;
      grpBuild.add(e);
    }

    const labelPos = pos.clone();
    const lNorm = new THREE.Vector3(0,0,1).applyEuler(rot);
    labelPos.add(lNorm.multiplyScalar(0.3));
    const labelText = zoneName + '\n' + klP.toFixed(2) + ' kPa';
    grpBuild.add(spriteText(labelText, labelPos, 0xffffff, 0.4));
  });
}

// Local Pressure Zones: monoslope ±X when windward or leeward — trapezoid to roof (not vertical d×h rectangle).
function buildKlMonoEndWallTrapezoid(pf, faceKey, zones, h, monoRise, pd, ov){
  const F = S.R.faces;
  const faceData = F[faceKey];
  if(!faceData) return;

  const strips = getKlStrips(pd, zones);
  const isLeft = pf.name === 'left';
  const halfPlaneX = S.width / 2;
  const fixedX = isLeft ? -halfPlaneX : halfPlaneX;

  strips.forEach(strip => {
    const stripW = strip.x2 - strip.x1;
    if(stripW < 1e-6) return;
    const zLo = strip.x1 - pd / 2;
    const zHi = strip.x2 - pd / 2;
    const geo = monoSidewallStripGeometry(fixedX, zLo, zHi, h, pd, monoRise, isLeft, ov);

    const klP = localPnet(faceData, strip.Kl, false);
    const zoneName = strip.name || klZoneName(faceKey, strip.Kl);
    const col = klZoneColor(strip.Kl, zoneName);
    const isTr = S.viewMode==='transparent', isWf = S.viewMode==='wireframe';
    const mat = new THREE.MeshStandardMaterial({
      color: col, roughness:.6, metalness:.1,
      transparent: isTr, opacity: isTr?.75:1,
      wireframe: isWf, side: THREE.DoubleSide,
      polygonOffset: true, polygonOffsetFactor: -2, polygonOffsetUnits: -2
    });
    const mesh = new THREE.Mesh(geo, mat);
    const xOut = isLeft ? -0.04 : 0.04;
    mesh.position.x = xOut;
    mesh.castShadow=true; mesh.receiveShadow=true;
    mesh.renderOrder = 2;
    faceMap.set(mesh.uuid, {key:faceKey, ...faceData, Kl:strip.Kl, klZone:zoneName, pMod:klP});
    grpBuild.add(mesh);

    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),
        new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25,depthWrite:false}));
      e.position.x = xOut;
      e.renderOrder = 3;
      grpBuild.add(e);
    }

    const yLoT = monoSidewallTopY(zLo, h, pd, monoRise, ov);
    const yHiT = monoSidewallTopY(zHi, h, pd, monoRise, ov);
    const labelPos = new THREE.Vector3(fixedX + xOut + (isLeft ? -0.35 : 0.35), (yLoT + yHiT) / 4, (zLo + zHi) / 2);
    const lNorm = new THREE.Vector3(isLeft ? -1 : 1, 0, 0);
    labelPos.add(lNorm.multiplyScalar(0.3));
    grpBuild.add(spriteText(zoneName + '\n' + klP.toFixed(2) + ' kPa', labelPos, 0xffffff, 0.4));
  });
}

// Local Pressure Zones: monoslope sidewalls — trapezoidal strips (same layout as addMonoSidewallCpZonesImpl).
function buildKlMonoSidewallFace(pf, fm, swZones, fromHighEnd, wwFaceData, wa1Kl, pd, h, monoRise, ov){
  const F = S.R.faces;
  const faceKey = fm[pf.name];
  const faceData = F[faceKey];
  if(!faceData) return;

  // Strips 0..pd map to z ∈ [−pd/2, pd/2]; must match monoSidewallTopY (roof plane incl. overhang) or quads bow-tie / render black.
  const strips = getKlStripsOneSided(pd, swZones, fromHighEnd);
  const isLeft = pf.name === 'left';
  const halfPlaneX = S.width / 2;
  const fixedX = isLeft ? -halfPlaneX : halfPlaneX;

  strips.forEach(strip => {
    const stripW = strip.x2 - strip.x1;
    if(stripW < 1e-6) return;
    const zLo = strip.x1 - pd/2;
    const zHi = strip.x2 - pd/2;
    const geo = monoSidewallStripGeometry(fixedX, zLo, zHi, h, pd, monoRise, isLeft, ov);

    const klP = localPnet(faceData, strip.Kl, true);
    const zoneName = 'WA1 & ' + (strip.name || 'Other');
    const col = klZoneColor(strip.Kl, zoneName);
    const isTr = S.viewMode==='transparent', isWf = S.viewMode==='wireframe';
    const mat = new THREE.MeshStandardMaterial({
      color: col, roughness:.6, metalness:.1,
      transparent: isTr, opacity: isTr?.75:1,
      wireframe: isWf, side: THREE.DoubleSide,
      polygonOffset: true, polygonOffsetFactor: -2, polygonOffsetUnits: -2
    });
    const mesh = new THREE.Mesh(geo, mat);
    const xOut = isLeft ? -0.04 : 0.04;
    mesh.position.x = xOut;
    mesh.castShadow=true; mesh.receiveShadow=true;
    mesh.renderOrder = 2;
    faceMap.set(mesh.uuid, {key:faceKey, ...faceData, Kl:strip.Kl, klZone:zoneName, pMod:klP});
    grpBuild.add(mesh);

    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),
        new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25,depthWrite:false}));
      e.position.x = xOut;
      e.renderOrder = 3;
      grpBuild.add(e);
    }

    const yLoT = monoSidewallTopY(zLo, h, pd, monoRise, ov);
    const yHiT = monoSidewallTopY(zHi, h, pd, monoRise, ov);
    const labelPos = new THREE.Vector3(fixedX + xOut + (isLeft ? -0.35 : 0.35), (yLoT + yHiT) / 4, (zLo + zHi) / 2);
    const lNorm = new THREE.Vector3(isLeft ? -1 : 1, 0, 0);
    labelPos.add(lNorm.multiplyScalar(0.3));
    grpBuild.add(spriteText(zoneName + '\n' + klP.toFixed(2) + ' kPa', labelPos, 0xffffff, 0.4));
  });
}

// Symmetric strips — for windward/leeward walls (zones from both edges)
function getKlStrips(totalW, zones){
  const strips = [];
  const halfW = totalW / 2;

  if(zones.length <= 1 || zones[0].dist >= halfW){
    strips.push({x1:0, x2:totalW, Kl:zones[0].Kl, name:zones[0].name||'', label:'full'});
    return strips;
  }

  let leftEdge = 0;
  for(let i=0; i<zones.length; i++){
    const zDist = Math.min(zones[i].dist, halfW);
    if(zDist > leftEdge){
      strips.push({x1:leftEdge, x2:zDist, Kl:zones[i].Kl, name:zones[i].name||'', label:'edge'});
      strips.push({x1:totalW-zDist, x2:totalW-leftEdge, Kl:zones[i].Kl, name:zones[i].name||'', label:'edge'});
      leftEdge = zDist;
    }
    if(zDist >= halfW) break;
  }

  if(leftEdge < halfW){
    const last = zones[zones.length-1];
    strips.push({x1:leftEdge, x2:totalW-leftEdge, Kl:last.Kl, name:last.name||'', label:'center'});
  }

  return strips;
}

// One-sided strips — for side walls (zones from windward corner)
// fromHighEnd: true = windward corner at x=totalW end, false = at x=0 end
function fillGapsInKlOneSidedStrips(strips, totalW, zones){
  const last = zones && zones.length ? zones[zones.length - 1] : null;
  const fillKl = last ? last.Kl : 1.0;
  const fillName = last ? (last.name || 'Other') : 'Other';
  if(!totalW || totalW <= 0) return strips;
  if(!strips.length){
    return [{ x1: 0, x2: totalW, Kl: fillKl, name: fillName, label: 'gap-fill' }];
  }
  const sorted = strips.filter(s => s.x2 > s.x1 + 1e-12).sort((a, b) => a.x1 - b.x1);
  const out = [];
  let expected = 0;
  for(const s of sorted){
    if(s.x1 > expected + 1e-4){
      out.push({ x1: expected, x2: s.x1, Kl: fillKl, name: fillName, label: 'gap-fill' });
    }
    out.push(s);
    expected = Math.max(expected, s.x2);
  }
  if(expected < totalW - 1e-4){
    out.push({ x1: expected, x2: totalW, Kl: fillKl, name: fillName, label: 'gap-fill' });
  }
  return out;
}

function getKlStripsOneSided(totalW, zones, fromHighEnd){
  const strips = [];
  if(fromHighEnd !== false){
    // Zones from high end (x=totalW toward x=0)
    let cursor = totalW;
    for(let i = 0; i < zones.length; i++){
      const leftBound = Math.max(totalW - zones[i].dist, 0);
      if(leftBound < cursor){
        strips.push({x1:leftBound, x2:cursor, Kl:zones[i].Kl, name:zones[i].name||'', label:zones[i].name||''});
        cursor = leftBound;
      }
      if(cursor <= 0) break;
    }
  } else {
    // Zones from low end (x=0 toward x=totalW)
    let cursor = 0;
    for(let i = 0; i < zones.length; i++){
      const rightBound = Math.min(zones[i].dist, totalW);
      if(rightBound > cursor){
        strips.push({x1:cursor, x2:rightBound, Kl:zones[i].Kl, name:zones[i].name||'', label:zones[i].name||''});
        cursor = rightBound;
      }
      if(cursor >= totalW) break;
    }
  }
  return fillGapsInKlOneSidedStrips(strips, totalW, zones);
}

function getKlVerticalStrips(totalH, zones){
  const strips = [];
  if(zones.length <= 1) return strips;

  // Top edge zone: from (totalH - dist) to totalH
  let topEdge = totalH;
  for(let i=0; i<zones.length-1; i++){
    const zDist = Math.min(zones[i].dist, totalH);
    if(zDist < totalH){
      strips.push({y1:totalH-zDist, y2:totalH, Kl:zones[i].Kl, label:'top-'+zones[i].Kl});
      break;
    }
  }
  return strips;
}

function addOpenings(w,d,h){
  const oc=0x112244;
  const fm = getWindFaceMap();

  // Find physical positions for each wind role
  // windward/leeward openings go on whichever physical face has that role
  const roleFace = {}; // role → {pos,rot,faceW}
  roleFace[fm.front] = {z: d/2+.02, axis:'z', faceW:w};
  roleFace[fm.back]  = {z:-d/2-.02, axis:'z', faceW:w};
  roleFace[fm.left]  = {x:-w/2-.02, axis:'x', faceW:d};
  roleFace[fm.right] = {x: w/2+.02, axis:'x', faceW:d};

  const wwFace = roleFace['windward'];
  const lwFace = roleFace['leeward'];
  const sw1Face = roleFace['sidewall1'];
  const sw2Face = roleFace['sidewall2'];

  function placeWins(face, openPct, maxWins, winFrac, hFrac){
    if(openPct <= 10) return;
    const nw = Math.min(Math.floor(openPct/10), maxWins);
    const ww = face.faceW * winFrac, wh = h * hFrac;
    for(let i=0;i<nw;i++){
      const m = new THREE.Mesh(new THREE.PlaneGeometry(ww,wh), new THREE.MeshStandardMaterial({color:oc,transparent:true,opacity:.7,side:THREE.DoubleSide}));
      const frac = (i+1)/(nw+1);
      if(face.axis === 'z'){
        m.position.set(-face.faceW/2 + face.faceW*frac, h*.55, face.z);
      } else {
        m.position.set(face.x, h*.55, -face.faceW/2 + face.faceW*frac);
        m.rotation.y = Math.PI/2;
      }
      grpBuild.add(m);
    }
  }

  placeWins(wwFace, S.openWW, 5, .12, .4);
  placeWins(lwFace, S.openLW, 5, .12, .4);
  placeWins(sw1Face, S.openSW, 4, .12, .35);
  placeWins(sw2Face, S.openSW, 4, .12, .35);
}

// ── ROOF ──
function buildRoof(){
  const w=S.width,d=S.depth,h=S.height,ov=S.overhang;
  const pRad=S.pitch*Math.PI/180, rh=S.roofType!=='flat'?Math.tan(pRad)*d/2:0;

  if(S.roofType==='flat'){
    if(S.showPressureMap){
      buildKlFlatRoof(w,d,h,ov);
    } else {
      const F=S.R.faces;
      if(F.roof_ww?.zones?.length){
        addMonoOrFlatCrosswindStripsWallU(d,h,0,ov,w,F);
      } else {
        addFace(new THREE.PlaneGeometry(w+ov*2,d+ov*2),'roof_ww',0x886644,
          new THREE.Vector3(0,h,0),new THREE.Euler(-Math.PI/2,0,0));
      }
    }
  } else if(S.roofType==='gable'){
    if(S.showPressureMap){ buildKlGableRoof(w,d,h,rh,ov,pRad); }
    else { buildGable(w,d,h,rh,ov,pRad); }
  } else if(S.roofType==='hip'){
    if(S.showPressureMap){ buildKlHipRoof(w,d,h,rh,ov); }
    else { buildHip(w,d,h,rh,ov); }
  } else if(S.roofType==='monoslope'){
    if(S.showPressureMap){ buildKlMonoRoof(w,d,h,rh,ov,pRad); }
    else { buildMono(w,d,h,rh,ov,pRad); }
  }
  // Add parapet for ALL roof types
  if(S.parapet>0) addParapet(w, d, h);
}

function buildGable(w,d,h,rh,ov,pRad){
  const F=S.R.faces;
  const fm=getWindFaceMap();
  // Ridge endpoints along X-axis at peak height
  const r1=new THREE.Vector3(-w/2-ov,h+rh,0), r2=new THREE.Vector3(w/2+ov,h+rh,0);
  // Windward eave (front, +z)
  const ww1=new THREE.Vector3(-w/2-ov,h,d/2+ov), ww2=new THREE.Vector3(w/2+ov,h,d/2+ov);
  // Leeward eave (back, -z)
  const lw1=new THREE.Vector3(-w/2-ov,h,-d/2-ov), lw2=new THREE.Vector3(w/2+ov,h,-d/2-ov);

  // Roof panels — along-ridge: eave→ridge strips; crosswind: u from windward wall plane
  const roofExtent = d/2+ov;
  if(F.roof_ww?.zones?.length){
    if(F.roof_lw?.zones?.length)
      addGableRoofCrosswindStripsFromWallU(d,h,rh,ov,w,F);
    else {
      addRoofUpwindCpZones(d/2+ov, roofExtent, w+ov*2, h, rh, F);
      triFace([lw1,lw2,r2,r1],'roof_lw',F);
    }
  } else {
    let kWW = 'roof_ww', kLW = 'roof_lw';
    if(fm.back === 'windward'){ kWW = 'roof_lw'; kLW = 'roof_ww'; }
    else if(fm.left === 'windward' || fm.right === 'windward'){ kWW = 'roof_ww'; kLW = 'roof_ww'; }
    triFace([ww2,ww1,r1,r2], kWW, F);
    triFace([lw1,lw2,r2,r1], kLW, F);
  }

  // Gable end walls at SIDES (x = ±w/2+ov) — use correct face per wind direction: left=gable uses fm.left, right uses fm.right
  const leftFace=F[fm.left]||{p:0}, rightFace=F[fm.right]||{p:0};
  const leftGMat=new THREE.MeshStandardMaterial({
    color:S.showHeatmap?heatCol(leftFace.p):0xcc8844,
    roughness:.6,side:THREE.DoubleSide,
    transparent:S.viewMode==='transparent'||S.showHeatmap,
    opacity:S.viewMode==='transparent'?.4:1,
    wireframe:S.viewMode==='wireframe'
  });
  // Left gable end (x=-w/2) — uses fm.left face pressure
  const lgG=new THREE.BufferGeometry();
  lgG.setAttribute('position',new THREE.Float32BufferAttribute(new Float32Array([
    -w/2-ov,h,-d/2-ov, -w/2-ov,h,d/2+ov, -w/2-ov,h+rh,0]),3));
  lgG.computeVertexNormals();
  const lgM=new THREE.Mesh(lgG,leftGMat);lgM.castShadow=true;grpBuild.add(lgM);
  faceMap.set(lgM.uuid,{key:fm.left,...leftFace});
  // Right gable end (x=+w/2) — uses fm.right face pressure
  const rgG=new THREE.BufferGeometry();
  rgG.setAttribute('position',new THREE.Float32BufferAttribute(new Float32Array([
    w/2+ov,h,d/2+ov, w/2+ov,h,-d/2-ov, w/2+ov,h+rh,0]),3));
  rgG.computeVertexNormals();
  const rgM=new THREE.Mesh(rgG,new THREE.MeshStandardMaterial({
    color:S.showHeatmap?heatCol(rightFace.p):0xcc8844,
    roughness:.6,side:THREE.DoubleSide,
    transparent:S.viewMode==='transparent'||S.showHeatmap,
    opacity:S.viewMode==='transparent'?.4:1,
    wireframe:S.viewMode==='wireframe'
  }));rgM.castShadow=true;grpBuild.add(rgM);
  faceMap.set(rgM.uuid,{key:fm.right,...rightFace});

  // Ridge line highlight
  const rg=new THREE.BufferGeometry().setFromPoints([r1,r2]);
  grpBuild.add(new THREE.Line(rg,new THREE.LineBasicMaterial({color:0xffaa00})));
  // Edge outlines
  if(S.viewMode!=='wireframe'){
    [[ww1,ww2],[ww2,r2],[r1,ww1],[lw1,lw2],[lw2,r2],[lw1,r1]].forEach(([a,b])=>{
      const eg=new THREE.BufferGeometry().setFromPoints([a,b]);
      grpBuild.add(new THREE.Line(eg,new THREE.LineBasicMaterial({color:0,transparent:true,opacity:.2})));
    });
  }
}

function buildHip(w,d,h,rh,ov){
  const F=S.R.faces;
  const rl=Math.max(w-d,2);
  const a1=new THREE.Vector3(-rl/2,h+rh,0),a2=new THREE.Vector3(rl/2,h+rh,0);
  const c=[
    new THREE.Vector3(-w/2-ov,h,-d/2-ov),new THREE.Vector3(w/2+ov,h,-d/2-ov),
    new THREE.Vector3(w/2+ov,h,d/2+ov),new THREE.Vector3(-w/2-ov,h,d/2+ov)];
  // Main slopes are trapezoids (ridge shorter than eave). Do NOT use addGableRoofCrosswindStripsFromWallU /
  // addRoofUpwindCpZones here — those emit full-width rectangles (gable geometry). Mesh stays hip-shaped;
  // α < 10° hip-end U/D use Table 5.3(A) strip zones in calc() (hipEndUpwind/DownwindCpZonesTableA); heatmap uses face-average p.
  // Map +Z / −Z trapezoids to roof_ww vs roof_lw by wind (same rules as buildKlHipRoof) — not fixed building N/S.
  const fmHip = getWindFaceMap();
  let frontWWEdge = 'none', backWWEdge = 'none';
  if(fmHip.front === 'windward'){ frontWWEdge = 'eave'; }
  else if(fmHip.back === 'windward'){ backWWEdge = 'eave'; }
  else if(fmHip.right === 'windward'){ frontWWEdge = 'u0'; backWWEdge = 'u1'; }
  else if(fmHip.left === 'windward'){ frontWWEdge = 'u1'; backWWEdge = 'u0'; }
  const frontRoofKey = frontWWEdge !== 'none' ? 'roof_ww' : 'roof_lw';
  const backRoofKey  = backWWEdge !== 'none' ? 'roof_ww' : 'roof_lw';
  triFace([c[2],c[3],a1,a2], frontRoofKey, F);
  triFace([c[0],c[1],a2,a1], backRoofKey, F);
  // Hip-end roof facets — eave edge on building line (±w/2) so it meets ±X wall tops; overhang is on main slopes only
  triFace([new THREE.Vector3(-w/2,h,d/2), new THREE.Vector3(-w/2,h,-d/2), a1],'roof_hip_l',F);
  triFace([new THREE.Vector3(w/2,h,-d/2), new THREE.Vector3(w/2,h,d/2), a2],'roof_hip_r',F);
  const rg=new THREE.BufferGeometry().setFromPoints([a1,a2]);
  grpBuild.add(new THREE.Line(rg,new THREE.LineBasicMaterial({color:0xffaa00})));
}

function triFace(pts,fk,F){
  const v=[];
  if(pts.length===3){pts.forEach(p=>v.push(p.x,p.y,p.z))}
  else{v.push(pts[0].x,pts[0].y,pts[0].z,pts[1].x,pts[1].y,pts[1].z,pts[2].x,pts[2].y,pts[2].z,
    pts[0].x,pts[0].y,pts[0].z,pts[2].x,pts[2].y,pts[2].z,pts[3].x,pts[3].y,pts[3].z)}
  const g=new THREE.BufferGeometry();
  g.setAttribute('position',new THREE.Float32BufferAttribute(v,3));g.computeVertexNormals();
  const fd=F[fk]||{p:0};
  const mat=roofMat(fk);
  const m=new THREE.Mesh(g,mat);m.castShadow=true;
  faceMap.set(m.uuid,{key:fk,...fd});grpBuild.add(m);
}

// Table 5.3(A) — Subdivide upwind roof slope into Cpe zone strips (α < 10° only)
// wwZ: windward-edge z; roofExtent: horizontal distance windward→ridge; roofWidth: crosswind
function addRoofUpwindCpZones(wwZ,roofExtent,roofWidth,h,rh,F){
  const fd=F.roof_ww; if(!fd?.zones?.length) return false;
  const xL=-S.width/2-S.overhang, xR=S.width/2+S.overhang;
  let cum=0;
  fd.zones.forEach(z=>{
    const w=Math.min(z.width??z.area/roofWidth, Math.max(0, roofExtent-cum));
    if(w<0.01) return;
    const z1=wwZ-cum, z2=wwZ-cum-w;
    // z from ridge (0) to windward eave (≈roofExtent): low y at eave, high y at ridge
    const slope=(a)=>h+rh*(1-a/roofExtent);
    const y1=slope(z1), y2=slope(z2);
    const v=[xL,y1,z1, xR,y1,z1, xR,y2,z2, xL,y2,z2];
    const g=new THREE.BufferGeometry();
    g.setAttribute('position',new THREE.Float32BufferAttribute(v,3));
    g.setIndex([0,1,2, 0,2,3]);
    g.computeVertexNormals();
    const col=S.showHeatmap?heatCol(z.p):0xcc6622;
    const mat=new THREE.MeshStandardMaterial({color:col,roughness:.5,side:THREE.DoubleSide,
      transparent:S.viewMode==='transparent'||S.showHeatmap, opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe'});
    const m=new THREE.Mesh(g,mat); m.castShadow=true;
    faceMap.set(m.uuid,{key:'roof_ww',name:fd.name+' ('+z.dist+')',Cp_e:z.Cpe,Cp_i:fd.Cp_i,p:z.p,area:z.area,force:z.force,clause:fd.clause});
    grpBuild.add(m);
    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(g),new THREE.LineBasicMaterial({color:0,transparent:true,opacity:.25}));
      grpBuild.add(e);
    }
    cum+=w;
  });
  return true;
}

// Crosswind R (Fig 5.2): Δu from windward wall — strips along Z if wind on front/back, along X if wind on gable ends
function addGableRoofCrosswindStripsFromWallU(d,h,rh,ov,w,F){
  const fm = getWindFaceMap();
  const xL = -w/2 - ov, xR = w/2 + ov;
  const roofExtent = d/2 + ov;
  // +Z half: z ∈ [0, roofExtent], eave at z=roofExtent (y=h), ridge at z=0 (y=h+rh)
  const slopePosZ = zz=>h + rh * (1 - zz / roofExtent);
  // −Z half: z ∈ [−roofExtent, 0], eave at z=−roofExtent (y=h), ridge at z=0 (y=h+rh)
  const slopeNegZ = zz=>h + rh * (1 + zz / roofExtent);
  function emitRoofStrip(faceKey, zn, zHi, zLo, slopeFn){
    if(Math.abs(zHi - zLo) < 0.001) return;
    const y1 = slopeFn(zHi), y2 = slopeFn(zLo);
    const v = [xL,y1,zHi, xR,y1,zHi, xR,y2,zLo, xL,y2,zLo];
    const g = new THREE.BufferGeometry();
    g.setAttribute('position', new THREE.Float32BufferAttribute(v,3));
    g.setIndex([0,1,2, 0,2,3]);
    g.computeVertexNormals();
    const fd = F[faceKey];
    const col = S.showHeatmap ? heatCol(zn.p) : (faceKey==='roof_ww' ? 0xcc6622 : 0x996633);
    const mat = new THREE.MeshStandardMaterial({color:col, roughness:.5, side:THREE.DoubleSide,
      transparent:S.viewMode==='transparent'||S.showHeatmap, opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe'});
    const m = new THREE.Mesh(g,mat); m.castShadow = true;
    faceMap.set(m.uuid, {key:faceKey, name:fd.name+' ('+zn.dist+')', Cp_e:zn.Cpe, Cp_i:fd.Cp_i, p:zn.p, area:zn.area, force:zn.force, clause:fd.clause});
    grpBuild.add(m);
    if(S.viewMode!=='wireframe'){
      const e = new THREE.LineSegments(new THREE.EdgesGeometry(g), new THREE.LineBasicMaterial({color:0, transparent:true, opacity:.25}));
      grpBuild.add(e);
    }
  }
  function emitRoofStripX(faceKey, zn, xLo, xHi, zTop, zBot, slopeFn){
    if(Math.abs(xHi - xLo) < 0.001) return;
    const yT = slopeFn(zTop), yB = slopeFn(zBot);
    const v = [xLo,yT,zTop, xHi,yT,zTop, xHi,yB,zBot, xLo,yB,zBot];
    const g = new THREE.BufferGeometry();
    g.setAttribute('position', new THREE.Float32BufferAttribute(v,3));
    g.setIndex([0,1,2, 0,2,3]);
    g.computeVertexNormals();
    const fd = F[faceKey];
    const col = S.showHeatmap ? heatCol(zn.p) : (faceKey==='roof_ww' ? 0xcc6622 : 0x996633);
    const mat = new THREE.MeshStandardMaterial({color:col, roughness:.5, side:THREE.DoubleSide,
      transparent:S.viewMode==='transparent'||S.showHeatmap, opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe'});
    const m = new THREE.Mesh(g,mat); m.castShadow = true;
    faceMap.set(m.uuid, {key:faceKey, name:fd.name+' ('+zn.dist+')', Cp_e:zn.Cpe, Cp_i:fd.Cp_i, p:zn.p, area:zn.area, force:zn.force, clause:fd.clause});
    grpBuild.add(m);
    if(S.viewMode!=='wireframe'){
      const e = new THREE.LineSegments(new THREE.EdgesGeometry(g), new THREE.LineBasicMaterial({color:0, transparent:true, opacity:.25}));
      grpBuild.add(e);
    }
  }
  const zE = d/2 + ov;
  if(fm.front === 'windward'){
    let cum = 0, uW = -ov;
    F.roof_ww.zones.forEach(zn=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = uW + cum, uB = uA + du;
      emitRoofStrip('roof_ww', zn, d/2 - uA, d/2 - uB, slopePosZ);
      cum += du;
    });
    if(F.roof_lw?.zones?.length){
      cum = 0; uW = d/2;
      F.roof_lw.zones.forEach(zn=>{
        const du = zn.width;
        if(du < 0.001) return;
        const uA = uW + cum, uB = uA + du;
        emitRoofStrip('roof_lw', zn, d/2 - uA, d/2 - uB, slopeNegZ);
        cum += du;
      });
    }
  } else if(fm.back === 'windward'){
    let cum = 0, uW = -ov;
    F.roof_ww.zones.forEach(zn=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = uW + cum, uB = uA + du;
      emitRoofStrip('roof_ww', zn, uB - d/2, uA - d/2, slopeNegZ);
      cum += du;
    });
    if(F.roof_lw?.zones?.length){
      cum = 0; uW = d/2;
      F.roof_lw.zones.forEach(zn=>{
        const du = zn.width;
        if(du < 0.001) return;
        const uA = uW + cum, uB = uA + du;
        emitRoofStrip('roof_lw', zn, uB - d/2, uA - d/2, slopePosZ);
        cum += du;
      });
    }
  } else if(fm.right === 'windward'){
    let cum = 0, uW = -ov;
    F.roof_ww.zones.forEach((zn, i)=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = uW + cum, uB = uA + du;
      const xHi = w/2 - uA, xLo = w/2 - uB;
      emitRoofStripX('roof_ww', zn, xLo, xHi, zE, 0, slopePosZ);
      const znL = F.roof_lw?.zones?.[i];
      if(znL) emitRoofStripX('roof_lw', znL, xLo, xHi, 0, -zE, slopeNegZ);
      cum += du;
    });
  } else if(fm.left === 'windward'){
    let cum = 0, uW = -ov;
    F.roof_ww.zones.forEach((zn, i)=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = uW + cum, uB = uA + du;
      const xLo = uA - w/2, xHi = uB - w/2;
      emitRoofStripX('roof_ww', zn, xLo, xHi, zE, 0, slopePosZ);
      const znL = F.roof_lw?.zones?.[i];
      if(znL) emitRoofStripX('roof_lw', znL, xLo, xHi, 0, -zE, slopeNegZ);
      cum += du;
    });
  }
}

// Mono / flat: R crosswind strips — u along Z if wind on front/back, along X if wind on gable ends (Fig 5.2)
function addMonoOrFlatCrosswindStripsWallU(d,h,monoRise,ov,w,F){
  const fm = getWindFaceMap();
  const fd = F.roof_ww;
  if(!fd?.zones?.length) return false;
  const xL = -w/2 - ov, xR = w/2 + ov;
  const span = d + 2*ov;
  const rise = monoRise || 0;
  const yAtZ = zz=>h + rise * (d/2 + ov - zz) / span;
  const zR = d/2 + ov, zL = -d/2 - ov;
  function emitStripZ(zn, zHi, zLo){
    if(Math.abs(zHi - zLo) < 0.001) return;
    const y1 = yAtZ(zHi), y2 = yAtZ(zLo);
    const v = [xL,y1,zHi, xR,y1,zHi, xR,y2,zLo, xL,y2,zLo];
    const g = new THREE.BufferGeometry();
    g.setAttribute('position', new THREE.Float32BufferAttribute(v,3));
    g.setIndex([0,1,2, 0,2,3]);
    g.computeVertexNormals();
    const col = S.showHeatmap ? heatCol(zn.p) : 0xcc6622;
    const mat = new THREE.MeshStandardMaterial({color:col, roughness:.5, side:THREE.DoubleSide,
      transparent:S.viewMode==='transparent'||S.showHeatmap, opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe'});
    const m = new THREE.Mesh(g,mat); m.castShadow = true;
    faceMap.set(m.uuid, {key:'roof_ww', name:fd.name+' ('+zn.dist+')', Cp_e:zn.Cpe, Cp_i:fd.Cp_i, p:zn.p, area:zn.area, force:zn.force, clause:fd.clause});
    grpBuild.add(m);
    if(S.viewMode!=='wireframe'){
      const e = new THREE.LineSegments(new THREE.EdgesGeometry(g), new THREE.LineBasicMaterial({color:0, transparent:true, opacity:.25}));
      grpBuild.add(e);
    }
  }
  function emitStripX(zn, xLo, xHi){
    if(Math.abs(xHi - xLo) < 0.001) return;
    const yT = yAtZ(zR), yB = yAtZ(zL);
    const v = [xLo,yT,zR, xHi,yT,zR, xHi,yB,zL, xLo,yB,zL];
    const g = new THREE.BufferGeometry();
    g.setAttribute('position', new THREE.Float32BufferAttribute(v,3));
    g.setIndex([0,1,2, 0,2,3]);
    g.computeVertexNormals();
    const col = S.showHeatmap ? heatCol(zn.p) : 0xcc6622;
    const mat = new THREE.MeshStandardMaterial({color:col, roughness:.5, side:THREE.DoubleSide,
      transparent:S.viewMode==='transparent'||S.showHeatmap, opacity:(S.viewMode==='transparent'||S.showHeatmap)?.75:1,
      wireframe:S.viewMode==='wireframe'});
    const m = new THREE.Mesh(g,mat); m.castShadow = true;
    faceMap.set(m.uuid, {key:'roof_ww', name:fd.name+' ('+zn.dist+')', Cp_e:zn.Cpe, Cp_i:fd.Cp_i, p:zn.p, area:zn.area, force:zn.force, clause:fd.clause});
    grpBuild.add(m);
    if(S.viewMode!=='wireframe'){
      const e = new THREE.LineSegments(new THREE.EdgesGeometry(g), new THREE.LineBasicMaterial({color:0, transparent:true, opacity:.25}));
      grpBuild.add(e);
    }
  }
  let cum = 0, u0 = -ov;
  if(fm.front === 'windward'){
    fd.zones.forEach(zn=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = u0 + cum, uB = uA + du;
      emitStripZ(zn, d/2 - uA, d/2 - uB);
      cum += du;
    });
    return true;
  }
  if(fm.back === 'windward'){
    fd.zones.forEach(zn=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = u0 + cum, uB = uA + du;
      emitStripZ(zn, uB - d/2, uA - d/2);
      cum += du;
    });
    return true;
  }
  if(fm.right === 'windward'){
    fd.zones.forEach(zn=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = u0 + cum, uB = uA + du;
      emitStripX(zn, w/2 - uB, w/2 - uA);
      cum += du;
    });
    return true;
  }
  if(fm.left === 'windward'){
    fd.zones.forEach(zn=>{
      const du = zn.width;
      if(du < 0.001) return;
      const uA = u0 + cum, uB = uA + du;
      emitStripX(zn, uA - w/2, uB - w/2);
      cum += du;
    });
    return true;
  }
  return addRoofUpwindCpZones(d/2+ov, d+ov*2, w+ov*2, h, monoRise, F);
}

function buildMono(w,d,h,rh,ov,pRad){
  const F=S.R.faces;
  const monoRise=Math.tan(pRad)*d; // full rise across building depth
  // Low side (front, +z) at height h
  const fl=new THREE.Vector3(-w/2-ov,h,d/2+ov), fr=new THREE.Vector3(w/2+ov,h,d/2+ov);
  // High side (back, -z) at height h+monoRise
  const bl=new THREE.Vector3(-w/2-ov,h+monoRise,-d/2-ov), br=new THREE.Vector3(w/2+ov,h+monoRise,-d/2-ov);

  if(F.roof_ww?.zones?.length){
    addMonoOrFlatCrosswindStripsWallU(d,h,monoRise,ov,w,F);
  } else {
    triFace([fl,fr,br,bl],'roof_ww',F);
  }

  // Sidewalls at ±X are trapezoids in buildWalls / buildKlWalls (roof line along top).
  // Back (−z) vertical face is a single plane h+monoRise tall in buildWalls (no separate strip mesh).

  // Edge outlines
  if(S.viewMode!=='wireframe'){
    [[fl,fr],[fr,br],[br,bl],[bl,fl]].forEach(([a,b])=>{
      const eg=new THREE.BufferGeometry().setFromPoints([a,b]);
      grpBuild.add(new THREE.Line(eg,new THREE.LineBasicMaterial({color:0,transparent:true,opacity:.2})));
    });
  }
}

function addParapet(w,d,h){
  const mat=new THREE.MeshStandardMaterial({color:0x777777,roughness:.7});
  [{g:[w,S.parapet,.15],p:[0,h+S.parapet/2,d/2]},{g:[w,S.parapet,.15],p:[0,h+S.parapet/2,-d/2]},
   {g:[.15,S.parapet,d],p:[-w/2,h+S.parapet/2,0]},{g:[.15,S.parapet,d],p:[w/2,h+S.parapet/2,0]}
  ].forEach(({g,p})=>{const m=new THREE.Mesh(new THREE.BoxGeometry(...g),mat);m.position.set(...p);m.castShadow=true;grpBuild.add(m)});
}

// ── Local Pressure Zones: flat roof subdivided into corner/edge/interior zones ──
// Per AS/NZS 1170.2 Table 5.6 & Figure 5.3
function buildKlFlatRoof(w,d,h,ov){
  const F = S.R.faces;
  const faceData = F['roof_ww'];
  if(!faceData) return;

  // Clause 5.4.4 — dimension a for roofs
  const a = (h/w >= 0.2 || h/d >= 0.2) ? 0.2 * Math.min(w, d) : 2 * h;
  const isCladding = S.showPressureMap;
  const isGlazing  = false;
  let edgeDist, edgeKl, cornerKl, interiorKl;
  if(isGlazing){
    edgeDist = a; cornerKl = 3.0; edgeKl = 2.0; interiorKl = 1.5;
  } else if(isCladding){
    edgeDist = a; cornerKl = 3.0; edgeKl = 1.5; interiorKl = 1.0;
  } else {
    edgeDist = 0; cornerKl = 1.0; edgeKl = 1.0; interiorKl = 1.0;
  }

  const totalW = w + ov*2, totalD = d + ov*2;
  const eW = Math.min(edgeDist, totalW/2);
  const eD = Math.min(edgeDist, totalD/2);
  const halfW = Math.min(edgeDist/2, totalW/2);  // 0.5a boundary
  const halfD = Math.min(edgeDist/2, totalD/2);
  const centerW = totalW - 2*eW;
  const centerD = totalD - 2*eD;
  const ra2Kl = 2.0;  // RA2 Kl per Table 5.6

  const addRoofPatch = (pW, pD, cx, cz, Kl, zoneName) => {
    if(pW < 0.01 || pD < 0.01) return;
    const klP = localPnet(faceData, Kl, true);
    const col = klZoneColor(Kl, zoneName);
    const isTr = S.viewMode==='transparent', isWf = S.viewMode==='wireframe';
    const mat = new THREE.MeshStandardMaterial({
      color: col, roughness:.5, side: THREE.DoubleSide,
      transparent: isTr, opacity: isTr?.75:1, wireframe: isWf,
      polygonOffset: true, polygonOffsetFactor: 1, polygonOffsetUnits: 1
    });
    const geo = new THREE.PlaneGeometry(pW, pD);
    const mesh = new THREE.Mesh(geo, mat);
    mesh.position.set(cx, h + 0.01, cz);
    mesh.rotation.x = -Math.PI/2;
    mesh.castShadow=true; mesh.receiveShadow=true;
    mesh.renderOrder = 1;
    faceMap.set(mesh.uuid, {key:'roof_ww', ...faceData, Kl:Kl, klZone:zoneName, pMod:klP});
    grpBuild.add(mesh);
    if(S.viewMode!=='wireframe'){
      const e=new THREE.LineSegments(new THREE.EdgesGeometry(geo),new THREE.LineBasicMaterial({color:0x000000,transparent:true,opacity:.25,depthWrite:false}));
      e.position.copy(mesh.position); e.rotation.copy(mesh.rotation);
      e.renderOrder = 2;
      grpBuild.add(e);
    }
    grpBuild.add(spriteText(zoneName+' '+klP.toFixed(2)+' kPa', new THREE.Vector3(cx, h+0.3, cz), 0xffffff, 0.3));
  };

  if(edgeDist <= 0){
    addRoofPatch(totalW, totalD, 0, 0, 1.0, 'MR');
    return;
  }

  // RC1 = upwind corners — corners adjacent to the windward wall
  const fm = getWindFaceMap();
  const corners = [
    {cx: -totalW/2+eW/2, cz:  totalD/2-eD/2, adj:['front','left']},
    {cx:  totalW/2-eW/2, cz:  totalD/2-eD/2, adj:['front','right']},
    {cx: -totalW/2+eW/2, cz: -totalD/2+eD/2, adj:['back','left']},
    {cx:  totalW/2-eW/2, cz: -totalD/2+eD/2, adj:['back','right']},
  ];
  corners.forEach(c => {
    const isWW = c.adj.some(f => fm[f] === 'windward');
    addRoofPatch(eW, eD, c.cx, c.cz, isWW ? cornerKl : ra2Kl, isWW ? 'RC1' : 'RC2');
  });

  // 4 edge strips (between corners) — each split into RA2 (outer 0.5a) and RA1 (inner 0.5a to a)
  const innerW = eW - halfW;  // width of RA1 band (0.5a)
  const innerD = eD - halfD;

  if(centerW > 0.01){
    // Front edge (+z)
    addRoofPatch(centerW, halfD, 0,  totalD/2-halfD/2, ra2Kl, 'RA2');
    if(innerD > 0.01) addRoofPatch(centerW, innerD, 0,  totalD/2-halfD-innerD/2, edgeKl, 'RA1');
    // Back edge (-z)
    addRoofPatch(centerW, halfD, 0, -totalD/2+halfD/2, ra2Kl, 'RA2');
    if(innerD > 0.01) addRoofPatch(centerW, innerD, 0, -totalD/2+halfD+innerD/2, edgeKl, 'RA1');
  }
  if(centerD > 0.01){
    // Left edge
    addRoofPatch(halfW, centerD, -totalW/2+halfW/2, 0, ra2Kl, 'RA2');
    if(innerW > 0.01) addRoofPatch(innerW, centerD, -totalW/2+halfW+innerW/2, 0, edgeKl, 'RA1');
    // Right edge
    addRoofPatch(halfW, centerD,  totalW/2-halfW/2, 0, ra2Kl, 'RA2');
    if(innerW > 0.01) addRoofPatch(innerW, centerD,  totalW/2-halfW-innerW/2, 0, edgeKl, 'RA1');
  }

  // Center patch (interior / MR)
  if(centerW > 0.01 && centerD > 0.01){
    addRoofPatch(centerW, centerD, 0, 0, interiorKl, 'MR');
  }
}

// ── Local Pressure Zones: gable roof — subdivide each slope into edge/ridge/interior zones ──
function buildKlGableRoof(w,d,h,rh,ov,pRad){
  const F = S.R.faces;
  const a = (h/w >= 0.2 || h/d >= 0.2) ? 0.2 * Math.min(w, d) : 2 * h;
  const isCladding = S.showPressureMap;
  const isGlazing  = false;

  let cornerKl, edgeKl, ridgeKl, interiorKl;
  if(isGlazing){
    cornerKl = 3.0; edgeKl = 2.0; ridgeKl = 2.0; interiorKl = 1.5;
  } else if(isCladding){
    cornerKl = 3.0; edgeKl = 1.5; ridgeKl = 1.5; interiorKl = 1.0;
  } else {
    cornerKl = 1.0; edgeKl = 1.0; ridgeKl = 1.0; interiorKl = 1.0;
  }

  // RC1 applies to all roof corners (Kl=3.0)
  const pitchDeg = S.pitch || 0;
  const useRC2 = pitchDeg >= 10;

  // Ridge endpoints
  const r1=new THREE.Vector3(-w/2-ov,h+rh,0), r2=new THREE.Vector3(w/2+ov,h+rh,0);
  const ww1=new THREE.Vector3(-w/2-ov,h,d/2+ov), ww2=new THREE.Vector3(w/2+ov,h,d/2+ov);
  const lw1=new THREE.Vector3(-w/2-ov,h,-d/2-ov), lw2=new THREE.Vector3(w/2+ov,h,-d/2-ov);

  // Slope length (eave to ridge along slope)
  const slopeLen = (d/2+ov) / Math.cos(pRad);
  const totalW = w + ov*2;

  // Determine which edge of each slope is windward
  // Front slope: pts=[ww2(+x), ww1(-x), r1, r2] → u=0 is right(+x), u=1 is left(-x)
  // Back slope:  pts=[lw1(-x), lw2(+x), r2, r1] → u=0 is left(-x), u=1 is right(+x)
  const fm = getWindFaceMap();
  let frontWWEdge = 'none', backWWEdge = 'none';
  if(fm.front === 'windward'){ frontWWEdge = 'eave'; }
  else if(fm.back === 'windward'){ backWWEdge = 'eave'; }
  else if(fm.right === 'windward'){ frontWWEdge = 'u0'; backWWEdge = 'u1'; }
  else if(fm.left === 'windward'){ frontWWEdge = 'u1'; backWWEdge = 'u0'; }
  const frontRoofKey = frontWWEdge !== 'none' ? 'roof_ww' : 'roof_lw';
  const backRoofKey  = backWWEdge !== 'none' ? 'roof_ww' : 'roof_lw';
  buildKlRoofSlope([ww2,ww1,r1,r2], frontRoofKey, F, totalW, slopeLen, a,
    cornerKl, edgeKl, ridgeKl, interiorKl, h, rh, d/2+ov, frontWWEdge, useRC2);
  buildKlRoofSlope([lw1,lw2,r2,r1], backRoofKey, F, totalW, slopeLen, a,
    cornerKl, edgeKl, ridgeKl, interiorKl, h, rh, d/2+ov, backWWEdge, useRC2);

  // Gable end walls — windward gable gets edge Kl, leeward gets MR
  const gableEdgeKl = isCladding ? 1.5 : (isGlazing ? 2.0 : 1.0);
  const leftIsWW = (fm.left === 'windward');
  const rightIsWW = (fm.right === 'windward');
  const leftKl = leftIsWW ? gableEdgeKl : interiorKl;
  const rightKl = rightIsWW ? gableEdgeKl : interiorKl;
  const gMatL=new THREE.MeshStandardMaterial({
    color: klZoneColor(leftKl, leftIsWW ? 'WA1' : 'MR'),
    roughness:.6,side:THREE.DoubleSide,
    transparent:S.viewMode==='transparent',
    opacity:S.viewMode==='transparent'?.4:1,
    wireframe:S.viewMode==='wireframe'
  });
  const gMatR=new THREE.MeshStandardMaterial({
    color: klZoneColor(rightKl, rightIsWW ? 'WA1' : 'MR'),
    roughness:.6,side:THREE.DoubleSide,
    transparent:S.viewMode==='transparent',
    opacity:S.viewMode==='transparent'?.4:1,
    wireframe:S.viewMode==='wireframe'
  });
  const lgG=new THREE.BufferGeometry();
  lgG.setAttribute('position',new THREE.Float32BufferAttribute(new Float32Array([
    -w/2-ov,h,-d/2-ov, -w/2-ov,h,d/2+ov, -w/2-ov,h+rh,0]),3));
  lgG.computeVertexNormals();
  grpBuild.add(new THREE.Mesh(lgG,gMatL));
  const rgG=new THREE.BufferGeometry();
  rgG.setAttribute('position',new THREE.Float32BufferAttribute(new Float32Array([
    w/2+ov,h,d/2+ov, w/2+ov,h,-d/2-ov, w/2+ov,h+rh,0]),3));
  rgG.computeVertexNormals();
  grpBuild.add(new THREE.Mesh(rgG,gMatR));

  // Ridge line
  const rg=new THREE.BufferGeometry().setFromPoints([r1,r2]);
  grpBuild.add(new THREE.Line(rg,new THREE.LineBasicMaterial({color:0xffaa00})));
  if(S.viewMode!=='wireframe'){
    [[ww1,ww2],[ww2,r2],[r1,ww1],[lw1,lw2],[lw2,r2],[lw1,r1]].forEach(([aa,bb])=>{
      const eg=new THREE.BufferGeometry().setFromPoints([aa,bb]);
      grpBuild.add(new THREE.Line(eg,new THREE.LineBasicMaterial({color:0,transparent:true,opacity:.2})));
    });
  }
}

// Build a single roof slope with Kl pressure zones
// useRC2: if true, ridge-side corners get cornerKl (RC2); eave corners get edgeKl
//         if false (pitch < 10°), no corners get RC designation
function buildKlRoofSlope(pts, faceKey, F, slopeW, slopeLen, a,
  cornerKl, edgeKl, ridgeKl, interiorKl, eaveH, rh, halfD, windwardEdge, useRC2){

  const faceData = F[faceKey];
  if(!faceData) return;

  const eA = Math.min(a, slopeLen/2);
  const sA = Math.min(a, slopeW/2);
  const halfA_v = Math.min(a/2, slopeLen/2);
  const halfA_u = Math.min(a/2, slopeW/2);

  // u = 0..1 along width; v = 0..1 along slope (eave=0 to ridge=1)
  const uEdge = sA / slopeW;           // a boundary
  const vEdge = eA / slopeLen;
  const uHalf = halfA_u / slopeW;      // 0.5a boundary
  const vHalf = halfA_v / slopeLen;
  const uInner = 1 - uEdge;
  const vInner = 1 - vEdge;
  const uInnerHalf = 1 - uHalf;
  const vInnerHalf = 1 - vHalf;
  const ra2Kl = 2.0;  // RA2 Kl per Table 5.6

  // Create zone patches by interpolating quad corners
  // pts order: [eaveR, eaveL, ridgeL, ridgeR] for windward slope

  const lerp3 = (a,b,t) => new THREE.Vector3(
    a.x+(b.x-a.x)*t, a.y+(b.y-a.y)*t, a.z+(b.z-a.z)*t);

  const quadPt = (u,v) => {
    const eave = lerp3(pts[0], pts[1], u);
    const ridge = lerp3(pts[3], pts[2], u);
    return lerp3(eave, ridge, v);
  };

  const addSlopePatch = (u1,v1,u2,v2,Kl,zoneName) => {
    if(Math.abs(u2-u1)<0.001 || Math.abs(v2-v1)<0.001) return;
    const c00=quadPt(u1,v1), c10=quadPt(u2,v1), c11=quadPt(u2,v2), c01=quadPt(u1,v2);
    const verts = new Float32Array([
      c00.x,c00.y,c00.z, c10.x,c10.y,c10.z, c11.x,c11.y,c11.z,
      c00.x,c00.y,c00.z, c11.x,c11.y,c11.z, c01.x,c01.y,c01.z
    ]);
    const geo = new THREE.BufferGeometry();
    geo.setAttribute('position', new THREE.Float32BufferAttribute(verts,3));
    geo.computeVertexNormals();

    const klP = localPnet(faceData, Kl, true);
    const col = klZoneColor(Kl, zoneName);
    const isTr = S.viewMode==='transparent', isWf = S.viewMode==='wireframe';
    const mat = new THREE.MeshStandardMaterial({
      color:col, roughness:.5, side:THREE.DoubleSide,
      transparent:isTr, opacity:isTr?.75:1, wireframe:isWf,
      polygonOffset: true, polygonOffsetFactor: 1, polygonOffsetUnits: 1
    });
    const mesh = new THREE.Mesh(geo, mat);
    mesh.castShadow=true; mesh.receiveShadow=true;
    mesh.renderOrder = 1;
    faceMap.set(mesh.uuid, {key:faceKey, ...faceData, Kl:Kl, klZone:zoneName, pMod:klP});
    grpBuild.add(mesh);

    const center = quadPt((u1+u2)/2, (v1+v2)/2);
    center.y += 0.25;
    const patchArea = (u2-u1)*slopeW * (v2-v1)*slopeLen;
    if(patchArea > 0.5){
      grpBuild.add(spriteText(zoneName+'\n'+klP.toFixed(2)+' kPa', center, 0xffffff, 0.35));
    }
  };

  // Zone layout based on which edge faces the wind
  // windwardEdge: 'eave' (v=0), 'u0' (u=0 side), 'u1' (u=1 side), 'none' (leeward)

  if(windwardEdge === 'eave'){
    if(useRC2){
      // Pitch >= 10°: no corner zones on upwind — RA2/RA1 full-width strips per Figure 5.3
      addSlopePatch(0, 0, 1, vHalf, ra2Kl, 'RA2');
      addSlopePatch(0, vHalf, 1, vEdge, edgeKl, 'RA1');
      addSlopePatch(0, vEdge, 1, 1, interiorKl, 'MR');
    } else {
      // Pitch < 10°: RC1 corner zones (< a from two edges) per Table 5.6
      addSlopePatch(0, 0, uEdge, vEdge, cornerKl, 'RC1');
      addSlopePatch(uInner, 0, 1, vEdge, cornerKl, 'RC1');
      addSlopePatch(uEdge, 0, uInner, vHalf, ra2Kl, 'RA2');
      addSlopePatch(uEdge, vHalf, uInner, vEdge, edgeKl, 'RA1');
      addSlopePatch(0, vEdge, 1, 1, interiorKl, 'MR');
    }

  } else if(windwardEdge === 'u0'){
    if(useRC2){
      addSlopePatch(0, 0, uHalf, 1, ra2Kl, 'RA2');
      addSlopePatch(uHalf, 0, uEdge, 1, edgeKl, 'RA1');
      addSlopePatch(uEdge, 0, 1, 1, interiorKl, 'MR');
    } else {
      addSlopePatch(0, 0, uEdge, vEdge, cornerKl, 'RC1');
      addSlopePatch(0, vInner, uEdge, 1, cornerKl, 'RC1');
      addSlopePatch(0, vEdge, uHalf, vInner, ra2Kl, 'RA2');
      addSlopePatch(uHalf, vEdge, uEdge, vInner, edgeKl, 'RA1');
      addSlopePatch(uEdge, 0, 1, 1, interiorKl, 'MR');
    }

  } else if(windwardEdge === 'u1'){
    if(useRC2){
      addSlopePatch(uInnerHalf, 0, 1, 1, ra2Kl, 'RA2');
      addSlopePatch(uInner, 0, uInnerHalf, 1, edgeKl, 'RA1');
      addSlopePatch(0, 0, uInner, 1, interiorKl, 'MR');
    } else {
      addSlopePatch(uInner, 0, 1, vEdge, cornerKl, 'RC1');
      addSlopePatch(uInner, vInner, 1, 1, cornerKl, 'RC1');
      addSlopePatch(uInnerHalf, vEdge, 1, vInner, ra2Kl, 'RA2');
      addSlopePatch(uInner, vEdge, uInnerHalf, vInner, edgeKl, 'RA1');
      addSlopePatch(0, 0, uInner, 1, interiorKl, 'MR');
    }

  } else {
    // ═══ LEEWARD SLOPE: RA3/RA4 along ridge (v=1) + RC2 corners if pitch ≥ 10° ═══
    if(useRC2){
      // RC2 corners at ridge ends (< a from both ridge and side edge)
      addSlopePatch(0, vInner, uEdge, 1, cornerKl, 'RC2');
      addSlopePatch(uInner, vInner, 1, 1, cornerKl, 'RC2');
      // RA4 strip (< 0.5a from ridge, between corners) — Kl=2.0
      addSlopePatch(uEdge, vInnerHalf, uInner, 1, ra2Kl, 'RA4');
      // RA3 strip (0.5a to a from ridge, between corners) — Kl=1.5
      addSlopePatch(uEdge, vInner, uInner, vInnerHalf, ridgeKl, 'RA3');
      // MR — everything below the ridge zone
      addSlopePatch(0, 0, 1, vInner, interiorKl, 'MR');
    } else {
      // pitch < 10°: no RA3/RA4/RC2, all MR
      addSlopePatch(0, 0, 1, 1, interiorKl, 'MR');
    }
  }
}

// ── Local Pressure Zones: hip roof — 2 quad slopes + 2 triangular hip faces ──
function buildKlHipRoof(w,d,h,rh,ov){
  const F = S.R.faces;
  const a = (h/w >= 0.2 || h/d >= 0.2) ? 0.2 * Math.min(w, d) : 2 * h;
  const isCladding = S.showPressureMap;
  const isGlazing  = false;

  let cornerKl, edgeKl, ridgeKl, interiorKl;
  if(isGlazing){
    cornerKl = 3.0; edgeKl = 2.0; ridgeKl = 2.0; interiorKl = 1.5;
  } else if(isCladding){
    cornerKl = 3.0; edgeKl = 1.5; ridgeKl = 1.5; interiorKl = 1.0;
  } else {
    cornerKl = 1.0; edgeKl = 1.0; ridgeKl = 1.0; interiorKl = 1.0;
  }

  // RC1 applies to all roof corners (Kl=3.0)
  const pitchDeg = S.pitch || 0;
  const useRC2 = pitchDeg >= 10;

  const rl = Math.max(w-d, 2);
  const a1 = new THREE.Vector3(-rl/2, h+rh, 0);
  const a2 = new THREE.Vector3(rl/2, h+rh, 0);
  const c = [
    new THREE.Vector3(-w/2-ov, h, -d/2-ov),
    new THREE.Vector3(w/2+ov, h, -d/2-ov),
    new THREE.Vector3(w/2+ov, h, d/2+ov),
    new THREE.Vector3(-w/2-ov, h, d/2+ov)
  ];

  // Quad slopes: determine windward/leeward from face map
  const slopeLen = (d/2+ov) / Math.cos(Math.atan2(rh, d/2+ov));
  const totalW = w + ov*2;
  const fm = getWindFaceMap();
  // Hip front: pts=[c[2](+x,+z), c[3](-x,+z), a1, a2] → u=0 right(+x), u=1 left(-x)
  // Hip back:  pts=[c[0](-x,-z), c[1](+x,-z), a2, a1] → u=0 left(-x), u=1 right(+x)
  let frontWWEdge = 'none', backWWEdge = 'none';
  if(fm.front === 'windward'){ frontWWEdge = 'eave'; }
  else if(fm.back === 'windward'){ backWWEdge = 'eave'; }
  else if(fm.right === 'windward'){ frontWWEdge = 'u0'; backWWEdge = 'u1'; }
  else if(fm.left === 'windward'){ frontWWEdge = 'u1'; backWWEdge = 'u0'; }
  const frontRoofKey = frontWWEdge !== 'none' ? 'roof_ww' : 'roof_lw';
  const backRoofKey  = backWWEdge !== 'none' ? 'roof_ww' : 'roof_lw';
  buildKlRoofSlope([c[2],c[3],a1,a2], frontRoofKey, F, totalW, slopeLen, a,
    cornerKl, edgeKl, ridgeKl, interiorKl, h, rh, d/2+ov, frontWWEdge, useRC2);
  buildKlRoofSlope([c[0],c[1],a2,a1], backRoofKey, F, totalW, slopeLen, a,
    cornerKl, edgeKl, ridgeKl, interiorKl, h, rh, d/2+ov, backWWEdge, useRC2);

  // Triangular hip faces (sides) — simpler: just a single triangle with Kl coloring
  const addHipTriangle = (pts3, faceKey) => {
    const fd = F[faceKey];
    if(!fd) return;
    const klP = localPnet(fd, edgeKl, true); // Hip face = edge zone, Kp applies
    const zoneName = klZoneName(faceKey, edgeKl);
    const col = klZoneColor(edgeKl, zoneName);
    const isTr = S.viewMode==='transparent', isWf = S.viewMode==='wireframe';
    const mat = new THREE.MeshStandardMaterial({
      color:col, roughness:.5, side:THREE.DoubleSide,
      transparent:isTr, opacity:isTr?.75:1, wireframe:isWf,
      polygonOffset: true, polygonOffsetFactor: 1, polygonOffsetUnits: 1
    });
    const verts = new Float32Array([
      pts3[0].x,pts3[0].y,pts3[0].z,
      pts3[1].x,pts3[1].y,pts3[1].z,
      pts3[2].x,pts3[2].y,pts3[2].z
    ]);
    const geo = new THREE.BufferGeometry();
    geo.setAttribute('position', new THREE.Float32BufferAttribute(verts,3));
    geo.computeVertexNormals();
    const mesh = new THREE.Mesh(geo, mat);
    mesh.castShadow = true; mesh.receiveShadow = true;
    mesh.renderOrder = 1;
    faceMap.set(mesh.uuid, {key:faceKey, ...fd, Kl:edgeKl, klZone:zoneName, pMod:klP});
    grpBuild.add(mesh);
    // Zone name + pressure label at centroid
    const cx = (pts3[0].x+pts3[1].x+pts3[2].x)/3;
    const cy = (pts3[0].y+pts3[1].y+pts3[2].y)/3 + 0.25;
    const cz = (pts3[0].z+pts3[1].z+pts3[2].z)/3;
    grpBuild.add(spriteText(zoneName+'\n'+klP.toFixed(2)+' kPa', new THREE.Vector3(cx,cy,cz), 0xffffff, 0.35));
  };

  addHipTriangle([new THREE.Vector3(-w/2,h,d/2), new THREE.Vector3(-w/2,h,-d/2), a1], 'roof_hip_l');
  addHipTriangle([new THREE.Vector3(w/2,h,-d/2), new THREE.Vector3(w/2,h,d/2), a2], 'roof_hip_r');

  // Ridge line
  const rg = new THREE.BufferGeometry().setFromPoints([a1,a2]);
  grpBuild.add(new THREE.Line(rg, new THREE.LineBasicMaterial({color:0xffaa00})));
}

// ── Local Pressure Zones: monoslope roof — single slope with edge/corner/interior zones ──
function buildKlMonoRoof(w,d,h,rh,ov,pRad){
  const F = S.R.faces;
  const a = (h/w >= 0.2 || h/d >= 0.2) ? 0.2 * Math.min(w, d) : 2 * h;
  const isCladding = S.showPressureMap;
  const isGlazing  = false;

  let cornerKl, edgeKl, ridgeKl, interiorKl;
  if(isGlazing){
    cornerKl = 3.0; edgeKl = 2.0; ridgeKl = 2.0; interiorKl = 1.5;
  } else if(isCladding){
    cornerKl = 3.0; edgeKl = 1.5; ridgeKl = 1.5; interiorKl = 1.0;
  } else {
    cornerKl = 1.0; edgeKl = 1.0; ridgeKl = 1.0; interiorKl = 1.0;
  }

  // RC1 applies to all roof corners (Kl=3.0)
  const pitchDeg = S.pitch || 0;
  const useRC2 = pitchDeg >= 10;

  const monoRise = Math.tan(pRad) * d;
  // Low side (front, +z) at height h
  const fl = new THREE.Vector3(-w/2-ov, h, d/2+ov);
  const fr = new THREE.Vector3(w/2+ov, h, d/2+ov);
  // High side (back, -z) at height h+monoRise
  const bl = new THREE.Vector3(-w/2-ov, h+monoRise, -d/2-ov);
  const br = new THREE.Vector3(w/2+ov, h+monoRise, -d/2-ov);

  const slopeLen = (d+ov*2) / Math.cos(pRad);
  const totalW = w + ov*2;

  // Determine windward edge for mono slope
  // Default pts [fl,fr,br,bl]: quadPt v=0 = low (+z) eave, v=1 = high (−z) — correct for fm.front windward.
  // fm.back windward: wind hits high wall first; Kl strips must start at high eave → swap to [bl,br,fr,fl] so v=0 is windward.
  // u0/u1: wind on L/R gable ends — keep default pts (u measures across width).
  const fm = getWindFaceMap();
  let monoWWEdge = 'eave';
  if(fm.left === 'windward') monoWWEdge = 'u0';
  else if(fm.right === 'windward') monoWWEdge = 'u1';
  const monoPts = fm.back === 'windward' ? [bl, br, fr, fl] : [fl, fr, br, bl];
  buildKlRoofSlope(monoPts, 'roof_ww', F, totalW, slopeLen, a,
    cornerKl, edgeKl, ridgeKl, interiorKl, h, monoRise, d/2+ov, monoWWEdge, useRC2);

  // Sidewall trapezoids: buildKlWalls; back vertical face: single plane in buildKlWalls

  // Edge outlines
  if(S.viewMode !== 'wireframe'){
    [[fl,fr],[fr,br],[br,bl],[bl,fl]].forEach(([aa,bb]) => {
      const eg = new THREE.BufferGeometry().setFromPoints([aa,bb]);
      grpBuild.add(new THREE.Line(eg, new THREE.LineBasicMaterial({color:0, transparent:true, opacity:.2})));
    });
  }
}

// ═══════════════════════════════════════════════
//   DIMENSION LINES
// ═══════════════════════════════════════════════
function buildDims(){
  const w=S.width,d=S.depth,h=S.height;
  const pRad=S.pitch*Math.PI/180;
  const rh=S.roofType!=='flat'?Math.tan(pRad)*d/2:0;
  const off=2.5;
  dimLine(new THREE.Vector3(-w/2,-.4,d/2+off),new THREE.Vector3(w/2,-.4,d/2+off),w.toFixed(1)+'m',0xffaa44);
  dimLine(new THREE.Vector3(w/2+off,-.4,d/2),new THREE.Vector3(w/2+off,-.4,-d/2),d.toFixed(1)+'m',0x44aaff);
  dimLine(new THREE.Vector3(w/2+off,0,d/2+off),new THREE.Vector3(w/2+off,h,d/2+off),h.toFixed(1)+'m',0x44ff88);
  if(S.pitch>0&&S.roofType!=='flat'){
    dimLine(new THREE.Vector3(-w/2-off,h,0),new THREE.Vector3(-w/2-off,h+rh,0),rh.toFixed(1)+'m',0xff4488);
    grpDim.add(spriteText(S.pitch+'°',new THREE.Vector3(-w/2-off-1.5,h+rh/2,d/4),0xffcc00,.45));
    dimLine(new THREE.Vector3(-w/2-off-2.5,0,d/2+off),new THREE.Vector3(-w/2-off-2.5,h+rh,d/2+off),(h+rh).toFixed(1)+'m total',0xff8844);
  }
}
function dimLine(a,b,label,col,targetGrp){
  const grp=targetGrp||grpDim;
  const g=new THREE.BufferGeometry().setFromPoints([a,b]);
  const l=new THREE.Line(g,new THREE.LineDashedMaterial({color:col,dashSize:.3,gapSize:.15}));
  l.computeLineDistances();grp.add(l);
  const dir=new THREE.Vector3().subVectors(b,a).normalize();
  let perp=new THREE.Vector3(-dir.z,0,dir.x).multiplyScalar(.3);
  if(perp.lengthSq()<.001)perp.set(.3,0,0);
  [a,b].forEach(p=>{
    const tg=new THREE.BufferGeometry().setFromPoints([p.clone().add(perp),p.clone().sub(perp)]);
    grp.add(new THREE.Line(tg,new THREE.LineBasicMaterial({color:col})));
  });
  const mid=new THREE.Vector3().addVectors(a,b).multiplyScalar(.5);mid.y+=.5;
  grp.add(spriteText(label,mid,col,.6));
}
function dimLineNoLabel(a,b,col,targetGrp){
  const grp=targetGrp||grpDim;
  const g=new THREE.BufferGeometry().setFromPoints([a,b]);
  const l=new THREE.Line(g,new THREE.LineDashedMaterial({color:col,dashSize:.3,gapSize:.15}));
  l.computeLineDistances();grp.add(l);
  const dir=new THREE.Vector3().subVectors(b,a).normalize();
  let perp=new THREE.Vector3(-dir.z,0,dir.x).multiplyScalar(.3);
  if(perp.lengthSq()<.001)perp.set(.3,0,0);
  [a,b].forEach(p=>{
    const tg=new THREE.BufferGeometry().setFromPoints([p.clone().add(perp),p.clone().sub(perp)]);
    grp.add(new THREE.Line(tg,new THREE.LineBasicMaterial({color:col})));
  });
}

// ═══════════════════════════════════════════════
//   LABELS
// ═══════════════════════════════════════════════
function buildLabels(){
  const w=S.width,d=S.depth,h=S.height,F=S.R.faces;
  const fm = getWindFaceMap();
  const rn = {windward:'Windward', leeward:'Leeward', sidewall1:'Side Wall L', sidewall2:'Side Wall R'};
  const rc = {windward:0x66bbff, leeward:0x66ff88, sidewall1:0xffaa44, sidewall2:0xffaa44};
  const rs = {windward:.55, leeward:.55, sidewall1:.45, sidewall2:.45};

  // Place labels at physical face centres with wind-role text
  grpLabel.add(spriteText(rn[fm.front], new THREE.Vector3(0,h/2,d/2+1.5), rc[fm.front], rs[fm.front]));
  grpLabel.add(spriteText(rn[fm.back],  new THREE.Vector3(0,h/2,-d/2-1.5), rc[fm.back], rs[fm.back]));
  grpLabel.add(spriteText(rn[fm.left],  new THREE.Vector3(-w/2-2,h/2,0), rc[fm.left], rs[fm.left]));
  grpLabel.add(spriteText(rn[fm.right], new THREE.Vector3(w/2+2,h/2,0), rc[fm.right], rs[fm.right]));

  if(S.showHeatmap){
    // Reverse map: wind role → physical position for heatmap pressure values
    const rp = {};
    rp[fm.front] = new THREE.Vector3(0,h*.25,d/2+1.5);
    rp[fm.back]  = new THREE.Vector3(0,h*.25,-d/2-1.5);
    rp[fm.left]  = new THREE.Vector3(-w/2-1.5,h*.25,0);
    rp[fm.right] = new THREE.Vector3(w/2+1.5,h*.25,0);
    for(let k in F){
      let p;
      if(rp[k]) p = rp[k];
      else if(k==='roof_ww') p = new THREE.Vector3(w/4,h+1.5,d/4);
      else if(k==='roof_lw') p = new THREE.Vector3(w/4,h+1.5,-d/4);
      else if(k==='roof_hip_l') p = new THREE.Vector3(-w/2.6,h+1.55,0);
      else if(k==='roof_hip_r') p = new THREE.Vector3(w/2.6,h+1.55,0);
      if(p) grpLabel.add(spriteText(F[k].p.toFixed(2)+' kPa',p,0xffffff,.45));
    }
  }
}

// ═══════════════════════════════════════════════
//   PRESSURE ARROWS & INTERNAL PRESSURE
// ═══════════════════════════════════════════════
function buildPArrows(){
  const w=S.width,d=S.depth,h=S.height,F=S.R.faces;
  const fm = getWindFaceMap();
  let mx = 0.1;
  for(const f of Object.values(F)){
    if(!f) continue;
    const ap = Math.abs(Number(f.p));
    if(Number.isFinite(ap) && ap > mx) mx = ap;
  }
  function ar(o,dir,p,ml){
    const len=Math.min(2.8, Math.max(0.25, Math.abs(Number(p)||0)/mx*ml));
    const c=p>0?0xff4444:0x4444ff;
    const dd=dir.clone().normalize();if(p<0)dd.negate();
    grpArrows.add(new THREE.ArrowHelper(dd,o,len,c));
  }
  // Front/back faces — pressure from wind role
  for(let r=0;r<3;r++)for(let c=0;c<3;c++){
    const fx=(c+.5)/3-.5,fy=(r+.5)/3;
    ar(new THREE.Vector3(fx*w,fy*h,d/2),new THREE.Vector3(0,0,-1),F[fm.front].p,3);
    ar(new THREE.Vector3(fx*w,fy*h,-d/2),new THREE.Vector3(0,0,1),F[fm.back].p,3);
  }
  // Left/right faces — pressure from wind role
  for(let r=0;r<3;r++)for(let c=0;c<3;c++){
    const fz=(c+.5)/3-.5,fy=(r+.5)/3;
    ar(new THREE.Vector3(-w/2,fy*h,fz*d),new THREE.Vector3(1,0,0),F[fm.left].p,3);
    ar(new THREE.Vector3(w/2,fy*h,fz*d),new THREE.Vector3(-1,0,0),F[fm.right].p,3);
  }
  // Roof — optional arrows (hidden with heatmap: zone colors + labels are enough; avoids cone clutter)
  if(S.showHeatmap) return;
  const ov = S.overhang || 0;
  const pRad = (S.pitch || 0) * Math.PI / 180;
  const rhGable = S.roofType !== 'flat' && S.roofType !== 'monoslope' ? Math.tan(pRad) * d / 2 : 0;
  const hz = d / 2 + ov;
  const bump = 0.35;
  if(S.roofType === 'flat'){
    const pr = F.roof_ww?.p ?? 0;
    for(let r=0;r<2;r++)for(let c=0;c<3;c++){
      const fx=(c+.5)/3-.5, fz=(r+.5)/2-.5;
      const o=new THREE.Vector3(fx*w, h+bump, fz*d);
      ar(o, new THREE.Vector3(0,1,0), pr, 3);
    }
  } else if(S.roofType === 'monoslope'){
    const monoRise = Math.tan(pRad) * d;
    const n = new THREE.Vector3(0, d + 2 * ov, monoRise).normalize();
    const o = new THREE.Vector3(0, h + monoRise / 2, 0);
    o.addScaledVector(n, bump);
    ar(o, n, F.roof_ww?.p ?? 0, 3);
  } else {
    const nFront = new THREE.Vector3(0, hz, rhGable).normalize();
    const nBack  = new THREE.Vector3(0, hz, -rhGable).normalize();
    const frontIsWW = fm.back !== 'windward';
    const pF = frontIsWW ? F.roof_ww.p : F.roof_lw.p;
    const pB = frontIsWW ? F.roof_lw.p : F.roof_ww.p;
    const zc = (d / 2 + ov) / 2;
    const oF = new THREE.Vector3(0, h + rhGable / 2, zc);
    const oB = new THREE.Vector3(0, h + rhGable / 2, -zc);
    oF.addScaledVector(nFront, bump);
    oB.addScaledVector(nBack, bump);
    for(let i=-1;i<=1;i+=2){
      const off = i * w / 6;
      ar(oF.clone().setX(off), nFront, pF, 3);
      ar(oB.clone().setX(off), nBack, pB, 3);
    }
  }
}

function buildInternal(){
  const w=S.width,d=S.depth,h=S.height;
  const pi=S.R.qz*S.R.Cpi, col=pi>0?0xff8844:0x44aaff;
  [{d:new THREE.Vector3(0,0,1),p:new THREE.Vector3(0,h/2,d/4)},
   {d:new THREE.Vector3(0,0,-1),p:new THREE.Vector3(0,h/2,-d/4)},
   {d:new THREE.Vector3(-1,0,0),p:new THREE.Vector3(-w/4,h/2,0)},
   {d:new THREE.Vector3(1,0,0),p:new THREE.Vector3(w/4,h/2,0)},
   {d:new THREE.Vector3(0,1,0),p:new THREE.Vector3(0,h*.7,0)}
  ].forEach(({d:dir,p:pos})=>{
    const len=Math.abs(pi)*2+.5;
    const dd=pi<0?dir.clone().negate():dir;
    grpInternal.add(new THREE.ArrowHelper(dd,pos,len,col,len*.2,len*.12));
  });
  grpLabel.add(spriteText('Cp,i = '+S.R.Cpi.toFixed(2),new THREE.Vector3(0,h*.35,0),0xffcc00,.45));
}

// ═══════════════════════════════════════════════
//   WIND INDICATOR & PARTICLES
// ═══════════════════════════════════════════════
function buildWindIndicator(){
  const d=S.depth,h=S.height;
  // Same relative bearing as getWindFaceMap / calc() hitsSide — otherwise arrows disagree with roof U vs R (Fig 5.2)
  const relW = ((S.windAngle - S.mapBuildingAngle) % 360 + 360) % 360;
  const ang=(relW+(S.R.angleOff||0))*Math.PI/180;
  const dir=new THREE.Vector3(-Math.sin(ang),0,-Math.cos(ang));
  const org=new THREE.Vector3(Math.sin(ang)*(d/2+10),h/2,Math.cos(ang)*(d/2+10));
  grpWind.add(new THREE.ArrowHelper(dir,org,6,0xff3333,1.5,.8));
  grpWind.add(spriteText('WIND',org.clone().add(new THREE.Vector3(0,2.5,0)),0xff4444,.55));
  const perp=new THREE.Vector3(Math.cos(ang),0,-Math.sin(ang));
  for(let i=-2;i<=2;i++){
    if(!i)continue;
    grpWind.add(new THREE.ArrowHelper(dir,org.clone().add(perp.clone().multiplyScalar(i*1.5)),4,0xff6666,1,.5));
  }
}

function buildNorthArrow(){
  // North arrow indicator on the ground plane — NOT rotated with building
  // It stays fixed so user can see True North vs building orientation
  const maxDim = Math.max(S.width, S.depth);
  const arrowDist = maxDim/2 + 8;
  const arrowLen = 4;
  // North is -Z in the scene coordinate (bearing 0)
  const northDir = new THREE.Vector3(0, 0, -1);
  const arrowOrigin = new THREE.Vector3(0, 0.15, -arrowDist);
  const arrow = new THREE.ArrowHelper(northDir, arrowOrigin, arrowLen, 0xff2222, 1.2, 0.7);
  scene.add(arrow);
  // Store for cleanup
  if(!window._northArrowParts) window._northArrowParts = [];
  window._northArrowParts.forEach(o => scene.remove(o));
  window._northArrowParts = [arrow];

  // "N" label
  const nLabel = spriteText('N', arrowOrigin.clone().add(new THREE.Vector3(0, 1.5, -arrowLen/2)), 0xff2222, 0.6);
  scene.add(nLabel);
  window._northArrowParts.push(nLabel);

  // Angle label if building is rotated
  if(S.mapBuildingAngle !== 0){
    const angLabel = spriteText('Building: ' + S.mapBuildingAngle + '°', new THREE.Vector3(0, 0.15, -arrowDist + arrowLen + 2), 0xffaa44, 0.5);
    scene.add(angLabel);
    window._northArrowParts.push(angLabel);
  }
}

function buildParticles(){
  if(particleSys){scene.remove(particleSys);particleSys.geometry.dispose();particleSys.material.dispose();particleSys=null}
  const N=720,pos=new Float32Array(N*3),cols=new Float32Array(N*3),meta=[];
  const w=S.width,d=S.depth,h=S.height;
  const ang=(S.windAngle+(S.R.angleOff||0))*Math.PI/180;
  const wx=-Math.sin(ang), wz=-Math.cos(ang); // wind travel direction
  const px=Math.cos(ang), pz=-Math.sin(ang);  // perpendicular to wind
  const spread=Math.max(w,d)*1.75;
  const windwardP=Math.abs(S.R?.faces?.windward?.p||0.85);
  const leewardP=Math.abs(S.R?.faces?.leeward?.p||0.45);
  const roofP=Math.max(Math.abs(S.R?.faces?.roof_ww?.p||0),Math.abs(S.R?.faces?.roof_lw?.p||0));
  const sideP=Math.max(
    Math.abs(S.R?.faces?.side1_zone1?.p||0),
    Math.abs(S.R?.faces?.side1_zone2?.p||0),
    Math.abs(S.R?.faces?.side1_zone3?.p||0)
  );
  const pressureGain=Math.min(2.2, Math.max(0.75, windwardP + roofP*0.55));
  const wakeGain=Math.min(2.1, Math.max(0.55, leewardP + roofP*0.35));
  const sideGain=Math.min(1.8, Math.max(0.45, sideP + windwardP*0.2));
  for(let i=0;i<N;i++){
    const upDist=Math.random()*30+4; // distance upwind
    const side=(Math.random()-.5)*spread;
    pos[i*3]  = -wx*upDist + px*side;
    pos[i*3+1]= Math.random()*Math.max(h*2.2,7);
    pos[i*3+2]= -wz*upDist + pz*side;
    cols[i*3]=0.56; cols[i*3+1]=0.86; cols[i*3+2]=1.0;
    meta.push({
      jitterX:(Math.random()-.5)*.06,
      jitterY:(Math.random()-.5)*.03,
      jitterZ:(Math.random()-.5)*.06,
      seed:Math.random()*Math.PI*2,
      pressureGain,
      wakeGain,
      sideGain
    });
  }
  const g=new THREE.BufferGeometry();
  g.setAttribute('position',new THREE.BufferAttribute(pos,3));
  g.setAttribute('color',new THREE.BufferAttribute(cols,3));
  particleSys=new THREE.Points(g,new THREE.PointsMaterial({size:.18,vertexColors:true,transparent:true,opacity:.72,blending:THREE.AdditiveBlending,depthWrite:false}));
  particleSys.userData.meta=meta;scene.add(particleSys);
}
function tickParticles(){
  if(!particleSys||!S.showParticles)return;
  const p=particleSys.geometry.attributes.position;
  const c=particleSys.geometry.attributes.color;
  const meta=particleSys.userData.meta||[];
  const w=S.width,d=S.depth,h=S.height;
  const ang=(S.windAngle+(S.R.angleOff||0))*Math.PI/180;
  const wx=-Math.sin(ang), wz=-Math.cos(ang); // wind travel direction
  const px=Math.cos(ang), pz=-Math.sin(ang);  // perpendicular
  const spread=Math.max(w,d)*1.75;
  for(let i=0;i<p.count;i++){
    const m=meta[i]||{jitterX:0,jitterY:0,jitterZ:0,seed:0,pressureGain:1,wakeGain:1,sideGain:1};
    let x=p.getX(i),y=p.getY(i),z=p.getZ(i);
    const downwind = x*wx + z*wz;
    const side = x*px + z*pz;
    const isFrontBand = downwind > -(d/2+5) && downwind < 1.5;
    const isRoofBand = downwind >= -1 && downwind <= d/2+4 && Math.abs(side) < w/2+2;
    const isWakeBand = downwind > d/2 && downwind < d/2+14 && Math.abs(side) < w/2+5;
    const nearSide = Math.abs(side) >= w/2-1 && Math.abs(side) <= w/2+5 && downwind > -(d/2+4) && downwind < d/2+6;

    let speed = .18 + m.pressureGain*.05;
    let vx=wx*speed + m.jitterX, vy=m.jitterY, vz=wz*speed + m.jitterZ;

    if(isFrontBand && y < h+2.5){
      const wallLift = (1 - Math.min(1, Math.abs(side)/(w/2+3))) * m.pressureGain;
      vx *= 0.58;
      vz *= 0.58;
      vy += 0.055 + wallLift*0.05;
    }
    if(isRoofBand){
      const roofBoost = (1 - Math.min(1, Math.abs(side)/(w/2+2))) * m.pressureGain;
      vx += wx*(0.045 + roofBoost*0.03);
      vz += wz*(0.045 + roofBoost*0.03);
      vy += 0.035 + roofBoost*0.035;
    }
    if(nearSide && y < h+2){
      const sgn = side>=0 ? 1 : -1;
      vx += px*sgn*(0.04 + m.sideGain*0.03);
      vz += pz*sgn*(0.04 + m.sideGain*0.03);
    }
    if(isWakeBand){
      const swirl = Math.sin((downwind-d/2)*0.65 + m.seed)*m.wakeGain;
      vx *= 0.55;
      vz *= 0.55;
      vx += px*swirl*0.045;
      vz += pz*swirl*0.045;
      vy += Math.cos((downwind-d/2)*0.55 + m.seed)*0.03*m.wakeGain;
    }
    x+=vx;y+=vy;z+=vz;
    // Reset if particle traveled too far downwind or out of bounds
    if(downwind>Math.max(w,d)+15||y<-1||y>h*3){
      const upDist=Math.random()*8+18;
      const side2=(Math.random()-.5)*spread;
      x=-wx*upDist+px*side2;
      y=Math.random()*Math.max(h*2.2,7);
      z=-wz*upDist+pz*side2;
    }
    p.setXYZ(i,x,y,z);

    const localSpeed = Math.sqrt(vx*vx + vy*vy + vz*vz);
    const hot = Math.min(1, Math.max(0, (localSpeed - 0.12) / 0.2));
    const wakeTint = isWakeBand ? 0.25 : 0;
    c.setXYZ(i,
      0.45 + hot*0.55,
      0.75 + (1-hot)*0.15 - wakeTint*0.2,
      1.0 - hot*0.65 - wakeTint*0.25
    );
  }
  p.needsUpdate=true;
  c.needsUpdate=true;
}

// ═══════════════════════════════════════════════
//   COMPASS
// ═══════════════════════════════════════════════
function setupCompass(){
  const cv=document.getElementById('compass-canvas');
  if(!cv) return;
  const compassRect=()=>cv.getBoundingClientRect();
  function angleFrom(cx,cy,ex,ey){return Math.atan2(ex-cx,-(ey-cy))*180/Math.PI}
  cv.addEventListener('mousedown',()=>compassDrag=true);
  window.addEventListener('mouseup',()=>compassDrag=false);
  window.addEventListener('mousemove',e=>{
    if(!compassDrag)return;
    const r=compassRect();S.windAngle=((angleFrom(r.left+r.width/2,r.top+r.height/2,e.clientX,e.clientY)%360)+360)%360;
    onInput();
  });
  cv.addEventListener('touchstart',e=>{compassDrag=true;e.preventDefault()},{passive:false});
  cv.addEventListener('touchend',()=>compassDrag=false);
  cv.addEventListener('touchmove',e=>{
    if(!compassDrag)return;e.preventDefault();
    const t=e.touches[0],r=compassRect();
    S.windAngle=((angleFrom(r.left+r.width/2,r.top+r.height/2,t.clientX,t.clientY)%360)+360)%360;
    onInput();
  },{passive:false});
}
function drawCompass(){
  const cv=document.getElementById('compass-canvas');
  if(!cv) return;
  const c=cv.getContext('2d');
  const cx=65,cy=65,R=52;
  c.clearRect(0,0,130,130);
  c.fillStyle='rgba(15,15,35,.8)';c.beginPath();c.arc(cx,cy,R+5,0,Math.PI*2);c.fill();
  c.strokeStyle='rgba(255,255,255,.15)';c.lineWidth=2;c.beginPath();c.arc(cx,cy,R,0,Math.PI*2);c.stroke();
  for(let i=0;i<36;i++){
    const a=i*10*Math.PI/180;const inner=i%9===0?R-12:R-6;
    c.strokeStyle=i%9===0?'rgba(255,255,255,.45)':'rgba(255,255,255,.12)';c.lineWidth=i%9===0?2:1;
    c.beginPath();c.moveTo(cx+Math.sin(a)*inner,cy-Math.cos(a)*inner);
    c.lineTo(cx+Math.sin(a)*R,cy-Math.cos(a)*R);c.stroke();
  }
  c.font='bold 12px Arial';c.textAlign='center';c.textBaseline='middle';
  c.fillStyle='#ff4444';c.fillText('N',cx,cy-R+10);
  c.fillStyle='#888';c.fillText('S',cx,cy+R-10);c.fillText('E',cx+R-10,cy);c.fillText('W',cx-R+10,cy);
  const a=S.windAngle*Math.PI/180;
  c.save();c.translate(cx,cy);c.rotate(a);
  c.fillStyle='#ff4444';c.beginPath();
  c.moveTo(0,-R+15);c.lineTo(6,0);c.lineTo(2,0);c.lineTo(2,R-20);
  c.lineTo(-2,R-20);c.lineTo(-2,0);c.lineTo(-6,0);c.closePath();c.fill();
  c.restore();
  c.fillStyle='#fff';c.beginPath();c.arc(cx,cy,3,0,Math.PI*2);c.fill();
  c.fillStyle='#ffaa44';c.font='10px Arial';c.fillText(Math.round(S.windAngle)+'°',cx,cy+R+9);
}

// ═══════════════════════════════════════════════
//   RESULTS UI (Structure tab)
// ═══════════════════════════════════════════════
function updateLegend(){
  const container = document.getElementById('kl-legend-swatches');
  if(!container) return;

  const r = S.height / Math.min(S.width, S.depth);
  const saShort = r <= 1 ? 'SA1/SA2' : 'SA3–SA5';
  const saFull = r <= 1 ? 'SA1/SA2 zones (r ≤ 1)' : 'SA3/SA4/SA5 zones (r > 1)';

  const zones = [
    {short:'3.0 RC1', full:'Kl = 3.0  (RC1, SA5)', color:'#DD2222'},
    {short:'2.0 RA2', full:'Kl = 2.0  (RA2, SA2/SA4)', color:'#EECC00'},
    {short:'1.5 WW', full:'Kl = 1.5  WA1 windward wall', color:'#22BB44'},
    {short:'1.5 RA', full:'Kl = 1.5  RA roof edge (teal)', color:'#2a9d8f'},
    {short:'1.5 SA', full:'Kl = 1.5  SA sidewall', color:'#4caf50'},
    {short:'1.0 MR', full:'Kl = 1.0  (MR, Other)', color:'#2288DD'}
  ];
  const esc = t => String(t).replace(/&/g,'&amp;').replace(/"/g,'&quot;');
  let html = '<span class="kl-legend-meta" title="r = h / min(b,d). '+esc(saFull)+'">r='
    + r.toFixed(2)+' · '+saShort+'</span>';
  zones.forEach(z=>{
    html += '<span class="kl-legend-item" title="'+esc(z.full)+'">'
      + '<span class="kl-legend-swatch" style="background:'+z.color+'"></span>'
      + '<span class="kl-legend-item-text">'+z.short+'</span></span>';
  });
  if(S.Kp !== '1.0'){
    const kpTip = S.Kp === 'auto'
      ? 'Kp from Table 5.8 (permeable cladding) — reduces negative p on roofs and side walls (Cl 5.4.5)'
      : 'Uniform Kp = '+S.Kp+' — reduces negative p on roofs and side walls (Cl 5.4.5)';
    const kpShort = S.Kp === 'auto' ? 'Kp Table 5.8' : 'Kp '+S.Kp;
    html += '<span class="kl-legend-kp" title="'+esc(kpTip)+'">'+kpShort+'</span>';
  }
  container.innerHTML = html;
}

// ═══════════════════════════════════════════════
//   8-DIRECTION MULTIPLIER TABS — polar compass UI
// ═══════════════════════════════════════════════
function switchDirTab(tab){
  S.activeDirTab = tab;
  dirPolarClosePopover();
  ['Md','Mzcat','Ms','Mt','Mlee','Vsit'].forEach(t=>{
    const btn = document.getElementById('dtab-'+t);
    if(btn) btn.classList.toggle('active', t===tab);
  });
  refreshDirectionalWindUI();
}

/** TC label for direction i (matches former table dropdown display). */
function tcLabelForDir(i){
  const tcVal = S.TC_dir[i];
  const zones = S.terrainZones && S.terrainZones[i];
  const hasMultiZone = zones && zones.length > 1;
  if(hasMultiZone){
    const distinctTCs = [...new Set(zones.map(z=>z.tc))];
    if(distinctTCs.length > 1){
      return 'TC ' + distinctTCs[0] + ' → TC ' + distinctTCs[distinctTCs.length-1];
    }
    return 'TC ' + distinctTCs[0];
  }
  if(tcVal == null || tcVal === '' || !Number.isFinite(Number(tcVal))) return '';
  return Number.isInteger(tcVal) ? 'TC' + tcVal : 'TC ' + Number(tcVal).toFixed(2);
}

function formatTcForReport(i){
  const tc = S.TC_dir[i];
  if(tc == null || tc === '' || !Number.isFinite(Number(tc))) return '—';
  return Number.isInteger(tc) ? 'TC' + tc : 'TC ' + Number(tc).toFixed(2);
}

function formatMzcatForReport(i, decimals){
  const d = decimals != null ? decimals : 3;
  const v = S.Mzcat[i];
  return (v != null && Number.isFinite(v)) ? v.toFixed(d) : '—';
}

function dirPolarThetaCenterRad(bearingDeg){
  return bearingDeg * Math.PI / 180 - Math.PI / 2;
}

/** Annular sector wedge from rInner to rOuter (variable per direction). */
function dirPolarSectorPath(cx, cy, rInner, rOuter, thetaCenterRad){
  const a0 = thetaCenterRad - Math.PI / 8;
  const a1 = thetaCenterRad + Math.PI / 8;
  const xi0 = cx + rInner * Math.cos(a0), yi0 = cy + rInner * Math.sin(a0);
  const xo0 = cx + rOuter * Math.cos(a0), yo0 = cy + rOuter * Math.sin(a0);
  const xo1 = cx + rOuter * Math.cos(a1), yo1 = cy + rOuter * Math.sin(a1);
  const xi1 = cx + rInner * Math.cos(a1), yi1 = cy + rInner * Math.sin(a1);
  return `M ${xi0.toFixed(2)} ${yi0.toFixed(2)} L ${xo0.toFixed(2)} ${yo0.toFixed(2)} A ${rOuter} ${rOuter} 0 0 1 ${xo1.toFixed(2)} ${yo1.toFixed(2)} L ${xi1.toFixed(2)} ${yi1.toFixed(2)} A ${rInner} ${rInner} 0 0 0 ${xi0.toFixed(2)} ${yi0.toFixed(2)} Z`;
}

function dirPolarValueBounds(vals, tab){
  let vmin, vmax;
  if(tab==='Mzcat'){
    const finite = vals.filter(v=>v!=null && Number.isFinite(v));
    if(finite.length === 0){ vmin = 0.85; vmax = 1.15; }
    else { vmin = Math.min(...finite); vmax = Math.max(...finite); }
  } else {
    vmin = Math.min(...vals); vmax = Math.max(...vals);
  }
  if(tab==='Vsit'){
    vmax = Math.max(vmax, 20);
    vmin = 0;
  } else if(tab!=='Mzcat'){
    if(!Number.isFinite(vmin) || !Number.isFinite(vmax)) vmin = 0.8, vmax = 1.2;
    if(Math.abs(vmax - vmin) < 1e-6){
      vmin = Math.min(vmin - 0.1, 0.85);
      vmax = Math.max(vmax + 0.1, 1.15);
    }
  } else {
    if(!Number.isFinite(vmin) || !Number.isFinite(vmax)) vmin = 0.85, vmax = 1.15;
    if(Math.abs(vmax - vmin) < 1e-6){
      vmin = Math.min(vmin - 0.1, 0.85);
      vmax = Math.max(vmax + 0.1, 1.15);
    }
  }
  return {vmin, vmax};
}

function dirPolarNormRadius(v, vmin, vmax, rIn, rOut){
  const t = (v - vmin) / (vmax - vmin);
  const n = 0.14 + Math.max(0, Math.min(1, t)) * 0.86;
  return rIn + n * (rOut - rIn);
}

function dirPolarClosePopover(){
  S.dirPolarPopoverIdx = null;
  const pop = document.getElementById('dir-polar-popover');
  if(pop){
    pop.innerHTML = '';
    pop.hidden = true;
  }
}

function dirPolarSectorClick(i){
  if(S.analysisLocked){
    toast('Unlock analysis to edit multipliers');
    return;
  }
  const tab = S.activeDirTab;
  if(isTerrainPolarTab(tab)){
    if(!currentUser){
      toast('Terrain multipliers are locked in this local build for now.');
      showAuthOverlay();
      return;
    }
    if(!isPaidPlan() && getTerrainPolarUseCount() >= TERRAIN_POLAR_FREE_USES){
      toast('Unlimited terrain multiplier edits are enabled in this local build.');
      openPaymentModal();
      dirPolarClosePopover();
      refreshDirectionalWindUI();
      return;
    }
    if(!isPaidPlan()) incrementTerrainPolarUseCount();
  }
  const dirs = S.dirs;
  const angles = [0,45,90,135,180,225,270,315];
  const pop = document.getElementById('dir-polar-popover');
  if(!pop) return;
  S.dirPolarPopoverIdx = i;

  const title = `${dirs[i]} (${angles[i]}°)`;
  let body = '';

  if(tab==='Vsit'){
    const v = S.Vsit_dir[i];
    body = `<div class="dir-polar-pop-readonly"><strong>V<sub>sit,β</sub></strong> = ${v.toFixed(1)} m/s</div>`;
  } else if(tab==='Mlee'){
    const m = S.Mlee[i] || 1;
    body = `<div class="dir-polar-pop-readonly"><strong>M<sub>lee</sub></strong> = ${m.toFixed(2)}</div>`;
  } else {
    const arrKey = tab==='Mzcat' ? 'Mzcat' : tab;
    const arr = S[arrKey];
    const val = arr[i];
    const safeVal = tab==='Mzcat'
      ? ((val != null && Number.isFinite(val)) ? val.toFixed(2) : '')
      : ((val != null && Number.isFinite(val)) ? val.toFixed(2) : '1.00');
    body = `<label class="dir-polar-pop-lbl" for="dir-polar-pop-val">${tab==='Md'?'M<sub>d</sub>':tab==='Mzcat'?'M<sub>z,cat</sub>':tab==='Ms'?'M<sub>s</sub>':'M<sub>t</sub>'}</label>`
      + `<input type="text" id="dir-polar-pop-val" class="dir-polar-pop-input" value="${safeVal}" inputmode="decimal" autocomplete="off" `
      + `onchange="dirPolarPopValueChange()" aria-label="Multiplier value for ${dirs[i]}">`;
    if(tab==='Mzcat'){
      const tcVal = S.TC_dir[i];
      const sel = (v)=>{
        if(v === '') return (tcVal == null || tcVal === '' || !Number.isFinite(Number(tcVal))) ? ' selected' : '';
        const n = Number(v);
        return (tcVal != null && Number.isFinite(Number(tcVal)) && Math.abs(Number(tcVal)-n) < 1e-6) ? ' selected' : '';
      };
      body += `<label class="dir-polar-pop-lbl" for="dir-polar-pop-tc">Terrain</label>`
        + `<select id="dir-polar-pop-tc" class="dir-polar-pop-select" onchange="dirPolarPopTcChange()" aria-label="Terrain category for ${dirs[i]}">`
        + `<option value=""${sel('')}>—</option>`
        + `<option value="1"${sel(1)}>TC1</option><option value="2"${sel(2)}>TC2</option><option value="2.5"${sel(2.5)}>TC2.5</option>`
        + `<option value="3"${sel(3)}>TC3</option><option value="4"${sel(4)}>TC4</option>`
        + `</select>`;
    }
  }

  pop.innerHTML = `<div class="dir-polar-pop-inner">`
    + `<div class="dir-polar-pop-head"><span>${title}</span>`
    + `<button type="button" class="dir-polar-pop-close" onclick="dirPolarClosePopover()" aria-label="Close">×</button></div>`
    + body
    + `</div>`;
  pop.hidden = false;
  if(isTerrainPolarTab(tab) && !isPaidPlan()) refreshDirectionalWindUI();
}

function dirPolarPopValueChange(){
  const i = S.dirPolarPopoverIdx;
  if(i==null) return;
  const tab = S.activeDirTab;
  const inp = document.getElementById('dir-polar-pop-val');
  if(!inp) return;
  const arrKey = tab==='Mzcat' ? 'Mzcat' : tab;
  updateDirVal(arrKey, i, inp.value);
}

function dirPolarPopTcChange(){
  const i = S.dirPolarPopoverIdx;
  if(i==null) return;
  const sel = document.getElementById('dir-polar-pop-tc');
  if(!sel) return;
  updateDirTC(i, sel.value);
}

function dirPolarSyncPopoverIfOpen(){
  const i = S.dirPolarPopoverIdx;
  if(i==null) return;
  const pop = document.getElementById('dir-polar-popover');
  if(!pop || pop.hidden) return;
  const tab = S.activeDirTab;
  const vInp = document.getElementById('dir-polar-pop-val');
  if(vInp && tab!=='Vsit' && tab!=='Mlee'){
    const arrKey = tab==='Mzcat' ? 'Mzcat' : tab;
    const arr = S[arrKey];
    if(tab==='Mzcat'){
      vInp.value = (arr && arr[i]!=null && Number.isFinite(arr[i])) ? arr[i].toFixed(2) : '';
    } else if(arr && arr[i]!=null && Number.isFinite(arr[i])){
      vInp.value = arr[i].toFixed(2);
    }
  }
  const tcSel = document.getElementById('dir-polar-pop-tc');
  if(tcSel && tab==='Mzcat'){
    const tcVal = S.TC_dir[i];
    const sel = (v)=>{
      if(v === '') return (tcVal == null || tcVal === '' || !Number.isFinite(Number(tcVal))) ? ' selected' : '';
      return (tcVal != null && Number.isFinite(Number(tcVal)) && Math.abs(Number(tcVal)-v) < 1e-6) ? ' selected' : '';
    };
    let h = `<option value=""${sel('')}>—</option>`;
    [[1,'TC1'],[2,'TC2'],[2.5,'TC2.5'],[3,'TC3'],[4,'TC4']].forEach(([v, lab])=>{
      h += `<option value="${v}"${sel(v)}>${lab}</option>`;
    });
    tcSel.innerHTML = h;
  }
}

function renderDirTable(){
  const el = document.getElementById('dir-table-content');
  if(!el) return;
  const tab = S.activeDirTab;
  const locked = S.analysisLocked;
  const terrainPolarBlocked = isTerrainPolarTab(tab) && !canUseTerrainPolarInteraction();
  const sectorLocked = locked || terrainPolarBlocked;
  const showMzMsLoading = (tab === 'Mzcat' || tab === 'Ms') && detectPendingOsm;
  /** Mt sampling is done when profiles + sub-direction Mh exist and site elevation is known (same markers as elev pipeline). */
  const mtElevSamplingDone = !!(S.detectedProfiles && Array.isArray(S.mhSub) && Number.isFinite(S.detectedSiteElev));
  const showMtLoading = tab === 'Mt' && detectPendingElev && !mtElevSamplingDone;
  if(showMzMsLoading || showMtLoading){
    let caption = 'Calculating multipliers…';
    if(tab === 'Mzcat') caption = 'Calculating M<sub>z,cat</sub>…';
    else if(tab === 'Ms') caption = 'Calculating M<sub>s</sub>…';
    else if(tab === 'Mt') caption = 'Calculating M<sub>t</sub>…';
    el.innerHTML = `<div class="dir-polar-wrap dir-polar-loading${locked || terrainPolarBlocked ? ' dir-polar-locked' : ''}">`
      + `<div class="dir-polar-loading-box" role="status" aria-live="polite" aria-busy="true">`
      + `<div class="dir-polar-loading-spinner" aria-hidden="true"></div>`
      + `<div class="dir-polar-loading-caption">${caption}</div>`
      + `</div></div>`;
    dirPolarSyncPopoverIfOpen();
    return;
  }

  const dirs = S.dirs;
  const angles = [0,45,90,135,180,225,270,315];
  const cx = 110, cy = 110;
  const rIn = 26, rMax = 92;

  let values;
  let quantityLabel;
  let valueDecimals = 2;
  if(tab==='Vsit'){
    values = S.Vsit_dir.slice();
    quantityLabel = 'V sit beta m/s';
    valueDecimals = 1;
  } else if(tab==='Mlee'){
    values = S.Mlee.map(m=>m||1);
    quantityLabel = 'M lee';
  } else {
    const arrKey = tab==='Mzcat' ? 'Mzcat' : tab;
    values = S[arrKey].slice();
    quantityLabel = tab==='Md' ? 'Md' : tab==='Mzcat' ? 'Mz cat' : tab==='Ms' ? 'Ms' : 'Mt combined';
  }

  const {vmin, vmax} = dirPolarValueBounds(values, tab);
  const ariaTab = tab==='Mzcat' ? 'Mz,cat' : tab==='Vsit' ? 'Vsit,beta' : tab;

  let paths = '';
  let texts = '';
  for(let i=0;i<8;i++){
    const v = values[i];
    const θ = dirPolarThetaCenterRad(angles[i]);
    const vNorm = (tab==='Mzcat' && (v==null || !Number.isFinite(v))) ? (vmin+vmax)/2 : v;
    const rOut = dirPolarNormRadius(vNorm, vmin, vmax, rIn, rMax);
    let fill = 'rgba(0,210,255,0.18)';
    let stroke = 'rgba(255,255,255,0.22)';
    if(tab==='Mlee' && (S.Mlee[i]||1) > 1){
      fill = 'rgba(76,175,80,0.35)';
      stroke = 'rgba(129,199,132,0.85)';
    } else if(tab==='Vsit'){
      fill = 'rgba(0,210,255,0.22)';
      stroke = 'rgba(0,210,255,0.45)';
    }
    const pathD = dirPolarSectorPath(cx, cy, rIn, rOut, θ);
    const cursor = sectorLocked ? 'default' : 'pointer';
    const titleNum = (tab==='Mzcat' && (v==null || !Number.isFinite(v))) ? '—' : (valueDecimals===1?v.toFixed(1):v.toFixed(2));
    const title = `${dirs[i]} ${angles[i]}°: ${titleNum}${tab==='Vsit'?' m/s':''}${tab==='Mzcat'?' · '+tcLabelForDir(i):''}`;
    paths += `<path class="dir-polar-sector" d="${pathD}" fill="${fill}" stroke="${stroke}" stroke-width="1.2" `
      + `data-dir-idx="${i}" style="cursor:${cursor}" `
      + `onclick="dirPolarSectorClick(${i})" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();dirPolarSectorClick(${i});}" `
      + `tabindex="${sectorLocked?-1:0}" role="button" aria-label="${title.replace(/"/g,'&quot;')}">`
      + `<title>${title.replace(/</g,'&lt;')}</title></path>`;

    const rLab = (rIn + rOut) / 2 + 2;
    const tx = cx + Math.cos(θ) * rLab;
    const ty = cy + Math.sin(θ) * rLab;
    const disp = (tab==='Mzcat' && (v==null || !Number.isFinite(v))) ? '—' : (valueDecimals===1 ? v.toFixed(1) : v.toFixed(2));
    texts += `<text x="${tx.toFixed(1)}" y="${ty.toFixed(1)}" text-anchor="middle" dominant-baseline="middle" `
      + `class="dir-polar-valtxt" style="pointer-events:none">${disp}</text>`;
  }

  const rCard = rMax + 22;
  for(let i=0;i<8;i++){
    const θ = dirPolarThetaCenterRad(angles[i]);
    const bx = cx + Math.cos(θ) * rCard;
    const by = cy + Math.sin(θ) * rCard;
    texts += `<text x="${bx.toFixed(1)}" y="${by.toFixed(1)}" text-anchor="middle" dominant-baseline="middle" `
      + `class="dir-polar-dirlab" style="pointer-events:none">${dirs[i]}</text>`;
  }

  let notes = '';
  if(tab==='Mlee'){
    if(S.leeZone) notes += `<div class="dir-polar-note">Zone: ${S.leeZone.name}</div>`;
    if(S.leeOverride) notes += `<button type="button" class="dir-polar-lee-btn" onclick="enableLeeZone()">Re-enable Lee Zone</button>`;
    notes += '<div class="dir-polar-note">M<sub>lee</sub> is folded into M<sub>t</sub> on the Mt tab; V<sub>sit</sub> uses M<sub>t</sub> once (Cl 4.4.3 NZ lee zones).</div>';
  }
  if(tab==='Md' && mdOverrideKlDesign()){
    notes += '<div class="dir-polar-note">With Local Pressure Zones on: design uses M<sub>d</sub> = 1.0 (Cl 3.3(b), B2/C/D). Values shown are Table 3.2 reference.</div>';
  }
  if(tab==='Mt'){
    notes += '<div class="dir-polar-note">M<sub>t</sub> from M<sub>h</sub> and M<sub>lee</sub>; used once in V<sub>sit</sub>.</div>';
  }
  if(tab==='Mzcat'){
    notes += '<div class="dir-polar-hint">Click a sector to edit M<sub>z,cat</sub> and terrain category.</div>';
  } else if(tab!=='Vsit' && tab!=='Mlee'){
    notes += '<div class="dir-polar-hint">Click a sector to edit.</div>';
  } else {
    notes += '<div class="dir-polar-hint">Read-only; use other tabs to edit inputs.</div>';
  }
  if(isTerrainPolarTab(tab) && terrainPolarBlocked){
    if(!currentUser) notes += '<div class="dir-polar-note">Terrain multipliers are read-only in this local build for now.</div>';
    else notes += '<div class="dir-polar-note">Unlimited terrain multiplier edits are enabled in this local build.</div>';
  }

  el.innerHTML = `<div class="dir-polar-wrap${locked||terrainPolarBlocked?' dir-polar-locked':''}">`
    + `<svg class="dir-polar-svg" viewBox="0 0 220 220" role="img" aria-label="Directional ${quantityLabel} for 8 bearings, tab ${ariaTab}">`
    + `<circle cx="${cx}" cy="${cy}" r="${rIn-2}" fill="rgba(0,0,0,0.2)" stroke="rgba(255,255,255,0.12)" stroke-width="1"/>`
    + `<g class="dir-polar-sectors">${paths}</g>`
    + `<g class="dir-polar-labels">${texts}</g>`
    + `</svg>${notes}</div>`;

  dirPolarSyncPopoverIfOpen();
}

function updateDirVal(key, idx, val){
  const trimmed = String(val).trim();
  if(key==='Mzcat' && trimmed === ''){
    S.TC_dir[idx] = null;
    S.Mzcat[idx] = null;
    calc();
    refreshDirectionalWindUI();
    if(S.overlayTC) drawTCZones();
    if(S.overlayMs) drawMsOverlay();
    if(S.overlayMt) drawMtOverlay();
    return;
  }
  const v = parseFloat(val)||1;
  if(key==='Mt'){
    const el = (S.detectedSiteElev !== undefined && S.detectedSiteElev !== null && !isNaN(S.detectedSiteElev)) ? S.detectedSiteElev : 0;
    S.Mt_hill[idx] = combinedMtTargetToMhill(v, S.Mlee[idx]||1, S.region, el);
  } else {
    S[key][idx] = v;
  }
  calc();
  refreshDirectionalWindUI();
  if(key==='Ms' && S.overlayMs) drawMsOverlay();
  if(key==='Mt' && S.overlayMt) drawMtOverlay();
}
function updateDirTC(idx, val){
  if(val === '' || val == null){
    S.TC_dir[idx] = null;
    S.Mzcat[idx] = null;
  } else {
    const tc = parseFloat(val);
    if(!Number.isFinite(tc)){
      S.TC_dir[idx] = null;
      S.Mzcat[idx] = null;
    } else {
      S.TC_dir[idx] = tc;
      const h = S.R.h || S.height;
      S.Mzcat[idx] = mzCat(h, tc);
    }
  }
  calc();
  refreshDirectionalWindUI();
  if(S.overlayTC) drawTCZones();
}
/** Keep S.lat/S.lng and hidden inputs aligned with the site pin on the map (instant Detect/Clear without re-clicking the map). */
function syncSiteCoordsFromMapPin(){
  if(typeof mapMarker !== 'undefined' && mapMarker && typeof mapMarker.getLatLng === 'function'){
    const ll = mapMarker.getLatLng();
    if(ll && Number.isFinite(ll.lat) && Number.isFinite(ll.lng)){
      S.lat = ll.lat;
      S.lng = ll.lng;
      const la = document.getElementById('inp-lat');
      const ln = document.getElementById('inp-lng');
      if(la) la.value = S.lat.toFixed(6);
      if(ln) ln.value = S.lng.toFixed(6);
      return;
    }
  }
  const la = document.getElementById('inp-lat');
  const ln = document.getElementById('inp-lng');
  if(la){
    const p = parseFloat(la.value);
    if(Number.isFinite(p)) S.lat = p;
  }
  if(ln){
    const p = parseFloat(ln.value);
    if(Number.isFinite(p)) S.lng = p;
  }
}

function detectMultipliers(){
  syncSiteCoordsFromMapPin();
  invalidateInFlightSiteDetect();
  toast('Detecting multipliers from map data…');
  autoDetectAllMultipliers();
}

/** Landuse PIP: bucket polygon indices by lat/lng grid to avoid O(points × all polys). */
function buildLanduseSpatialBuckets(landusePolys, bucketDeg){
  const buckets = Object.create(null);
  landusePolys.forEach((poly, idx)=>{
    const i0 = Math.floor(poly.minLat / bucketDeg), i1 = Math.floor(poly.maxLat / bucketDeg);
    const j0 = Math.floor(poly.minLng / bucketDeg), j1 = Math.floor(poly.maxLng / bucketDeg);
    for(let i=i0;i<=i1;i++){
      for(let j=j0;j<=j1;j++){
        const k = i+'_'+j;
        if(!buckets[k]) buckets[k] = [];
        buckets[k].push(idx);
      }
    }
  });
  return buckets;
}

function landusePolysNearPoint(tp, buckets, bucketDeg, landusePolys){
  const bi = Math.floor(tp.lat / bucketDeg);
  const bj = Math.floor(tp.lng / bucketDeg);
  const seen = new Set();
  const out = [];
  for(let di=-1;di<=1;di++)for(let dj=-1;dj<=1;dj++){
    const arr = buckets[(bi+di)+'_'+(bj+dj)];
    if(!arr) continue;
    for(const idx of arr){
      if(seen.has(idx)) continue;
      seen.add(idx);
      const poly = landusePolys[idx];
      if(tp.lat < poly.minLat || tp.lat > poly.maxLat || tp.lng < poly.minLng || tp.lng > poly.maxLng) continue;
      out.push(poly);
    }
  }
  return out;
}

// ═══════════════════════════════════════════════
// ═══════════════════════════════════════════════
//  ELEVATION DATA FETCHER
//  Production: same-origin POST to window.CW_ELEVATION_API_URL when signed in — server may use Google + Redis.
//  Guests: Open-Meteo → elevation-api.eu in-browser (no server/Google).
//  Optional IndexedDB cache (terrain grid only): opts.elevIdb { lat, lng, fast }
// ═══════════════════════════════════════════════

const CW_ELEV_IDB_NAME = 'cw_detection_v1';
const CW_ELEV_IDB_STORE = 'elev_grids';
const CW_ELEV_IDB_TTL_MS = 7 * 24 * 60 * 60 * 1000;

/** Point count: lib/terrain-samples.js (cwTerrainElevPointCount). */

function cwElevIdbKey(lat, lng, fast){
  return 'v1|'+Number(lat).toFixed(5)+'|'+Number(lng).toFixed(5)+'|'+(fast ? 'fast' : 'full');
}

async function cwElevIdbOpen(){
  if(typeof indexedDB === 'undefined') return null;
  return new Promise((res, rej)=>{
    const r = indexedDB.open(CW_ELEV_IDB_NAME, 1);
    r.onerror = ()=>rej(r.error);
    r.onsuccess = ()=>res(r.result);
    r.onupgradeneeded = (e)=>{
      const db = e.target.result;
      if(!db.objectStoreNames.contains(CW_ELEV_IDB_STORE)) db.createObjectStore(CW_ELEV_IDB_STORE);
    };
  });
}

async function cwElevIdbGet(key){
  try{
    const db = await cwElevIdbOpen();
    if(!db) return null;
    return new Promise((res)=>{
      const tx = db.transaction(CW_ELEV_IDB_STORE, 'readonly');
      const req = tx.objectStore(CW_ELEV_IDB_STORE).get(key);
      req.onsuccess = ()=>{
        const row = req.result;
        if(!row || !Array.isArray(row.elevations)) return res(null);
        if(Date.now() - row.savedAt > CW_ELEV_IDB_TTL_MS) return res(null);
        res(row.elevations);
      };
      req.onerror = ()=>res(null);
    });
  } catch(e){ return null; }
}

async function cwElevIdbPut(key, elevations){
  try{
    const db = await cwElevIdbOpen();
    if(!db) return;
    return new Promise((res)=>{
      const tx = db.transaction(CW_ELEV_IDB_STORE, 'readwrite');
      tx.objectStore(CW_ELEV_IDB_STORE).put({ elevations, savedAt: Date.now() }, key);
      tx.oncomplete = ()=>res();
      tx.onerror = ()=>res();
    });
  } catch(e){}
}

/**
 * @param {number[]} latArr
 * @param {number[]} lngArr
 * @param {{ openMeteoBatchGapMs?: number, euBatchGapMs?: number, elevIdb?: { lat:number, lng:number, fast:boolean }, elevationRefineMeta?: { nSub:number, nDist:number } }} [opts] — gaps between batch requests (Open-Meteo 429 avoidance; use lower values for fast detect)
 */
function applyElevRefineIfNeeded(arr, opts){
  if(!opts || !opts.elevationRefineMeta || !Array.isArray(arr)) return arr;
  const { nSub, nDist } = opts.elevationRefineMeta;
  const fn = typeof globalThis !== 'undefined' && globalThis.refineRadialElevationsMedian;
  if(typeof fn !== 'function' || arr.length !== 1 + nSub * nDist) return arr;
  return fn(arr, nSub, nDist);
}

async function fetchElevBatchGlobal(latArr, lngArr, opts){
  opts = opts || {};
  const idbMeta = opts.elevIdb;
  const terrainIdbKey =
    idbMeta &&
    typeof globalThis.cwTerrainElevPointCount === 'function' &&
    globalThis.cwTerrainElevPointCount(!!idbMeta.fast) === latArr.length
      ? cwElevIdbKey(idbMeta.lat, idbMeta.lng, !!idbMeta.fast)
      : null;
  if(terrainIdbKey){
    const cached = await cwElevIdbGet(terrainIdbKey);
    if(cached && cached.length === latArr.length && cached.every(e => Number.isFinite(Number(e)))){
      return cached.map(e => Number(e));
    }
  }

  const authHdrs = await cwDetectionAuthHeaders();
  const elevApi = typeof window !== 'undefined' && window.CW_ELEVATION_API_URL;
  // Server route uses Google Elevation when configured — only call it when signed in (guests use Open-Meteo in-browser).
  if(getActiveFirebaseUser() && elevApi && String(elevApi).trim()){
    const lats = Array.from(latArr, Number);
    const lngs = Array.from(lngArr, Number);
    if(lats.length !== lngs.length || lats.length === 0) throw new Error('Invalid lat/lng arrays');
    try{
      const hdrs = { 'Content-Type': 'application/json', ...authHdrs };
      const body = { lats, lngs };
      if(opts.elevationRefineMeta && opts.elevationRefineMeta.nSub && opts.elevationRefineMeta.nDist){
        body.refine = true;
        body.nSub = opts.elevationRefineMeta.nSub;
        body.nDist = opts.elevationRefineMeta.nDist;
      }
      const resp = await fetch(String(elevApi).trim(), {
        method: 'POST',
        headers: hdrs,
        body: JSON.stringify(body),
      });
      const data = await resp.json().catch(()=>null);
      if(resp.ok && data && Array.isArray(data.elevations) && data.elevations.length === lats.length){
        const out = data.elevations.map(e => Number(e));
        if(out.every(e => Number.isFinite(e))){
          let refined = out;
          if(!data.refined){
            refined = applyElevRefineIfNeeded(out, opts);
          }
          if(terrainIdbKey) await cwElevIdbPut(terrainIdbKey, refined);
          return refined;
        }
      }
      if(resp.status === 401){
        console.warn('Elevation API returned 401 — falling back to Open-Meteo / elevation-api.eu');
      } else {
        console.warn('Elevation API failed or invalid response (HTTP '+(resp && resp.status)+') — using direct providers');
      }
    } catch(e){
      console.warn('Elevation API error — using direct providers:', e && e.message);
    }
  }
  const gapOpenMeteo = opts.openMeteoBatchGapMs !== undefined ? opts.openMeteoBatchGapMs : 280;
  const gapEu = opts.euBatchGapMs !== undefined ? opts.euBatchGapMs : 120;
  const result = new Array(latArr.length);
  const BATCH = 100;
  let useOpenMeteo = true;

  function fillFromMeteoElev(start, end, elevationArr){
    const n = end - start;
    if(!Array.isArray(elevationArr) || elevationArr.length !== n) return false;
    for(let j=0; j<n; j++){
      const ev = elevationArr[j];
      if(ev == null || (typeof ev === 'number' && !Number.isFinite(ev))) return false;
      result[start+j] = Number(ev);
    }
    return true;
  }

  async function fetchOpenMeteoBatch(start, end){
    const bLats = latArr.slice(start,end).map(l=>Number(l).toFixed(6)).join(',');
    const bLngs = lngArr.slice(start,end).map(l=>Number(l).toFixed(6)).join(',');
    const url = `https://api.open-meteo.com/v1/elevation?latitude=${bLats}&longitude=${bLngs}`;
    const maxAttempts = 8;
    for(let attempt=0; attempt<maxAttempts; attempt++){
      try {
        const resp = await fetch(url);
        if(resp.status === 429){
          console.warn('Open-Meteo 429 — switching to fallback elevation provider');
          return false;
        }
        if(!resp.ok && resp.status >= 500 && attempt < maxAttempts - 1){
          await new Promise(r=>setTimeout(r, 800 * (attempt + 1)));
          continue;
        }
        const data = await resp.json().catch(()=>null);
        if(resp.ok && data && data.error !== true && fillFromMeteoElev(start, end, data.elevation)) return true;
        return false;
      } catch(e){
        if(attempt === maxAttempts - 1) console.warn('Open-Meteo error (batch '+start+'):', e.message);
        await new Promise(r=>setTimeout(r, 600 * (attempt + 1)));
      }
    }
    return false;
  }

  async function fetchElevationApiEuBatch(start, end){
    const n = end - start;
    const pts = [];
    for(let j=start; j<end; j++) pts.push([Number(latArr[j]), Number(lngArr[j])]);
    const euUrl = 'https://www.elevation-api.eu/v1/elevation?pts='+encodeURIComponent(JSON.stringify(pts));
    const resp = await fetch(euUrl);
    const data = await resp.json().catch(()=>null);
    if(!resp.ok || !data || !Array.isArray(data.elevations) || data.elevations.length !== n) return false;
    for(let j=0; j<n; j++){
      const ev = data.elevations[j];
      if(ev == null || !Number.isFinite(Number(ev))) return false;
      result[start+j] = Number(ev);
    }
    return true;
  }

  async function fetchOpenElevationBatch(start, end){
    const n = end - start;
    const pts = [];
    for(let j=start; j<end; j++) pts.push(Number(latArr[j]).toFixed(6)+','+Number(lngArr[j]).toFixed(6));
    const url = 'https://api.open-elevation.com/api/v1/lookup?locations=' + encodeURIComponent(pts.join('|'));
    const resp = await fetch(url);
    const data = await resp.json().catch(()=>null);
    if(!resp.ok || !data || !Array.isArray(data.results) || data.results.length !== n) return false;
    for(let j=0; j<n; j++){
      const ev = data.results[j] && data.results[j].elevation;
      if(ev == null || !Number.isFinite(Number(ev))) return false;
      result[start+j] = Number(ev);
    }
    return true;
  }

  for(let start=0; start<latArr.length; start+=BATCH){
    const end = Math.min(start+BATCH, latArr.length);
    if(useOpenMeteo && start > 0) await new Promise(r=>setTimeout(r, 280));

    if(useOpenMeteo){
      const ok = await fetchOpenMeteoBatch(start, end);
      if(ok) continue;
      console.warn('Open-Meteo failed (batch '+start+'), switching to elevation-api.eu');
      useOpenMeteo = false;
      start -= BATCH;
      continue;
    }
    if(start > 0) await new Promise(r=>setTimeout(r, 120));
    let ok = await fetchElevationApiEuBatch(start, end);
    if(!ok) ok = await fetchOpenElevationBatch(start, end);
    if(!ok) throw new Error('Elevation query failed (fallback)');
  }
  if(result.some(e => e === undefined)) throw new Error('Incomplete elevation data');
  const outFinal = applyElevRefineIfNeeded(result.slice(), opts);
  if(terrainIdbKey) await cwElevIdbPut(terrainIdbKey, outFinal);
  return outFinal;
}

// computeMhFromProfiles — AS/NZS 1170.2 Cl 4.4.2 — see lib/mh-topography.js

// ── Terrain TC: footprint vs sector/band (not centroid-only) ──
function pointInPolyLatLng(testLat, testLng, ring){
  let inside = false;
  for(let i=0, j=ring.length-1; i<ring.length; j=i++){
    const yi = ring[i][0], xi = ring[i][1];
    const yj = ring[j][0], xj = ring[j][1];
    if(((yi > testLat) !== (yj > testLat)) &&
       (testLng < (xj-xi)*(testLat-yi)/(yj-yi)+xi)){
      inside = !inside;
    }
  }
  return inside;
}

/**
 * Allocate building footprint area to each (sector, distance band) by grid-sampling
 * the polygon in site-local metres. Large buildings spanning multiple bands get
 * realistic coverage per band instead of 100% in the centroid band.
 */
function distributeFootprintToBands(footprintLatLng, areaM2, siteLat, siteLng, mPerDegLat, mPerDegLng, bandEdges, maxDist, centroidDistM, centroidSector){
  const nBands = bandEdges.length - 1;
  const alloc = Array.from({length:8}, ()=>Array(nBands).fill(0));
  const defaultArea = Math.max(areaM2 || 0, 64);
  if(!footprintLatLng || footprintLatLng.length < 3 || areaM2 < 1){
    const bi = (()=>{
      for(let b=0;b<nBands;b++){
        if(centroidDistM >= bandEdges[b] && centroidDistM < Math.min(bandEdges[b+1], maxDist)) return b;
      }
      return Math.max(0, nBands-1);
    })();
    alloc[centroidSector % 8][bi] += defaultArea;
    return alloc;
  }
  let minX=Infinity, maxX=-Infinity, minY=Infinity, maxY=-Infinity;
  footprintLatLng.forEach(([la, lo])=>{
    const x = (lo-siteLng)*mPerDegLng, y = (la-siteLat)*mPerDegLat;
    minX = Math.min(minX,x); maxX = Math.max(maxX,x);
    minY = Math.min(minY,y); maxY = Math.max(maxY,y);
  });
  const pad = 3;
  minX -= pad; maxX += pad; minY -= pad; maxY += pad;
  const step = Math.max(2, Math.min(12, Math.sqrt(Math.max(areaM2, 25)) / 18));
  const hits = [];
  for(let x = minX; x <= maxX; x += step){
    for(let y = minY; y <= maxY; y += step){
      const la = siteLat + y/mPerDegLat, lo = siteLng + x/mPerDegLng;
      if(!pointInPolyLatLng(la, lo, footprintLatLng)) continue;
      const dist = Math.hypot(x, y);
      if(dist < 0.5 || dist >= maxDist) continue;
      let bearing = Math.atan2(x, y) * 180 / Math.PI;
      bearing = ((bearing % 360) + 360) % 360;
      const sec = Math.round(bearing / 45) % 8;
      let bi = -1;
      for(let b=0;b<nBands;b++){
        const outer = Math.min(bandEdges[b+1], maxDist);
        if(dist >= bandEdges[b] && dist < outer){ bi = b; break; }
      }
      if(bi >= 0) hits.push({ sec, bi });
    }
  }
  if(hits.length === 0){
    let bi = 0;
    for(let b=0;b<nBands;b++){
      if(centroidDistM >= bandEdges[b] && centroidDistM < Math.min(bandEdges[b+1], maxDist)){ bi = b; break; }
    }
    alloc[centroidSector % 8][bi] += areaM2;
    return alloc;
  }
  const per = areaM2 / hits.length;
  hits.forEach(({sec, bi})=>{ alloc[sec][bi] += per; });
  return alloc;
}

/** True if site pin lies inside a mapped building footprint (roof), not on road/yard. */
function siteInsideAnyBuildingFootprint(siteLat, siteLng, buildingsList){
  if(!buildingsList || !buildingsList.length) return false;
  for(const bl of buildingsList){
    const fp = bl.footprint;
    if(!fp || fp.length < 3) continue;
    if(pointInPolyLatLng(siteLat, siteLng, fp)) return true;
  }
  return false;
}

function applyHeightOverridesToBuildings(buildingsList, hRef){
  const href = Math.max(hRef || 1, 0.01);
  buildingsList.forEach(b=>{
    if(b.heightInferred != null) b.height = b.heightInferred;
    b.heightScore = Math.min(1.0, b.height / href);
  });
}

/**
 * Recompute TC zones + Ms from cached building list (heights from tags/defaults/edits).
 * Mutates tcArr, msArr; sets S.terrainZones, S.Mzcat, S.shieldingDetails.
 */
function recomputeTerrainAndShieldingFromBuildings(buildingsList, ctx, tcArr, msArr){
  const { lat, lng, h, mPerDegLat, mPerDegLng, bandCoverage, luBandEdges, sectorBuildings, sectorWater, sectorOpen, shieldDist } = ctx;
  const siteInsideBuilding = siteInsideAnyBuildingFootprint(lat, lng, buildingsList);
  const avgDistXa = Math.max(500, 40 * h);
  const bandEdges = luBandEdges;
  const nBandsTC = bandEdges.length - 1;
  const bandAllocArea = Array.from({length:8}, ()=>Array(nBandsTC).fill(0));
  const bandHeightWeighted = Array.from({length:8}, ()=>Array(nBandsTC).fill(0));
  const bandBldgCount = Array.from({length:8}, ()=>Array(nBandsTC).fill(0));
  /** Max height among buildings contributing footprint here — catches towers when avg is diluted by car parks / large annulus */
  const bandMaxHeight = Array.from({length:8}, ()=>Array(nBandsTC).fill(0));
  buildingsList.forEach(bl=>{
    const mat = distributeFootprintToBands(bl.footprint, bl.area, lat, lng, mPerDegLat, mPerDegLng, bandEdges, avgDistXa, bl.distance, bl.sectorIdx);
    for(let s=0;s<8;s++){
      for(let bb=0;bb<nBandsTC;bb++){
        const a = mat[s][bb];
        if(a > 0.5){
          bandAllocArea[s][bb] += a;
          bandHeightWeighted[s][bb] += bl.height * a;
          bandBldgCount[s][bb]++;
          if(bl.height > bandMaxHeight[s][bb]) bandMaxHeight[s][bb] = bl.height;
        }
      }
    }
  });

  S.terrainZones = [];

  for(let i=0;i<8;i++){
    const zones = [];
    let sumMzW = 0, sumW = 0;

    for(let b=0; b<bandEdges.length-1; b++){
      const inner = bandEdges[b];
      const outer = Math.min(bandEdges[b+1], avgDistXa);
      if(inner >= avgDistXa) break;
      const bandWidth = outer - inner;
      if(bandWidth <= 0) continue;

      const bCount = bandBldgCount[i][b];
      const footprintArea = bandAllocArea[i][b];

      const bandArea = Math.PI * (outer*outer - inner*inner) / 8;
      // Z1 sector wedge (~157 m²) is tiny: same absolute roof m² yields huge % and density vs Z2+.
      // Floor the effective annulus area for TC thresholds so inner rings align better with outer bands.
      const Z1_EFFECTIVE_BAND_AREA_MIN_M2 = 400;
      const effectiveBandAreaForTc = (inner < 20)
        ? Math.max(bandArea, Z1_EFFECTIVE_BAND_AREA_MIN_M2)
        : bandArea;
      const densityPerHa = effectiveBandAreaForTc > 0 ? bCount / (effectiveBandAreaForTc / 10000) : 0;
      const footprintCoverage = effectiveBandAreaForTc > 0 ? (footprintArea / effectiveBandAreaForTc) * 100 : 0;

      const avgHeight = footprintArea > 0.01 ? bandHeightWeighted[i][b] / footprintArea : 0;
      const bandMaxH = bandMaxHeight[i][b];

      const cov = bandCoverage[i][b];
      const hasParking = cov.parking >= 2;
      const hasUrbanLanduse = cov.urban >= 2 && !hasParking;
      const hasForestLanduse = cov.forest >= 2;
      const hasWaterLanduse = cov.water >= 2;
      const hasOpenLanduse = cov.open >= 2;

      const isImmediateBand = inner < 25;
      const lowFootprintAtSite = footprintCoverage < 10 && bCount <= 2;
      let tc;
      if(hasWaterLanduse && !hasUrbanLanduse && bCount === 0){
        tc = 1;
      } else if(isImmediateBand && lowFootprintAtSite){
        tc = 2;
      } else if(hasParking && footprintCoverage < 12 && bCount < 5 && footprintArea < 2800){
        tc = bCount > 2 ? 2.5 : 2;
      } else if(hasOpenLanduse && !hasUrbanLanduse && footprintCoverage < 11){
        if(bCount > 0) tc = 2;
        else           tc = 2;
      } else if(hasUrbanLanduse || (footprintCoverage > 4 && bCount > 0)){
        const tallDense = densityPerHa >= 5 && (avgHeight >= 10 || bandMaxH >= 15);
        const highCover = footprintCoverage >= 20;
        const heavyFootprint = footprintArea >= 900;
        // Any ≥12 m fabric in the annulus: with high cover, treat as TC4 (mean height often low beside towers / car parks)
        if(tallDense || (highCover && (bCount === 0 || avgHeight >= 10 || bandMaxH >= 12))){
          tc = 4;
        } else if((highCover || heavyFootprint) && bCount > 0 && avgHeight < 10){
          tc = bandMaxH >= 12 ? 4 : 3;
        } else if(footprintCoverage >= 9 || footprintArea >= 500 || densityPerHa >= 4 ||
                  (hasUrbanLanduse && (densityPerHa >= 2 || footprintCoverage >= 6))){
          tc = (bandMaxH >= 18 && footprintCoverage >= 12 && hasUrbanLanduse) ? 4 : 3;
        } else if(footprintCoverage >= 3.5 || densityPerHa >= 1 || (hasUrbanLanduse && footprintCoverage >= 4)){
          tc = (bandMaxH >= 22 && footprintArea >= 350) ? 4 : 2.5;
        } else {
          tc = 2;
        }
      } else if(hasForestLanduse){
        tc = 2.5;
      } else if(densityPerHa >= 8 && (avgHeight >= 10 || bandMaxH >= 15)){
        tc = 4;
      } else if(densityPerHa >= 3.5){
        tc = (bandMaxH >= 18 && bCount > 0 && (footprintCoverage >= 8 || footprintArea >= 400)) ? 4 : 3;
      } else if(densityPerHa >= 1.2){
        tc = 2.5;
      } else if(bCount > 0){
        tc = 2;
      } else {
        tc = hasWaterLanduse ? 1 : 2;
      }

      // ── Z1 (0–20 m): default to open / low roughness (TC2–2.5) ──
      // OSM footprint bleed + landuse noise otherwise labels carriageways, yards and “lots” as TC3/4.
      // TC3/4 here only if this 20 m sector wedge is clearly roof-dominated (not façade spill).
      if(inner < 20 && !siteInsideBuilding){
        const base = tc;
        tc = Math.min(base, 2.5);
        // TC3: strong majority of wedge is building + landuse agrees + some height signal
        const z1OkTC3 = footprintCoverage >= 58 && footprintArea >= 88 && avgHeight >= 8 &&
          (cov.urban >= 3 || (cov.urban >= 2 && bandMaxH >= 14 && footprintCoverage >= 52));
        if(z1OkTC3) tc = Math.min(base, 3);
        // TC4 in Z1 only under roof — avoids TC4 on carriageway from façade clip + % inflation
        if(siteInsideBuilding){
          const z1OkTC4 = footprintCoverage >= 48 && footprintArea >= 120 && avgHeight >= 11 &&
            bandMaxH >= 14 && cov.urban >= 2;
          if(z1OkTC4) tc = Math.min(base, 4);
        }
        // Strong open signals (carpark / grass / low urban samples) — stay at TC2
        if((cov.parking >= 2 || (hasOpenLanduse && cov.urban <= 2) || cov.urban <= 1) && footprintCoverage < 55){
          tc = Math.min(tc, 2);
        }
      }

      // Outer rings: still TC3 but max height in annulus is high-rise (mean diluted) — bump to TC4
      if(inner >= 20 && tc < 4 && bandMaxH >= 16 && bCount > 0){
        const substantial = footprintArea >= 300 || footprintCoverage >= 7;
        if(substantial && (bandMaxH >= 22 || (bandMaxH >= 16 && footprintCoverage >= 14 && hasUrbanLanduse))){
          tc = 4;
        }
      }

      const mz = mzCat(h, tc);
      sumMzW += mz * bandWidth;
      sumW += bandWidth;
      zones.push({ from: inner, to: outer, tc, mz, width: bandWidth });
    }

    // Monotonicity: Z1 (0–20 m) TC should not exceed Z2 (20–50 m) when pin is not on a roof
    if(!siteInsideBuilding && zones.length >= 2 && bandEdges[0] === 0 && bandEdges[1] === 20){
      const z0 = zones[0];
      const z1b = zones[1];
      if(z0.tc > z1b.tc){
        const oldMz = z0.mz;
        z0.tc = z1b.tc;
        z0.mz = mzCat(h, z0.tc);
        sumMzW += (z0.mz - oldMz) * z0.width;
      }
    }

    if(sumW > 0){
      const avgMz = sumMzW / sumW;
      S.Mzcat[i] = avgMz;
      let tcLo = 1, tcHi = 4;
      for(let iter=0; iter<20; iter++){
        const tcMid = (tcLo + tcHi) / 2;
        if(mzCat(h, tcMid) > avgMz) tcLo = tcMid; else tcHi = tcMid;
      }
      tcArr[i] = parseFloat(((tcLo + tcHi) / 2).toFixed(2));
    } else {
      const b = sectorBuildings[i];
      if(sectorWater[i] && b<3)      tcArr[i] = 1;
      else if(sectorOpen[i] && b<5)  tcArr[i] = 2;
      else if(b < 3)                 tcArr[i] = 2;
      else if(b < 10)                tcArr[i] = 2.5;
      else if(b < 30)                tcArr[i] = 3;
      else                           tcArr[i] = 4;
    }
    S.terrainZones.push(zones);
  }

  const msTable = [[1.5, 0.7], [3.0, 0.8], [6.0, 0.9], [12.0, 1.0]];
  function tableLookupMs(sParam){
    if(sParam <= msTable[0][0]) return msTable[0][1];
    if(sParam >= msTable[msTable.length-1][0]) return msTable[msTable.length-1][1];
    for(let j=0; j<msTable.length-1; j++){
      if(sParam <= msTable[j+1][0]){
        const t = (sParam - msTable[j][0]) / (msTable[j+1][0] - msTable[j][0]);
        return msTable[j][1] + t * (msTable[j+1][1] - msTable[j][1]);
      }
    }
    return 1.0;
  }

  S.shieldingDetails = [];

  for(let i=0;i<8;i++){
    const sectorBldgs = buildingsList.filter(b=> b.sectorIdx===i && b.distance<=shieldDist && b.distance>5);

    const qualifying = sectorBldgs.filter(b=> b.height >= h);
    const ns = qualifying.length;

    if(ns === 0){
      msArr[i] = 1.0;
      S.shieldingDetails[i] = { ns:0, ls:0, hs:0, bs:0, s:Infinity, Ms:1.0 };
      continue;
    }

    const hs = qualifying.reduce((sum,b) => sum + b.height, 0) / ns;
    const bs = qualifying.reduce((sum,b) => sum + b.breadth, 0) / ns;
    const ls = h * (10 / ns + 5);
    const sParam = ls / Math.sqrt(Math.max(hs * bs, 0.01));
    const msVal = parseFloat(tableLookupMs(sParam).toFixed(2));
    msArr[i] = msVal;
    S.shieldingDetails[i] = { ns, ls:parseFloat(ls.toFixed(1)), hs:parseFloat(hs.toFixed(1)), bs:parseFloat(bs.toFixed(1)), s:parseFloat(sParam.toFixed(2)), Ms:msVal };
  }
}

function getDesignBuildingHeightH(){
  const ridgeRise = (S.pitch||0) > 0 && S.roofType !== 'flat' ? Math.tan((S.pitch||0)*Math.PI/180) * (S.depth||10) / 2 : 0;
  return (S.height||6) + (S.parapet||0) + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);
}

//  AUTO-DETECT TERRAIN, SHIELDING & TOPO FROM MAP
// ═══════════════════════════════════════════════

/** Clear OSM terrain/shielding cache when Overpass fails so stale footprints are not mixed with fallback multipliers. */
function clearStaleOsmTerrainCache(){
  S.detectedBuildingsList = null;
  S.terrainRecalcCtx = null;
  S.detectedBuildingsPerSector = null;
  S.detectedNearBuildings = null;
  S.detectedSectorWater = null;
  S.detectedSectorOpen = null;
}

/**
 * Wipe prior TC/Ms/Mt terrain profile when the site pin moves so map overlays (e.g. TC Z1/Z2 rings)
 * are not redrawn at the new coordinates using stale multi-zone data from the old location.
 */
function resetSiteTerrainForNewPinLocation(){
  S.terrainZones = [];
  clearStaleOsmTerrainCache();
  S.manualShieldBuildings = [];
  if(!Array.isArray(S.TC_dir)) S.TC_dir = new Array(8).fill(null);
  if(!Array.isArray(S.Ms)) S.Ms = new Array(8).fill(1);
  if(!Array.isArray(S.Mzcat)) S.Mzcat = new Array(8).fill(null);
  if(!Array.isArray(S.shieldingDetails)) S.shieldingDetails = new Array(8).fill(null).map(()=>({ ns:0, ls:0, hs:0, bs:0, s:Infinity, Ms:1.0 }));
  const h = getDesignBuildingHeightH();
  for(let i=0;i<8;i++){
    S.TC_dir[i] = null;
    S.Ms[i] = 1;
    S.Mzcat[i] = null;
    S.shieldingDetails[i] = { ns:0, ls:0, hs:0, bs:0, s:Infinity, Ms:1.0 };
  }
  S.Mt_hill = [1,1,1,1,1,1,1,1];
  S.detectedElevations = null;
  S.detectedSiteElev = null;
  S.detectedSampleDistances = null;
  S.detectedProfiles = null;
  S.mhSub = null;
  S.mhDetailsSub = null;
  S.elevBearingsSub = null;
  S.detectionTimestamp = null;
}

/**
 * Apply TC/Ms/Md to S immediately after OSM parse + recomputeTerrainAndShieldingFromBuildings.
 * Mzcat is already set on S by recompute; fallback fills missing slots.
 */
function applyOsmTcMsMdToState(h, tcArr, msArr, opts){
  opts = opts || {};
  const applyTc = opts.applyTc !== false;
  const applyMs = opts.applyMs !== false;
  const mdVals = MD_TABLE[S.region]||[1,1,1,1,1,1,1,1];
  for(let i=0;i<8;i++){
    if(applyTc){
      S.TC_dir[i] = tcArr[i];
      if(!S.Mzcat[i] || S.Mzcat[i] === 0) S.Mzcat[i] = mzCat(h, tcArr[i]);
    } else {
      S.TC_dir[i] = 3;
      if(!S.Mzcat[i] || S.Mzcat[i] === 0) S.Mzcat[i] = mzCat(h, 3);
    }
    if(applyMs){
      S.Ms[i] = msArr[i];
    } else {
      S.Ms[i] = 1;
    }
    S.Md[i] = mdVals[i];
  }
  calc();
  checkLeeZone();
  calc();
  refreshDirectionalWindUI();
  if(S.overlayTC) drawTCZones();
  if(S.overlayMs) drawMsOverlay();
  if(opts.toast !== false) toast('✓ Shielding & terrain category updated (topography may still be sampling)');
}

/** Apply Mt_hill after elevation sampling; updates combined Mt via calc(). */
function applyElevationMtHillToState(mtArr, opts){
  opts = opts || {};
  for(let i=0;i<8;i++) S.Mt_hill[i] = mtArr[i];
  calc();
  checkLeeZone();
  calc();
  refreshDirectionalWindUI();
  if(S.overlayMt) drawMtOverlay(opts.drawOpts || {});
  if(opts.toast === true) toast('✓ Topographic multipliers (Mt) updated');
}

/**
 * Master function — single Overpass query for TC+Ms, elevation API for Mt.
 * Runs both in parallel; TC/Ms apply as soon as OSM returns; Mt when elevations return.
 */
async function autoDetectAllMultipliers(opts){
  opts = opts || {};
  deactivateMapOverlaysIfFreeTierExceeded();
  if(autoDetectInFlight){
    autoDetectQueued = true;
    autoDetectQueuedOpts = opts;
    toast('⚠ Detection already running — will re-run when current finishes');
    return;
  }
  const ridgeRise = (S.pitch||0) > 0 && S.roofType !== 'flat' ? Math.tan((S.pitch||0)*Math.PI/180) * (S.depth||10) / 2 : 0;
  const h = (S.height||6) + (S.parapet||0) + ridgeRise * (S.roofType==='monoslope' ? 1 : 0.5);
  const lat = S.lat, lng = S.lng;
  if(!lat || !lng){ toast('⚠ Set a location first'); return; }

  const userFree = !!(currentUser && !isPaidPlan());
  const fromPin = opts.fromPinMove === true;
  const allowTc = !userFree || !fromPin || mapOverlayFreeTierAllowsMapAnalysis('tc');
  const allowMs = !userFree || !fromPin || mapOverlayFreeTierAllowsMapAnalysis('ms');
  const allowMt = !userFree || !fromPin || mapOverlayFreeTierAllowsMapAnalysis('mt');

  if(fromPin && userFree && !allowTc && !allowMs && !allowMt){
    applyOsmTcMsMdToState(h, [3,3,3,3,3,3,3,3], [1,1,1,1,1,1,1,1], { toast: false });
    applyElevationMtHillToState([1,1,1,1,1,1,1,1], { toast: false });
    setTerrainDataStatus('fallback', 'free tier — map terrain recalc disabled when moving site');
    toast('Terrain recalculation remains available in this local build.');
    refreshDirectionalWindUI();
    return;
  }

  const runGen = siteDetectGeneration;
  const ac = new AbortController();
  activeDetectAbort = ac;

  autoDetectInFlight = true;
  detectPendingOsm = true;
  detectPendingElev = true;
  // Clear prior elevation outputs so Mt loading UI matches this run (avoids spinner after Mt is already computed).
  S.detectedProfiles = null;
  S.mhSub = null;
  S.mhDetailsSub = null;
  S.detectedElevations = null;
  S.detectedSiteElev = null;
  S.detectedSampleDistances = null;
  S.elevBearingsSub = null;
  refreshDirectionalWindUI();
  try{
  toast(fromPin ? '⏳ Sampling terrain…' : '🔍 Detecting terrain, shielding & topography…');
  setTerrainDataStatus('loading');

  // Default arrays
  const tcArr = [3,3,3,3,3,3,3,3];
  const msArr = [1,1,1,1,1,1,1,1];
  const mtArr = [1,1,1,1,1,1,1,1];

  const fastDetect = false;
  const sampleDistancesFullFallback = [
    30, 60, 90, 120, 150, 180, 210, 240, 270, 300,
    340, 380, 420, 470, 520,
    580, 650, 730, 820, 920,
    1050, 1200, 1400, 1650, 2000, 2500, 3200, 4200, 5000,
  ];
  const sampleDistancesFastFallback = [40, 80, 120, 180, 240, 300, 400, 500, 650, 850, 1100, 1500, 2200, 4000];
  const sampleDistancesFull =
    typeof globalThis !== 'undefined' && globalThis.CW_SAMPLE_DISTANCES_FULL
      ? globalThis.CW_SAMPLE_DISTANCES_FULL
      : sampleDistancesFullFallback;
  const sampleDistancesFast =
    typeof globalThis !== 'undefined' && globalThis.CW_SAMPLE_DISTANCES_FAST
      ? globalThis.CW_SAMPLE_DISTANCES_FAST
      : sampleDistancesFastFallback;
  const sampleDistances = fastDetect ? sampleDistancesFast : sampleDistancesFull;
  const nDistMt = sampleDistances.length;
  const nSubMt = fastDetect ? 16 : 32;
  const elevBatchOpts = fastDetect
    ? {
        openMeteoBatchGapMs: 40,
        euBatchGapMs: 40,
        elevIdb: { lat, lng, fast: true },
        elevationRefineMeta: { nSub: nSubMt, nDist: nDistMt },
      }
    : {
        openMeteoBatchGapMs: 280,
        euBatchGapMs: 120,
        elevIdb: { lat, lng, fast: false },
        elevationRefineMeta: { nSub: nSubMt, nDist: nDistMt },
      };
  const tDetect0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : null;

  // ── 1) Single Overpass query for buildings + land use (TC + Ms) ──
  const osmPromise = (async ()=>{
    try{
    if(fromPin && userFree && !allowTc && !allowMs){
      if(runGen !== siteDetectGeneration) return;
      detectPendingOsm = false;
      applyOsmTcMsMdToState(h, [3,3,3,3,3,3,3,3], [1,1,1,1,1,1,1,1], { toast: false });
      refreshDirectionalWindUI();
      return;
    }
    const radius = Math.max(565, Math.ceil(40 * h));  // Match averaging distance + margin
    const query = `[out:json][timeout:55];(node["building"](around:${radius},${lat},${lng});way["building"](around:${radius},${lat},${lng});relation["building"](around:${radius},${lat},${lng});way["landuse"](around:${radius},${lat},${lng});relation["landuse"](around:${radius},${lat},${lng});way["natural"~"water|coastline|wood"](around:${radius},${lat},${lng});way["landuse"="forest"](around:${radius},${lat},${lng});way["amenity"="parking"](around:${radius},${lat},${lng});way["landuse"="garages"](around:${radius},${lat},${lng}););out body geom qt 3000;`;

    const cacheKey = cwOverpassCacheKey(lat, lng, radius, fastDetect);
    let elements = null;
    let osmFetchMeta = null;

    const cachedEls = readCwOsmCache(cacheKey);
    if(cachedEls && runGen === siteDetectGeneration){
      elements = cachedEls;
      setTerrainDataStatus('cached');
      console.log('Overpass OK — session cache —', elements.length, 'elements');
    } else {
      osmFetchMeta = await fetchOverpassElementsForSite({ query, runGen, ac, lat, lng });
      if(osmFetchMeta.cancelled){
        if(runGen === siteDetectGeneration) setTerrainDataStatus('unknown');
        detectPendingOsm = false;
        refreshDirectionalWindUI();
        return;
      }
      if(osmFetchMeta.elements !== null && osmFetchMeta.elements !== undefined){
        elements = osmFetchMeta.elements;
        writeCwOsmCache(cacheKey, elements);
        setTerrainDataStatus('live', formatTerrainLiveDetail(osmFetchMeta.viaProxy, osmFetchMeta.linz));
        console.log('Overpass OK — network —', elements.length, 'elements');
      }
    }

    if(elements === null){
      if(runGen !== siteDetectGeneration) return;
      if(ac.signal.aborted){
        if(runGen === siteDetectGeneration) setTerrainDataStatus('unknown');
        detectPendingOsm = false;
        refreshDirectionalWindUI();
        return;
      }
      const staleEls = readCwOsmCacheStaleFallback(cacheKey);
      if(staleEls && staleEls.length > 0){
        elements = staleEls;
        setTerrainDataStatus('cached', 'stale — live fetch failed; using last OSM snapshot for this site');
        toast('⚠ Live map fetch failed — using cached OpenStreetMap for this location.');
        console.warn('Overpass failed — using stale session cache', osmFetchMeta && osmFetchMeta.reason);
      }
    }

    if(elements === null){
      if(runGen !== siteDetectGeneration) return;
      const detail = osmFetchMeta && osmFetchMeta.reason ? osmFetchMeta.reason : 'unavailable';
      console.warn('All Overpass paths failed or timed out', detail);
      clearStaleOsmTerrainCache();
      let msg = '⚠ Map data unavailable — TC/Ms set to conservative defaults.';
      if(detail === 'unauthorized'){
        msg = '⚠ Local map data route did not authorize or respond — TC/Ms set to conservative defaults.';
      } else if(detail === 'timeout'){
        msg = '⚠ Map data did not finish in time — TC/Ms set to conservative defaults. Run Detect again or use a backend Overpass proxy for production.';
      } else if(String(detail).indexOf('http_429') === 0 || String(detail).indexOf('http_503') === 0){
        msg = '⚠ Map service busy — TC/Ms set to conservative defaults.';
      } else if(String(detail).indexOf('http_') === 0){
        msg = '⚠ Map service error (' + detail.replace('http_', '') + ') — TC/Ms set to conservative defaults.';
      }
      toast(msg);
      setTerrainDataStatus('fallback', detail);
      detectPendingOsm = false;
      applyOsmTcMsMdToState(h, tcArr, msArr, { toast: false, applyTc: allowTc, applyMs: allowMs });
      return;
    }

    await new Promise(r=>setTimeout(r, 0));
    if(runGen !== siteDetectGeneration) return;

    // ── Parse buildings with full geometry for Ms analysis (Cl 4.3) ──
    // Also parse landuse polygons for terrain category estimation
    const buildingsList = [];
    const sectorBuildings = [0,0,0,0,0,0,0,0];
    const sectorNearBuildings = [0,0,0,0,0,0,0,0];
    const sectorWater  = [false,false,false,false,false,false,false,false];
    const sectorOpen   = [false,false,false,false,false,false,false,false];
    const shieldDist = Math.min(20*h, 500);
    const mPerDegLat = 111320;
    const mPerDegLng = 111320 * Math.cos(lat*Math.PI/180);

    // Per-sector landuse coverage tracking using point-in-polygon tests
    // Finer bands near site (0–20, 20–50) so carparks/open areas at site get correct TC
    const luBandEdges = [0, 20, 50, 100, 150, 200, 250, 300, 400, 500, 600];
    const nLuBands = luBandEdges.length - 1;
    // Coverage arrays: [sector][band] = { urban: 0, forest: 0, open: 0, water: 0, parking: 0 }
    // Fast detect: 1 sample per band; full: 5 (center + radial offsets)
    const bandCoverage = Array.from({length:8}, ()=>
      Array.from({length:nLuBands}, ()=>({urban:0, forest:0, open:0, water:0, parking:0}))
    );

    // Ray-casting point-in-polygon test (lat/lng coords)
    function pointInPoly(testLat, testLng, polyCoords){
      let inside = false;
      for(let i=0, j=polyCoords.length-1; i<polyCoords.length; j=i++){
        const yi = polyCoords[i].lat, xi = polyCoords[i].lon;
        const yj = polyCoords[j].lat, xj = polyCoords[j].lon;
        if(((yi > testLat) !== (yj > testLat)) &&
           (testLng < (xj-xi)*(testLat-yi)/(yj-yi)+xi)){
          inside = !inside;
        }
      }
      return inside;
    }

    // Generate test points: multi-point grid per sector/band for robust coverage
    const bearings8 = [0,45,90,135,180,225,270,315];
    const testPoints = []; // [{lat,lng,sector,band}]
    const subOffsets = fastDetect ? [0.5] : [0, 0.25, 0.5, 0.75, 1]; // radial position within band
    for(let si=0;si<8;si++){
      const bRad = bearings8[si]*Math.PI/180;
      for(let bi=0;bi<nLuBands;bi++){
        const innerR = luBandEdges[bi], outerR = luBandEdges[bi+1];
        for(const frac of subOffsets){
          const r = innerR + frac * (outerR - innerR);
          // Add small angular offset (±8°) to sample across sector width
          const angOff = (frac - 0.5) * 16 * Math.PI/180;
          const rLat = lat + r*Math.cos(bRad + angOff)/mPerDegLat;
          const rLng = lng + r*Math.sin(bRad + angOff)/mPerDegLng;
          testPoints.push({ lat: rLat, lng: rLng, sector: si, band: bi });
        }
      }
    }

    // Flatten relation members into top-level elements (relations may contain ways with geometry)
    const flatElements = [];
    elements.forEach(el=>{
      if(el.type === 'relation' && el.members){
        el.members.forEach(m=>{
          if(m.geometry && m.geometry.length > 0){
            flatElements.push({ type:'way', tags: el.tags, geometry: m.geometry });
          }
        });
      }
      if(el.geometry || el.center || el.lat) flatElements.push(el);
    });

    // Collect landuse polygons for point-in-polygon testing
    const landusePolys = [];
    flatElements.forEach(el=>{
      if(!el.tags || !el.geometry || el.geometry.length < 3) return;
      const isUrban = el.tags.landuse==='residential'||el.tags.landuse==='commercial'||
                      el.tags.landuse==='industrial'||el.tags.landuse==='retail'||
                      el.tags.landuse==='construction';
      const isForest = el.tags.natural==='wood'||el.tags.landuse==='forest';
      const isOpenLand = el.tags.landuse==='farmland'||el.tags.landuse==='meadow'||
                         el.tags.landuse==='grass'||el.tags.landuse==='recreation_ground'||
                         el.tags.landuse==='orchard'||el.tags.landuse==='vineyard';
      const isWater = el.tags.natural==='water'||el.tags.natural==='coastline'||
                      el.tags.water!==undefined;
      // Carparks, parking lots, garages — open pavement, minimal wind obstruction (TC2)
      const isParking = el.tags.amenity==='parking'||el.tags.landuse==='garages';
      if(isUrban || isForest || isOpenLand || isWater || isParking){
        let minLat = 90, maxLat = -90, minLng = 180, maxLng = -180;
        el.geometry.forEach(n=>{
          if(n.lat < minLat) minLat = n.lat;
          if(n.lat > maxLat) maxLat = n.lat;
          if(n.lon < minLng) minLng = n.lon;
          if(n.lon > maxLng) maxLng = n.lon;
        });
        landusePolys.push({ geom: el.geometry, minLat, maxLat, minLng, maxLng, isUrban, isForest, isOpenLand, isWater, isParking });
      }
    });

    const LU_BUCKET_DEG = 0.01;
    const luPolyBuckets = landusePolys.length ? buildLanduseSpatialBuckets(landusePolys, LU_BUCKET_DEG) : null;

    // For each test point, check which landuse polygons contain it (count-based for multi-point)
    testPoints.forEach(tp=>{
      const cv = bandCoverage[tp.sector][tp.band];
      const polysToScan = luPolyBuckets
        ? landusePolysNearPoint(tp, luPolyBuckets, LU_BUCKET_DEG, landusePolys)
        : landusePolys;
      for(const poly of polysToScan){
        if(pointInPoly(tp.lat, tp.lng, poly.geom)){
          if(poly.isParking)  cv.parking++;
          if(poly.isUrban)    cv.urban++;
          if(poly.isForest)   cv.forest++;
          if(poly.isOpenLand) cv.open++;
          if(poly.isWater)    cv.water++;
        }
      }
    });

    console.log('Landuse PIP coverage:', bandCoverage.map((s,i)=>
      s.map((b,j)=>(b.urban||b.forest||b.open||b.water||b.parking)?
        `S${i}B${j}:P${b.parking}U${b.urban}F${b.forest}O${b.open}W${b.water}`:null
      ).filter(Boolean)
    ).filter(a=>a.length));

    // Process all elements for buildings
    flatElements.forEach(el=>{
      if(!el.tags) return;

      const isOpenLand = el.tags.landuse==='farmland'||el.tags.landuse==='meadow'||
                         el.tags.landuse==='grass'||el.tags.landuse==='recreation_ground'||
                         el.tags.landuse==='orchard'||el.tags.landuse==='vineyard';
      const isWater = el.tags.natural==='water'||el.tags.natural==='coastline'||
                      el.tags.water!==undefined;
      let eLat, eLng;
      if(el.center){ eLat=el.center.lat; eLng=el.center.lon; }
      else if(el.geometry && el.geometry.length > 0){
        let sLat2=0, sLng2=0;
        el.geometry.forEach(n=>{ sLat2+=n.lat; sLng2+=n.lon; });
        eLat = sLat2/el.geometry.length;
        eLng = sLng2/el.geometry.length;
      }
      else if(el.lat && el.lon){ eLat=el.lat; eLng=el.lon; }
      else return;

      const dLat = eLat - lat;
      const dLng = eLng - lng;
      let bearing = Math.atan2(dLng, dLat)*180/Math.PI;
      bearing = ((bearing%360)+360)%360;
      const idx = Math.round(bearing/45) % 8;
      const distM = Math.sqrt((dLat*mPerDegLat)**2 + (dLng*mPerDegLng)**2);

      // Sector-level flags
      if(isWater) sectorWater[idx]=true;
      if(isOpenLand) sectorOpen[idx]=true;

      // Buildings — full per-building analysis
      if(el.tags.building){
        sectorBuildings[idx]++;
        if(distM <= shieldDist) sectorNearBuildings[idx]++;

        // Footprint polygon from way geometry
        const footprint = (el.geometry && el.geometry.length >= 3)
          ? el.geometry.map(n=>[n.lat, n.lon]) : null;

        // Building height from OSM tags — improved estimation
        // Priority: explicit height tag > building:levels > building type default
        let bHeight = 0;
        if(el.tags.height){
          bHeight = parseFloat(el.tags.height) || 0;
        }
        if(!bHeight && el.tags['building:levels']){
          const levels = parseFloat(el.tags['building:levels']) || 1;
          // NZ residential ~2.7m/level, commercial ~3.5m, add ~1m for roof
          const isCommercial = el.tags.building==='commercial'||el.tags.building==='retail'||
            el.tags.building==='industrial'||el.tags.building==='warehouse';
          bHeight = levels * (isCommercial ? 3.5 : 2.7) + 1.0;
        }
        if(!bHeight){
          // Default by building type — NZ residential typically 5-8m to ridge
          // Most NZ OSM buildings are tagged building=yes so this default is critical
          const bt = el.tags.building;
          if(bt==='house'||bt==='residential'||bt==='detached'||bt==='semidetached_house'||bt==='terrace'){
            bHeight = 7; // typical NZ single-story house to ridge
          } else if(bt==='apartments'||bt==='dormitory'){
            bHeight = 12;
          } else if(bt==='commercial'||bt==='retail'||bt==='office'){
            bHeight = 11;
          } else if(bt==='industrial'||bt==='warehouse'){
            bHeight = 8;
          } else if(bt==='garage'||bt==='shed'||bt==='carport'){
            bHeight = 3.5;
          } else {
            // building=yes — mixed urban (OSM rarely has height); bias toward mid-rise for TC
            bHeight = 9;
          }
        }

        // Footprint area via Shoelace formula (m²)
        let area = 0;
        if(footprint && footprint.length >= 3){
          for(let j=0; j<footprint.length; j++){
            const k=(j+1)%footprint.length;
            const xj=(footprint[j][1]-lng)*mPerDegLng, yj=(footprint[j][0]-lat)*mPerDegLat;
            const xk=(footprint[k][1]-lng)*mPerDegLng, yk=(footprint[k][0]-lat)*mPerDegLat;
            area += xj*yk - xk*yj;
          }
          area = Math.abs(area)/2;
        }

        // Breadth perpendicular to bearing from site (shielding width)
        let breadth = Math.sqrt(Math.max(area, 1));
        if(footprint && footprint.length >= 3){
          const bRad = bearing*Math.PI/180;
          const perpAngle = bRad + Math.PI/2;
          let minP=Infinity, maxP=-Infinity;
          footprint.forEach(p=>{
            const dx=(p[1]-eLng)*mPerDegLng, dy=(p[0]-eLat)*mPerDegLat;
            const proj = dx*Math.cos(perpAngle) + dy*Math.sin(perpAngle);
            if(proj<minP) minP=proj;
            if(proj>maxP) maxP=proj;
          });
          if(isFinite(minP) && isFinite(maxP) && maxP-minP > 0.5) breadth = maxP - minP;
        }
        // Node buildings (no footprint) default to 8m breadth (typical NZ house width)
        if(!footprint || footprint.length < 3) breadth = 8;

        // Height score relative to designed building
        const heightScore = Math.min(1.0, bHeight / h);
        const osmKey = (el.type||'way')+'/'+el.id;

        buildingsList.push({
          id: el.id,
          osmKey,
          footprint,
          centroid: [eLat, eLng],
          distance: distM, bearing, sectorIdx: idx,
          height: bHeight,
          heightInferred: bHeight,
          area, breadth: Math.max(breadth, 3),
          heightScore,
          elevation: null, slope: 0,
          tags: el.tags
        });
      }
    });

    const terrainCtx = {
      lat, lng, h, mPerDegLat, mPerDegLng,
      bandCoverage, luBandEdges,
      sectorBuildings: sectorBuildings.slice(),
      sectorWater: sectorWater.slice(),
      sectorOpen: sectorOpen.slice(),
      shieldDist
    };
    S.terrainRecalcCtx = terrainCtx;
    applyHeightOverridesToBuildings(buildingsList, h);
    recomputeTerrainAndShieldingFromBuildings(buildingsList, terrainCtx, tcArr, msArr);

    // Store raw detection data
    S.detectedBuildingsPerSector = [...sectorBuildings];
    S.detectedNearBuildings = [...sectorNearBuildings];
    S.detectedSectorWater = [...sectorWater];
    S.detectedSectorOpen = [...sectorOpen];
    S.detectedBuildingsList = buildingsList;
    console.log('OSM detected —', buildingsList.length, 'buildings, per sector:', sectorBuildings, 'near:', sectorNearBuildings);
    if(runGen !== siteDetectGeneration) return;
    detectPendingOsm = false;
    applyOsmTcMsMdToState(h, tcArr, msArr, { applyTc: allowTc, applyMs: allowMs });
    } catch(e){
      console.error('OSM pipeline error:', e);
      detectPendingOsm = false;
      refreshDirectionalWindUI();
      if(runGen !== siteDetectGeneration) return;
      if(e && (e.name === 'AbortError' || e.message === 'The user aborted a request.')){
        setTerrainDataStatus('unknown');
        return;
      }
      setTerrainDataStatus('fallback', 'parse_error');
      toast('⚠ Map data processing failed — TC/Ms set to defaults.');
      applyOsmTcMsMdToState(h, tcArr, msArr, { toast: false, applyTc: true, applyMs: true });
    }
  })();

  // ── 2) Elevation query for Mt — AS/NZS 1170.2 Section 4.4.2 ──
  // Dense sampling at SRTM resolution (~30 m) near site to properly identify
  // terrain features (hills, ridges, escarpments, gullies) per Figures 4.3–4.5,
  // then coarser sampling for distant features.
  const elevPromise = (async ()=>{
    try{
    if(fromPin && userFree && !allowMt){
      if(runGen !== siteDetectGeneration) return;
      detectPendingElev = false;
      applyElevationMtHillToState([1,1,1,1,1,1,1,1], { toast: false });
      refreshDirectionalWindUI();
      return;
    }
    const nDist = nDistMt;
    const mPerDegLat = 111320;
    const mPerDegLng = 111320 * Math.cos(lat*Math.PI/180);
    const nSub = nSubMt;
    const bearingsSub = [];
    for(let i=0;i<nSub;i++) bearingsSub.push(i*(360/nSub));

    function ptAt(b, dist){
      const bRad = b*Math.PI/180;
      return [lat + dist*Math.cos(bRad)/mPerDegLat, lng + dist*Math.sin(bRad)/mPerDegLng];
    }

    const lats = [lat], lngs = [lng];
    for(const b of bearingsSub){
      for(const dist of sampleDistances){
        const [la,lo] = ptAt(b, dist);
        lats.push(la); lngs.push(lo);
      }
    }

    try {
      toast('⏳ Sampling terrain ('+lats.length+' pts'+(fastDetect ? ', fast' : '')+')…');
      const elevations = await fetchElevBatchGlobal(lats, lngs, elevBatchOpts);
      if(runGen !== siteDetectGeneration) return;
      const siteElev = elevations[0];
      S.elevation = siteElev.toFixed(2)+' m';
      const elevEl = document.getElementById('inp-elevation');
      if(elevEl) elevEl.value = S.elevation;

      // Build per-direction profiles
      const profiles = [];
      for(let si=0; si<nSub; si++){
        const baseIdx = 1 + si * nDist;
        const prof = [{dist:0, elev:siteElev}];
        for(let d=0; d<nDist; d++){
          prof.push({dist:sampleDistances[d], elev:elevations[baseIdx+d]});
        }
        profiles.push(prof);
      }

      const z = h; // reference height above ground

      // ── Compute Mh for each of 32 sub-directions per Cl 4.4.2 ──
      const mhSub = [];
      const mhDetailsSub = [];

      for(let si=0; si<nSub; si++){
        const upProfile = profiles[si];
        const oppSi = (si + nSub/2) % nSub;
        const downProfile = profiles[oppSi].slice(1); // exclude site duplicate

        const result = computeMhFromProfiles(upProfile, downProfile, siteElev, z);
        mhSub.push(result.Mh);
        mhDetailsSub.push(result);
      }

      // ── Map sub-direction Mh → 8 sector Mh (max over the four sub-rays in each 45° sector) ──
      const subPerSector = nSub / 8;
      for(let i=0;i<8;i++){
        const c = i * subPerSector;
        let sectorMh = 0;
        for(let off = 0; off < subPerSector; off++){
          const idx = c + off;
          if(mhSub[idx] > sectorMh) sectorMh = mhSub[idx];
        }
        mtArr[i] = parseFloat(Math.max(sectorMh, 1.0).toFixed(4));
      }

      // Store data for overlay & PDF
      S.detectedElevations = [...elevations];
      S.detectedSiteElev = siteElev;
      S.detectedSampleDistances = sampleDistances;
      S.nSubDirs = nSub;
      S.mhSub = mhSub;
      S.mhDetailsSub = mhDetailsSub;
      S.elevBearingsSub = bearingsSub;
      S.detectedProfiles = profiles;
      const totalPts = nSub * nDist + 1;
      console.log('Elevation detected — site:', siteElev.toFixed(1)+'m',
        totalPts, 'points (' + nDist + ' per dir ×', nSub, 'dirs)',
        'Mh('+nSub+'):', mhSub.map(v=>v.toFixed(4)),
        'Mt_hill(8):', mtArr.map(v=>v.toFixed(4)));
      detectPendingElev = false;
      applyElevationMtHillToState(mtArr, { toast: false });
    } catch(err){
      if(runGen !== siteDetectGeneration) return;
      console.warn('Elevation API failed:', err);
      toast('⚠ Elevation query failed — Mt set to 1.0');
      for(let i=0;i<8;i++) mtArr[i] = 1;
      detectPendingElev = false;
      applyElevationMtHillToState(mtArr, { toast: false });
    }
    } catch(e){
      console.error('Elevation pipeline error:', e);
      if(runGen !== siteDetectGeneration) return;
      detectPendingElev = false;
      refreshDirectionalWindUI();
      throw e;
    }
  })();

  // ── Wait for both to finish (TC/Ms and Mt already applied inside each promise) ──
  await Promise.allSettled([osmPromise, elevPromise]);

  if(runGen !== siteDetectGeneration) return;

  // ── Fetch building centroid elevations (full detect only; skipped in fast to save batches) ──
  const MAX_BUILDING_ELEV_POINTS = fastDetect ? 0 : 250;
  const buildingsList = S.detectedBuildingsList || [];
  if(buildingsList.length > 0 && MAX_BUILDING_ELEV_POINTS > 0){
    try {
      const siteElevMs = S.detectedSiteElev || 0;
      let listForElev = buildingsList;
      if(buildingsList.length > MAX_BUILDING_ELEV_POINTS){
        listForElev = buildingsList.slice().sort((a,b)=>(a.distance||0)-(b.distance||0)).slice(0, MAX_BUILDING_ELEV_POINTS);
        console.warn('Building elevation query capped to nearest', MAX_BUILDING_ELEV_POINTS, 'of', buildingsList.length, 'structures');
      }
      const bLats = listForElev.map(b=>b.centroid[0]);
      const bLngs = listForElev.map(b=>b.centroid[1]);
      const bElevs = await fetchElevBatchGlobal(bLats, bLngs, {
        openMeteoBatchGapMs: elevBatchOpts.openMeteoBatchGapMs,
        euBatchGapMs: elevBatchOpts.euBatchGapMs,
      });
      if(runGen !== siteDetectGeneration) return;
      bElevs.forEach((elev,j)=>{ listForElev[j].elevation = elev; });
      listForElev.forEach(b=>{
        if(b.elevation !== null && b.distance > 0){
          b.slope = (b.elevation - siteElevMs) / b.distance;
        }
      });
      console.log('Building elevations fetched for', listForElev.length, 'buildings');
    } catch(e){ console.warn('Building elevation fetch:', e.message); }
  }

  if(runGen !== siteDetectGeneration) return;

  console.log('Detection complete — TC:', S.TC_dir, 'Ms:', S.Ms, 'Mt_hill:', S.Mt_hill);
  if(tDetect0 != null && typeof performance !== 'undefined' && performance.now){
    console.log('autoDetectAllMultipliers total:', Math.round(performance.now() - tDetect0), 'ms');
  }
  S.detectionTimestamp = new Date().toLocaleString();
  refreshDirectionalWindUI();
  toast('✓ Multipliers auto-detected for all 8 directions');
 } catch(err){
  if(typeof runGen !== 'undefined' && runGen !== siteDetectGeneration) return;
  console.error('autoDetectAllMultipliers error:', err);
  if(err && err.name === 'AbortError'){
    if(runGen === siteDetectGeneration) setTerrainDataStatus('unknown');
    return;
  }
  if(runGen === siteDetectGeneration) setTerrainDataStatus('fallback', err && err.message ? String(err.message).slice(0, 80) : 'error');
  toast('⚠ Auto-detection failed: '+(err && err.message ? err.message : String(err)));
 } finally {
  if(activeDetectAbort === ac) activeDetectAbort = null;
  if(typeof runGen !== 'undefined' && runGen === siteDetectGeneration){
    autoDetectInFlight = false;
    detectPendingOsm = false;
    detectPendingElev = false;
    refreshDirectionalWindUI();
    if(autoDetectQueued){
      autoDetectQueued = false;
      const nextOpts = autoDetectQueuedOpts || {};
      autoDetectQueuedOpts = null;
      setTimeout(()=>{ autoDetectAllMultipliers(nextOpts); }, 0);
    }
  } else {
    refreshDirectionalWindUI();
  }
 }
}
function clearMultipliers(){
  syncSiteCoordsFromMapPin();
  const mdVals = MD_TABLE[S.region]||[1,1,1,1,1,1,1,1];
  for(let i=0;i<8;i++){
    S.Md[i]=mdVals[i]; S.Mzcat[i]=null; S.Ms[i]=1; S.Mt[i]=1; S.Mt_hill[i]=1; S.TC_dir[i]=null;
    S.Mlee[i]=1;
  }
  S.leeZone=null; S.leeOverride=false;
  S.mhSub=null; S.mhDetailsSub=null; S.elevBearingsSub=null; S.nSubDirs=32;
  S.detectedProfiles=null;
  S.detectedElevations=null; S.detectedSiteElev=null; S.detectedSampleDistances=null;
  S.detectedBuildingsList=null;
  S.terrainRecalcCtx=null;
  calc();
  refreshDirectionalWindUI();
  updateLeeNotification();
  if(S.overlayTC) drawTCZones();
  if(S.overlayMs) drawMsOverlay();
  if(S.overlayMt) drawMtOverlay();
  setTerrainDataStatus('unknown');
  toast('Multipliers cleared');
}

// ═══════════════════════════════════════════════
//   LEE ZONE MULTIPLIER  (Clause 4.4.3, Table 4.4)
// ═══════════════════════════════════════════════
// dirs = wind direction FROM which the wind blows that triggers Mlee
// shadow = shadow zone extent (km from crest)
// outer = outer zone width (km beyond shadow zone, linear interp Mlee → 1.0)
const LEE_ZONES = [
  // ── North Island ──
  {id:1, name:'Kaimai',
    dirs:['E','SE'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-37.38,175.85],[-37.45,175.83],[-37.52,175.81],[-37.60,175.79],[-37.68,175.76],[-37.75,175.74]]},
  {id:2, name:'Taranaki',
    dirs:['N','NE','E','SE','S','SW','W','NW'], Mlee:1.35, shadow:12, outer:18,
    crests:[[-39.2962,174.0634]]},
  {id:3, name:'Ruapehu',
    dirs:['NW','SE'], Mlee:1.35, shadow:12, outer:18,
    crests:[[-39.28,175.57]]},
  {id:4, name:'Tararua',
    dirs:['SE'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-40.72,175.42],[-40.82,175.38],[-40.90,175.35],[-40.98,175.30],[-41.05,175.26]]},
  {id:5, name:'Tararua and Orongorongo',
    dirs:['NW'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-40.72,175.42],[-40.82,175.38],[-40.90,175.35],[-40.98,175.30],[-41.05,175.26],[-41.15,175.08],[-41.22,175.00],[-41.30,174.96]]},
  // Zone 6 (Coastal Wairarapa) — NW direction, no separate values in Table 4.4; covered by zone 5

  // ── South Island ──
  {id:7, name:'West Coast North',
    dirs:['E','SE'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-42.00,171.80],[-42.10,171.75],[-42.20,171.70],[-42.30,171.65],[-42.40,171.60]]},
  {id:8, name:'West Coast Alps',
    dirs:['SE'], Mlee:1.35, shadow:12, outer:18,
    crests:[[-43.00,171.10],[-43.20,171.00],[-43.40,170.90],[-43.60,170.80],[-43.80,170.60]]},
  {id:9, name:'Awatere',
    dirs:['NW'], Mlee:1.35, shadow:12, outer:15,
    crests:[[-41.80,173.70],[-41.90,173.60],[-42.00,173.50],[-42.10,173.45]]},
  {id:10,name:'Inland Kaikoura',
    dirs:['NW'], Mlee:1.35, shadow:12, outer:20,
    crests:[[-42.15,173.40],[-42.25,173.35],[-42.35,173.30],[-42.45,173.25]]},
  {id:11,name:'Southern Alps',
    dirs:['NW'], Mlee:1.35, shadow:12, outer:18,
    crests:[[-42.50,171.50],[-42.80,171.35],[-43.10,171.10],[-43.40,170.90],[-43.70,170.70],[-44.00,170.40],[-44.30,169.80],[-44.60,169.40]]},
  {id:12,name:'Hunter',
    dirs:['SW'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-45.50,167.50],[-45.60,167.40],[-45.70,167.30]]},
  {id:13,name:'Hakataramea',
    dirs:['NW'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-44.40,170.50],[-44.50,170.45],[-44.60,170.40]]},
  {id:14,name:'St Mary\'s',
    dirs:['SW'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-44.65,170.15],[-44.75,170.10]]},
  {id:15,name:'Pisa',
    dirs:['NW'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-44.90,169.20],[-44.95,169.15],[-45.00,169.10]]},
  {id:16,name:'Dunstan',
    dirs:['NW'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-45.00,169.45],[-45.10,169.40],[-45.20,169.35]]},
  {id:17,name:'Rock and Pillar',
    dirs:['NW'], Mlee:1.20, shadow:8, outer:12,
    crests:[[-45.40,170.05],[-45.50,170.10],[-45.60,170.15]]}
];

const DIR_BEARING = {N:0,NE:45,E:90,SE:135,S:180,SW:225,W:270,NW:315};

function haversineKm(lat1,lon1,lat2,lon2){
  const R=6371, dLat=(lat2-lat1)*Math.PI/180, dLon=(lon2-lon1)*Math.PI/180;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}

function bearingDeg(lat1,lon1,lat2,lon2){
  const dLon=(lon2-lon1)*Math.PI/180;
  const y=Math.sin(dLon)*Math.cos(lat2*Math.PI/180);
  const x=Math.cos(lat1*Math.PI/180)*Math.sin(lat2*Math.PI/180)-Math.sin(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.cos(dLon);
  return (Math.atan2(y,x)*180/Math.PI+360)%360;
}

function distFromCrestDownwind(lat,lng,crests,windBearing){
  let minDist = Infinity;
  for(const [cLat,cLng] of crests){
    const brg = bearingDeg(cLat,cLng,lat,lng);
    let diff = Math.abs(brg - windBearing);
    if(diff>180) diff=360-diff;
    if(diff<90){
      const d = haversineKm(cLat,cLng,lat,lng);
      if(d<minDist) minDist=d;
    }
  }
  return minDist;
}

function checkLeeZone(){
  if(S.leeOverride) return;
  const lat=S.lat, lng=S.lng;
  if(!lat||!lng) return;

  for(let i=0;i<8;i++) S.Mlee[i]=1;
  S.leeZone=null;

  for(const zone of LEE_ZONES){
    for(const dirName of zone.dirs){
      const dirIdx = S.dirs.indexOf(dirName);
      if(dirIdx<0) continue;
      const windBearing = dirIdx*45;
      const downwindBearing = (windBearing+180)%360;
      const d = distFromCrestDownwind(lat,lng,zone.crests,downwindBearing);
      if(d===Infinity) continue;

      console.log('Lee check:', zone.name, dirName, 'dist='+d.toFixed(1)+'km',
        'shadow='+zone.shadow, 'outer='+zone.outer, 'total='+(zone.shadow+zone.outer));

      let mlee = 1;
      if(d <= zone.shadow){
        mlee = zone.Mlee;
      } else if(d <= zone.shadow + zone.outer){
        const frac = (d - zone.shadow) / zone.outer;
        mlee = zone.Mlee + (1.0 - zone.Mlee) * frac;
      }

      if(mlee > S.Mlee[dirIdx]){
        S.Mlee[dirIdx] = parseFloat(mlee.toFixed(3));
        S.leeZone = zone;
      }
    }
  }
  console.log('Lee result:', S.Mlee, S.leeZone?.name||'none');
  updateLeeNotification();
}

function updateLeeNotification(){
  const el = document.getElementById('lee-notification');
  if(!el) return;
  if(S.leeZone && !S.leeOverride && S.Mlee.some(m=>m>1)){
    const zone = S.leeZone;
    const maxM = Math.max(...S.Mlee).toFixed(2);
    el.querySelector('.lee-text').textContent =
      'Lee Zone Detected: '+zone.name+' (Mlee up to '+maxM+') \u2014 Clause 4.4.3';
    el.style.display = 'flex';
  } else {
    el.style.display = 'none';
  }
}

function dismissLeeZone(){
  S.leeOverride = true;
  for(let i=0;i<8;i++) S.Mlee[i]=1;
  S.leeZone=null;
  calc();
  recalcPressures();
  updateLeeNotification();
  toast('Lee zone multiplier dismissed');
}

function enableLeeZone(){
  S.leeOverride = false;
  checkLeeZone();
  calc();
  recalcPressures();
  updateLeeNotification();
  refreshDirectionalWindUI();
  toast('Lee zone multiplier re-enabled');
}

// ═══════════════════════════════════════════════
//   ANALYSIS TAB
// ═══════════════════════════════════════════════
/** Polar multiplier SVG + V<sub>sit</sub> canvas rose — call after calc() when directional results change. */
function refreshDirectionalWindUI(){
  renderDirTable();
  if(S.R && S.R.faces) drawVsitRose();
}

function updateAnalysisTab(){
  refreshDirectionalWindUI();
}

function drawVsitRose(){
  const cv = document.getElementById('vsit-rose-canvas');
  if(!cv) return;
  const c = cv.getContext('2d');
  const w=cv.width, h=cv.height;
  const cx=w/2, cy=h/2, R=Math.min(w,h)/2-30;
  c.clearRect(0,0,w,h);

  c.fillStyle='rgba(15,15,35,.6)';
  c.beginPath();c.arc(cx,cy,R+8,0,Math.PI*2);c.fill();

  const maxV = Math.max(...S.Vsit_dir, 20);
  // Grid rings — only label outermost ring
  for(let i=1;i<=4;i++){
    const r = (i/4)*R;
    c.strokeStyle='rgba(255,255,255,.1)';c.lineWidth=1;
    c.beginPath();c.arc(cx,cy,r,0,Math.PI*2);c.stroke();
  }
  c.fillStyle='rgba(255,255,255,.3)';c.font='8px Arial';c.textAlign='center';
  c.fillText(Math.round(maxV)+' m/s', cx, cy-R-2);

  const dirs = S.dirs;
  const angles = [0,45,90,135,180,225,270,315];
  c.fillStyle='rgba(255,255,255,.55)';c.font='bold 10px Arial';c.textAlign='center';c.textBaseline='middle';
  for(let i=0;i<8;i++){
    const a = angles[i]*Math.PI/180 - Math.PI/2;
    c.fillText(dirs[i], cx+Math.cos(a)*(R+16), cy+Math.sin(a)*(R+16));
  }

  // Polygon fill
  c.beginPath();
  for(let i=0;i<8;i++){
    const a = angles[i]*Math.PI/180 - Math.PI/2;
    const r = (S.Vsit_dir[i]/maxV)*R;
    const x = cx + Math.cos(a)*r;
    const y = cy + Math.sin(a)*r;
    if(i===0) c.moveTo(x,y); else c.lineTo(x,y);
  }
  c.closePath();
  c.fillStyle='rgba(0,210,255,.2)';c.fill();
  c.strokeStyle='rgba(0,210,255,.8)';c.lineWidth=2;c.stroke();

  // Data points — only show value labels when values differ
  const allSame = S.Vsit_dir.every(v=>Math.abs(v-S.Vsit_dir[0])<0.1);
  for(let i=0;i<8;i++){
    const a = angles[i]*Math.PI/180 - Math.PI/2;
    const r = (S.Vsit_dir[i]/maxV)*R;
    const x = cx+Math.cos(a)*r;
    const y = cy+Math.sin(a)*r;
    c.fillStyle='#00d2ff';c.beginPath();c.arc(x,y,3,0,Math.PI*2);c.fill();
    if(!allSame){
      c.fillStyle='#fff';c.font='8px Arial';c.textAlign='center';
      c.fillText(S.Vsit_dir[i].toFixed(1), x + Math.cos(a)*12, y + Math.sin(a)*12);
    }
  }
  // If all same, show single centered value
  if(allSame){
    c.fillStyle='#00d2ff';c.font='bold 11px Arial';c.textAlign='center';c.textBaseline='middle';
    c.fillText(S.Vsit_dir[0].toFixed(1)+' m/s (all dirs)', cx, cy+R+22);
  }
}

// ═══════════════════════════════════════════════
//   PRESSURES TAB
// ═══════════════════════════════════════════════
function switchPressureTab(tab){
  S.activePressureTab = tab;
  const ph = document.getElementById('ptab-heatmap'), pl = document.getElementById('ptab-local');
  const cH = document.getElementById('pressure-heatmap-content'), cL = document.getElementById('pressure-local-content');
  if(ph) ph.classList.toggle('active', tab==='heatmap');
  if(pl) pl.classList.toggle('active', tab==='local');
  if(cH) cH.style.display = tab==='heatmap'?'':'none';
  if(cL) cL.style.display = tab==='local'?'':'none';
}

// ═══════════════════════════════════════════════
//   PRESSURES TAB — hover on formula cells → floating tip
// ═══════════════════════════════════════════════
let pressureCalcById = {};
let pressureCalcSeq = 0;
function resetPressureCalcRegistry(){
  pressureCalcById = {};
  pressureCalcSeq = 0;
}
function registerPressureCalc(payload){
  const id = 'pc' + (pressureCalcSeq++);
  pressureCalcById[id] = payload;
  return id;
}
function pcTd(valueStr, posNegClass, payload){
  const id = registerPressureCalc(payload);
  return `<td class="${posNegClass} pressure-calc-cell" data-pcid="${id}" title="Hover for calculation">${valueStr}</td>`;
}
/** V<sub>sit,β</sub> and q<sub>z</sub> — hover shows floating tip (same as p<sub>e</sub>, p<sub>net</sub>, etc.) */
function pcParamCell(valueStr, payload){
  const id = registerPressureCalc(payload);
  return `<td class="pressure-calc-cell" data-pcid="${id}" title="Hover for calculation">${valueStr}</td>`;
}
function buildQzPayload(qz, rho, VsitGov, govDirLabel){
  return {
    title: 'Design wind pressure q<sub>z</sub>',
    subtitle: 'ρ<sub>air</sub> = '+rho.toFixed(1)+' kg/m³ (Appendix D)',
    lines: [
      { html: 'q<sub>z</sub> = ½ ρ<sub>air</sub> V<sub>sit,β</sub>² / 1000' },
      { html: `= 0.5 × ${rho.toFixed(1)} × (${VsitGov.toFixed(2)})² / 1000` },
      { html: `= <b>${qz.toFixed(3)} kPa</b>` },
      { html: `V<sub>sit,β</sub> from max-of-3 envelope — governing sector: <strong>${govDirLabel}</strong>` }
    ]
  };
}
function buildVsitPayload(VR, Md, Mz, Ms, Mt, Vsit, dirLabel){
  return {
    title: 'Site wind speed V<sub>sit,β</sub>',
    subtitle: 'AS/NZS 1170.2 — Equation 2.2 ('+dirLabel+')',
    lines: [
      { html: 'V<sub>sit,β</sub> = V<sub>R</sub> M<sub>d</sub> M<sub>z,cat</sub> M<sub>s</sub> M<sub>t</sub>' },
      { html: `= ${VR.toFixed(3)} × ${Md.toFixed(3)} × ${Mz.toFixed(3)} × ${Ms.toFixed(3)} × ${Mt.toFixed(3)}` },
      { html: `= <b>${Vsit.toFixed(2)} m/s</b>` }
    ]
  };
}
function buildPeExternalPayload(label, qz, Cpe, Ka, Kp, pe){
  return {
    title: label,
    subtitle: 'External pressure p<sub>e</sub> = q<sub>z</sub> C<sub>p,e</sub> K<sub>a</sub> K<sub>p</sub> (Cl. 5.3–5.4)',
    lines: [
      { html: 'p<sub>e</sub> = q<sub>z</sub> × C<sub>p,e</sub> × K<sub>a</sub> × K<sub>p</sub>' },
      { html: `= ${qz.toFixed(3)} × ${Cpe.toFixed(4)} × ${Ka.toFixed(1)} × ${Kp.toFixed(4)}` },
      { html: `= <b>${pe.toFixed(2)} kPa</b>` }
    ]
  };
}
function buildCshpePayload(Cpe, Kl, Kp, Cshpe){
  return {
    title: 'Combined coefficient C<sub>shp,e</sub>',
    subtitle: 'Table 5.6 — local pressure factor K<sub>l</sub>',
    lines: [
      { html: 'C<sub>shp,e</sub> = C<sub>p,e</sub> × K<sub>l</sub> × K<sub>p</sub>' },
      { html: `= ${Cpe.toFixed(4)} × ${Kl.toFixed(1)} × ${Kp.toFixed(4)}` },
      { html: `= <b>${Cshpe.toFixed(2)}</b>` }
    ]
  };
}
function buildPeLocalPayload(qz, Cshpe, pe){
  return {
    title: 'External pressure p<sub>e</sub> (local)',
    subtitle: 'p<sub>e</sub> = q<sub>z</sub> × C<sub>shp,e</sub>',
    lines: [
      { html: 'p<sub>e</sub> = q<sub>z</sub> × C<sub>shp,e</sub>' },
      { html: `= ${qz.toFixed(3)} × ${Cshpe.toFixed(4)}` },
      { html: `= <b>${pe.toFixed(2)} kPa</b>` }
    ]
  };
}
function buildPnetPayload(pe, kce, pi, pnet, caseNum, qz, cpi, kci, Kv){
  const caseLbl = caseNum === 1
    ? 'Case 1 (−ve C<sub>p,i</sub>)'
    : 'Case 2 (+ve C<sub>p,i</sub>)';
  return {
    title: 'Net pressure p<sub>net</sub>',
    subtitle: `Table 5.5 — K<sub>c,e</sub> with ${caseLbl}`,
    lines: [
      { html: 'p<sub>net</sub> = p<sub>e</sub> × K<sub>c,e</sub> − p<sub>i</sub>' },
      { html: `p<sub>i</sub> = q<sub>z</sub> × C<sub>p,i</sub> × K<sub>c,i</sub> × K<sub>v</sub> = ${qz.toFixed(3)} × ${cpi.toFixed(2)} × ${kci.toFixed(1)} × ${Kv.toFixed(3)} = ${pi.toFixed(2)} kPa` },
      { html: `= (${pe.toFixed(4)}) × ${kce.toFixed(2)} − (${pi.toFixed(2)})` },
      { html: `= <b>${pnet.toFixed(2)} kPa</b>` }
    ]
  };
}
function openPressureCalcModal(payload){
  hidePressureHoverTip();
  const overlay = document.getElementById('pressure-calc-modal');
  const titleEl = document.getElementById('pressure-calc-title');
  const subEl = document.getElementById('pressure-calc-sub');
  const bodyEl = document.getElementById('pressure-calc-body');
  if(!overlay || !titleEl || !bodyEl) return;
  titleEl.innerHTML = payload.title || '';
  if(subEl){
    subEl.innerHTML = payload.subtitle || '';
    subEl.style.display = payload.subtitle ? '' : 'none';
  }
  bodyEl.innerHTML = (payload.lines || []).map(l => `<div class="pc-line">${l.html}</div>`).join('');
  overlay.classList.add('show');
  overlay.setAttribute('aria-hidden', 'false');
}
function closePressureCalcModal(){
  const overlay = document.getElementById('pressure-calc-modal');
  if(overlay){
    overlay.classList.remove('show');
    overlay.setAttribute('aria-hidden', 'true');
  }
}

let pressureHoverHideTimer = null;
function cancelHidePressureHoverTip(){
  clearTimeout(pressureHoverHideTimer);
  pressureHoverHideTimer = null;
}
function scheduleHidePressureHoverTip(){
  clearTimeout(pressureHoverHideTimer);
  pressureHoverHideTimer = setTimeout(()=>{ hidePressureHoverTip(); }, 100);
}
function showPressureHoverTip(td, payload){
  const tip = document.getElementById('pressure-hover-tip');
  if(!tip) return;
  tip.innerHTML = ''
    + `<div class="pressure-hover-tip-head">${payload.title || ''}</div>`
    + (payload.subtitle ? `<div class="pressure-hover-tip-sub">${payload.subtitle}</div>` : '')
    + `<div class="pressure-hover-tip-body">${(payload.lines || []).map(l => `<div class="pc-line">${l.html}</div>`).join('')}</div>`;
  tip.hidden = false;
  tip.style.display = 'block';
  tip.setAttribute('aria-hidden', 'false');
  requestAnimationFrame(()=>{
    const r = td.getBoundingClientRect();
    const tw = tip.offsetWidth;
    const th = tip.offsetHeight;
    let left = r.left + r.width / 2 - tw / 2;
    let top = r.bottom + 6;
    if(top + th > window.innerHeight - 10) top = Math.max(8, r.top - th - 6);
    if(left < 8) left = 8;
    if(left + tw > window.innerWidth - 8) left = Math.max(8, window.innerWidth - tw - 8);
    tip.style.left = left + 'px';
    tip.style.top = top + 'px';
  });
}
function hidePressureHoverTip(){
  const tip = document.getElementById('pressure-hover-tip');
  if(!tip) return;
  cancelHidePressureHoverTip();
  tip.hidden = true;
  tip.style.display = 'none';
  tip.setAttribute('aria-hidden', 'true');
  tip.innerHTML = '';
}
function initPressureHoverTipDelegate(){
  const root = document.getElementById('model-pressure-tables');
  if(!root || root.dataset.pressureHoverTipInit) return;
  root.dataset.pressureHoverTipInit = '1';
  root.addEventListener('mouseover', ev=>{
    const td = ev.target.closest('td.pressure-calc-cell');
    if(!td || !root.contains(td)) return;
    cancelHidePressureHoverTip();
    const id = td.getAttribute('data-pcid');
    if(!id || !pressureCalcById[id]) return;
    showPressureHoverTip(td, pressureCalcById[id]);
  });
  root.addEventListener('mouseout', ev=>{
    const td = ev.target.closest('td.pressure-calc-cell');
    if(!td || !root.contains(td)) return;
    const rel = ev.relatedTarget;
    if(rel && td.contains(rel)) return;
    scheduleHidePressureHoverTip();
  });
  root.addEventListener('scroll', ()=>hidePressureHoverTip(), {passive:true});
}
/** Tap-to-modal when hover is unavailable (touch / coarse pointer). */
function initPressureCalcClickDelegate(){
  if(typeof matchMedia === 'function' && matchMedia('(hover: hover)').matches) return;
  const root = document.getElementById('model-pressure-tables');
  if(!root || root.dataset.pressureCalcInit) return;
  root.dataset.pressureCalcInit = '1';
  root.addEventListener('click', function(ev){
    const td = ev.target.closest('td.pressure-calc-cell');
    if(!td || !root.contains(td)) return;
    const id = td.getAttribute('data-pcid');
    if(!id || !pressureCalcById[id]) return;
    openPressureCalcModal(pressureCalcById[id]);
  });
}

function findFaceMapMesh(uuid){
  let found = null;
  const visit = g=>{
    if(!g) return;
    g.traverse(o=>{ if(o.isMesh && o.uuid===uuid) found=o; });
  };
  visit(grpBuild);
  if(!found) visit(uploadedModelGroup);
  return found;
}

function pressureTableFkAttr(keys){
  if(!keys) return '';
  const arr = (Array.isArray(keys)?keys:[keys]).filter(Boolean);
  if(!arr.length) return '';
  return ' data-fk="'+arr.join(' ')+'"';
}

const WALL_PRESSURE_LABEL_TO_FK = {Windward:'windward','Side Wall L':'sidewall1','Side Wall R':'sidewall2',Leeward:'leeward'};

function roofExternalRowToFaceKeys(name){
  if(!name) return null;
  if(name.startsWith('Upwind')) return 'roof_ww';
  if(name.startsWith('Downwind')) return 'roof_lw';
  if(name.startsWith('R crosswind (hip)')) return ['roof_hip_l','roof_hip_r'];
  if(name.startsWith('R crosswind')) return ['roof_ww','roof_lw'];
  return null;
}

function roofLocalRowToFaceKeys(label){
  if(!label) return null;
  if(label.includes('R crosswind (hip)')) return ['roof_hip_l','roof_hip_r'];
  if(/^R\s*\(/.test(label)) return ['roof_ww','roof_lw'];
  if(label.includes('Upwind')) return 'roof_ww';
  if(label.includes('Downwind')) return 'roof_lw';
  return null;
}

function highlightMeshesForFaceKeys(keys){
  if(!keys || !faceMap || !faceMap.size) return;
  const arr = Array.isArray(keys) ? keys : [keys];
  const set = new Set(arr.filter(Boolean));
  faceMap.forEach((data, uuid)=>{
    if(!set.has(data.key)) return;
    const obj = findFaceMapMesh(uuid);
    if(obj && obj.material && obj.material.emissive) obj.material.emissive.setHex(0x333355);
  });
}

function highlightMeshesForFaceKeysAndZone(keys, zoneRef){
  if(!keys || !faceMap || !faceMap.size) return;
  const arr = Array.isArray(keys) ? keys : [keys];
  const set = new Set(arr.filter(Boolean));
  const targetZone = zoneRef && String(zoneRef).trim();
  if(!targetZone){
    highlightMeshesForFaceKeys(keys);
    return;
  }
  faceMap.forEach((data, uuid)=>{
    if(!set.has(data.key)) return;
    if((data.klZone || '') !== targetZone) return;
    const obj = findFaceMapMesh(uuid);
    if(obj && obj.material && obj.material.emissive) obj.material.emissive.setHex(0x333355);
  });
}

function clearAllFaceMapMeshEmissive(){
  if(!faceMap || !faceMap.size) return;
  faceMap.forEach((data, uuid)=>{
    const obj = findFaceMapMesh(uuid);
    if(obj && obj.material && obj.material.emissive) obj.material.emissive.setHex(0);
  });
}

function setPressureTableRowModelHighlight(tr){
  if(pressureTableHoverTr === tr) return;
  clearPressureTableRowModelHighlight();
  pressureTableHoverTr = tr;
  if(!tr) return;
  const raw = tr.getAttribute('data-fk');
  if(!raw) return;
  const keys = raw.trim().split(/\s+/).filter(Boolean);
  const zoneRef = tr.getAttribute('data-zone');
  if(keys.length){
    if(zoneRef) highlightMeshesForFaceKeysAndZone(keys, zoneRef);
    else highlightMeshesForFaceKeys(keys);
  }
}

function clearPressureTableRowModelHighlight(){
  pressureTableHoverTr = null;
  clearAllFaceMapMeshEmissive();
}

function initPressureTableRowHoverDelegate(){
  const root = document.getElementById('model-pressure-tables');
  if(!root || root.dataset.ptRowHoverInit) return;
  root.dataset.ptRowHoverInit = '1';
  root.addEventListener('mouseover', ev=>{
    const tr = ev.target.closest('tr[data-fk]');
    if(tr && root.contains(tr)) setPressureTableRowModelHighlight(tr);
  });
  root.addEventListener('mouseout', ev=>{
    const tr = ev.target.closest('tr[data-fk]');
    if(!tr || !root.contains(tr)) return;
    const rel = ev.relatedTarget;
    if(rel && tr.contains(rel)) return;
    clearPressureTableRowModelHighlight();
  });
  root.addEventListener('scroll', ()=>clearPressureTableRowModelHighlight(), {passive:true});
}

function setLimitState(mode){
  document.querySelectorAll('.limitstate-btn').forEach(b=>{
    b.classList.toggle('active', b.dataset.limit === mode);
  });
  recalcPressures();
}

function recalcPressures(){
  const R = S.R;
  if(!R || !R.faces) return;

  clearPressureTableRowModelHighlight();
  hidePressureHoverTip();
  resetPressureCalcRegistry();

  const activeLimitBtn = document.querySelector('.limitstate-btn.active');
  const isUlt = !activeLimitBtn || activeLimitBtn.dataset.limit === 'ultimate';
  const VR = isUlt ? S.windSpeed : S.svcVr;
  const h = R.h;

  // Wind direction from Building tab slider (same 8 sectors as site multipliers)
  const theta = S.windAngle;
  const dirIdx = Math.round(((theta % 360 + 360) % 360) / 45) % 8;

  // Use the per-direction Vsit from the wind rose (consistent with what's displayed)
  const Md    = effectiveMd(dirIdx);
  const Mzcat = effectiveMzcat(dirIdx);
  const Ms    = S.Ms[dirIdx];
  const Mt    = S.Mt[dirIdx];
  const Mlee  = S.Mlee[dirIdx] || 1;
  const VRdir = isUlt ? VR : S.svcVr;
  const rhoAir = 1.2;
  const adjL = (dirIdx + 7) % 8, adjR = (dirIdx + 1) % 8;
  const VsitC = VRdir * Md * Mzcat * Ms * Mt;
  const VsitL = VRdir * effectiveMd(adjL) * effectiveMzcat(adjL) * S.Ms[adjL] * S.Mt[adjL];
  const VsitR = VRdir * effectiveMd(adjR) * effectiveMzcat(adjR) * S.Ms[adjR] * S.Mt[adjR];
  const qzC = 0.5 * rhoAir * VsitC * VsitC / 1000;
  const qzL = 0.5 * rhoAir * VsitL * VsitL / 1000;
  const qzR = 0.5 * rhoAir * VsitR * VsitR / 1000;
  let qz = Math.max(qzC, qzL, qzR);
  const qzEps = 1e-7;
  let govIdx = dirIdx;
  let govVsit = VsitC;
  if(Math.abs(qz - qzC) < qzEps){ govIdx = dirIdx; govVsit = VsitC; }
  else if(Math.abs(qz - qzL) < qzEps){ govIdx = adjL; govVsit = VsitL; }
  else { govIdx = adjR; govVsit = VsitR; }
  const gMd = effectiveMd(govIdx);
  const gMz = effectiveMzcat(govIdx);
  const gMs = S.Ms[govIdx];
  const gMt = S.Mt[govIdx];
  const govDirLabel = S.dirs[govIdx];
  const qzCol = pcParamCell(qz.toFixed(3), buildQzPayload(qz, rhoAir, govVsit, govDirLabel));
  const vsitCol = pcParamCell(govVsit.toFixed(2), buildVsitPayload(VRdir, gMd, gMz, gMs, gMt, govVsit, govDirLabel));

  const manualOn = document.getElementById('cpi-manual-toggle')?.checked;
  const cpiResult = getCpiCasesForDesign();
  const cpi1 = cpiResult.cpi1;
  const cpi2 = cpiResult.cpi2;
  if(!manualOn){
    const el1 = document.getElementById('auto-cpi1');
    const el2 = document.getElementById('auto-cpi2');
    const clauseEl = document.getElementById('perm-clause');
    if(el1) el1.textContent = (cpi1>=0?'+':'')+cpi1.toFixed(2);
    if(el2) el2.textContent = (cpi2>=0?'+':'')+cpi2.toFixed(2);
    if(clauseEl) clauseEl.textContent = cpiResult.clause;
    if(permCondition==='auto'){
      document.querySelectorAll('.perm-btn').forEach(b=>{
        if(b.dataset.perm!=='auto'){
          b.classList.toggle('detected', !isKlDesignMode() && b.dataset.perm===cpiResult.detected);
        }
      });
    }
  }
  const modeBanner = document.getElementById('cpi-mode-banner');
  if(modeBanner){
    if(manualOn){
      modeBanner.innerHTML = 'Cp,i: <b>manual</b> — overrides Table 5.1 auto and Local Pressure Zones defaults (−0.3 / +0.2).';
    } else if(isKlDesignMode()){
      const mdNote = KL_DESIGN_MD_REGIONS.has(S.region)
        ? ' Cl 3.3(b): <b>M<sub>d</sub> = 1</b> (B2/C/D).'
        : ' (M<sub>d</sub> from Table 3.2 in this region.)';
      modeBanner.innerHTML =
        'Design: <b>Local Pressure Zones on</b> — Table 5.6; default C<sub>p,i</sub> −0.3 / +0.2.' + mdNote +
        ' <i>No separate cladding/glazing toggles.</i>';
    } else {
      modeBanner.textContent = '';
    }
  }
  suggestKcFromPermeability();
  validateKcProduct();
  renderKcDiagram(document.getElementById('kc-design-case')?.value || 'f');
  const kci1 = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;
  const kci2 = kci1;
  const kce1 = parseFloat(document.getElementById('kce1-val')?.value) || 0.8;
  const kce2 = kce1;

  // Clause 5.3.4 — Kv open area/volume factor
  const { volGeom, vol, overridden: volumeOverridden } = prepareVolumeForKv();
  const wwWallArea = S.width * S.height;
  const lwWallArea = S.width * S.height;
  const swWallArea = S.depth * S.height;
  const aWW = S.openWW / 100 * wwWallArea;
  const aLW = S.openLW / 100 * lwWallArea;
  const aSW = S.openSW / 100 * swWallArea;
  const openArea = aWW + aLW + aSW * 2;
  const A = Math.max(aWW, aLW, aSW); // largest wall opening area (governs r per Cl 5.3.4)

  const kvAuto = computeKvAuto(vol, A);
  const Kv = kvAuto.Kv;

  // Internal pressures: pi = qz × Cpi × Kci × Kv (Clause 5.3)
  S.R.Kv = Kv;
  const pi1 = qz * cpi1 * kci1 * Kv;
  const pi2 = qz * cpi2 * kci2 * Kv;
  const qzLpz = qz;
  const cpi1Lpz = manualOn ? cpi1 : -0.3;
  const cpi2Lpz = manualOn ? cpi2 : 0.2;
  const pi1Lpz = qzLpz * cpi1Lpz * kci1 * Kv;
  const pi2Lpz = qzLpz * cpi2Lpz * kci2 * Kv;

  const elPi1 = document.getElementById('pi1-display');
  const elPi2 = document.getElementById('pi2-display');
  const elOa = document.getElementById('open-area-display');
  const elKv = document.getElementById('kv-display');
  const elKvComp = document.getElementById('kv-computed-display');
  const elA = document.getElementById('kv-a-display');
  if(elPi1) elPi1.textContent = pi1.toFixed(2)+' kPa';
  if(elPi2) elPi2.textContent = pi2.toFixed(2)+' kPa';
  if(elOa) elOa.textContent = openArea.toFixed(1)+' m²';
  if(elKv) elKv.textContent = Kv.toFixed(3);
  if(elKvComp) elKvComp.textContent = kvAuto.Kv.toFixed(3);
  if(elA) elA.textContent = A > 0 ? A.toFixed(2)+' m²' : '—';
  updateKvEquationPopupHtml(vol, volGeom, A, kvAuto.ratio, kvAuto.zone, kvAuto.Kv, Kv, qz, cpi1, kci1, pi1, volumeOverridden);

  const kaCellTitle = 'K_a — area reduction factor (Table 5.4). p_e = q_z C_p,e K_a K_p (Cl. 5.3–5.4).';
  const Ka = 1.0;
  const useTable56KlForLocalTables = true;
  const isGlazing  = false;
  const hRef = R.h;

  // Determine effective dimensions based on wind angle vs building orientation
  // relAngle: wind direction relative to building front face
  const buildAngle = S.mapBuildingAngle || 0;
  const relAngle = ((theta - buildAngle) % 360 + 360) % 360;
  // If wind hits the width face (front/back): 315-45° or 135-225°
  // If wind hits the depth face (sides): 45-135° or 225-315° — match calc() / getWindFaceMap (>= at 45, 225)
  const hitsSide = (relAngle >= 45 && relAngle < 135) || (relAngle >= 225 && relAngle < 315);
  const effW = hitsSide ? S.depth : S.width;
  const effD = hitsSide ? S.width : S.depth;

  // Recompute Cp values for this orientation
  const rHD = hRef / effD;
  const db = effD / effW;
  const CpWW = hRef > 25 ? 0.8 : 0.7;
  const CpLW = leewardCp(db, S.pitch);

  // Roof Cp,e — both min and max values; Table 5.3(C) D vs R (hip) for leeward vs crosswind R
  let CpRW_min, CpRW_max, CpRL_D, CpRL_R;
  // Determine if wind is along or across the ridge for this windward face
  const buildAngleRoof = S.mapBuildingAngle || 0;
  const relAngleRoof = ((theta - buildAngleRoof) % 360 + 360) % 360;
  const isAlongRidge = (relAngleRoof >= 45 && relAngleRoof < 135) || (relAngleRoof >= 225 && relAngleRoof < 315);
  const rtRoof = S.roofType || 'gable';
  const monoHighWindward = rtRoof === 'monoslope' && !isAlongRidge && relAngle >= 135 && relAngle < 225;

  // Table 5.3(A) R crosswind zones — except hip + α ≥ 10° → Table 5.3(C) for R (not distance bands from (A))
  // Monoslope: R strips when wind ∥ ridge, or when wind hits high (back) wall (Fig 5.2 R); else U/D rows for low-eave windward
  const useTableA_UD = S.pitch < 10 || rtRoof === 'monoslope' || rtRoof === 'flat';
  const rCrosswindSlopeR_TableC_UI = rtRoof === 'hip' && S.pitch >= 10;
  let crosswindZones = null;
  const ovR = S.overhang || 0;
  const rwSegR = isAlongRidge
    ? ((S.roofType==='gable'||S.roofType==='hip') ? (S.depth/2 + ovR) : (S.depth + 2*ovR))
    : (effW + 2*ovR);
  if((isAlongRidge || rtRoof === 'flat' || (S.pitch < 10 && rtRoof !== 'monoslope') || monoHighWindward) && !rCrosswindSlopeR_TableC_UI){
    const seg = roofCrosswindCpZonesWallSegment(rHD, hRef, -ovR, effD + ovR, rwSegR);
    crosswindZones = seg.map(z=>{
      const d0 = z.uLo < -0.001 ? '0 (wall incl. '+Math.abs(z.uLo).toFixed(2)+' m overhang)' : z.uDispLo.toFixed(2);
      return {
        dist: d0+' to '+z.uDispHi.toFixed(2)+' m from wall · '+z.dist,
        Cpe_min: z.Cpe_pair[0],
        Cpe_max: z.Cpe_pair[1]
      };
    });
  }

  if(useTableA_UD){
    if(rHD <= 0.5){ CpRW_min=-0.9; CpRW_max=-0.4; }
    else if(rHD <= 1.0){ CpRW_min=-0.9-(rHD-0.5)/0.5*0.4; CpRW_max=-0.4-(rHD-0.5)/0.5*0.35; }
    else { CpRW_min=-1.3; CpRW_max=-0.75; }
    CpRL_D=-0.5;
    CpRL_R=-0.5;
  } else {
    CpRW_min = roofUpwindCp(S.pitch, rHD);
    CpRW_max = roofUpwindCpMax(S.pitch, rHD);
    CpRL_D = roofDownwindSlopeD(S.pitch, rHD);
    CpRL_R = roofCrosswindHipR(S.pitch, rHD);
  }

  // Kr parapet reduction factor (Table 5.7)
  const wMin = Math.min(effW, effD);
  const Kr = calcKr(S.parapet, hRef, wMin);
  const CpRW_kr_min = CpRW_min < 0 ? CpRW_min * Kr : CpRW_min;
  const CpRW_kr_max = CpRW_max < 0 ? CpRW_max * Kr : CpRW_max;
  const CpRL_kr_D = CpRL_D < 0 ? CpRL_D * Kr : CpRL_D;
  const CpRL_kr_R = CpRL_R < 0 ? CpRL_R * Kr : CpRL_R;

  const w = effW, d = effD;

  // Table 5.2(C) — sidewall Cp,e zones by distance from windward edge
  const swZones = [
    {dist:'0 to 1h ('+hRef.toFixed(1)+'m)',   Cpe:-0.65},
    {dist:'1h to 2h ('+(2*hRef).toFixed(1)+'m)', Cpe:-0.5},
    {dist:'2h to 3h ('+(3*hRef).toFixed(1)+'m)', Cpe:-0.3},
    {dist:'> 3h',                               Cpe:-0.2}
  ];
  const wallData = [];
  if(hRef > 25){
    const wb = windwardWallHeightBands(S.height);
    if(wb.length){
      wb.forEach(({zLo, zHi})=>{
        const zMid = (zLo + zHi) / 2;
        wallData.push({
          name:'Windward',
          dist: zLo.toFixed(1)+'–'+zHi.toFixed(1)+' m AGL',
          Ka: Ka,
          Cpe: CpWW,
          applyKp: false,
          qzOverride: envelopeQzWindPressureAtZ(VRdir, dirIdx, zMid)
        });
      });
    } else {
      wallData.push({name:'Windward', dist:'Full face', Ka:Ka, Cpe:CpWW, applyKp:false});
    }
  } else {
    wallData.push({name:'Windward', dist:'Full face', Ka:Ka, Cpe:CpWW, applyKp:false});
  }
  // Add sidewall zones for both side walls
  ['Side Wall L','Side Wall R'].forEach(swName => {
    swZones.forEach(z => {
      wallData.push({name:swName, dist:z.dist, Ka:Ka, Cpe:z.Cpe, applyKp:true});
    });
  });
  wallData.push({name:'Leeward', dist:'Full face', Ka:Ka, Cpe:CpLW, applyKp:false});

  const qzWindwardForLpz = hRef > 25 ? windwardHeightMeanQz(VRdir, dirIdx, S.height) : qzLpz;

  // Kp uniform value (for non-auto modes)
  const KpUniform = getKpValue();

  let ewb = '';
  wallData.forEach(f=>{
    const Kp = (f.applyKp && f.Cpe < 0) ? KpUniform : 1.0;
    const qzRow = f.qzOverride != null ? f.qzOverride : qz;
    const pe = qzRow * f.Cpe * f.Ka * Kp;
    const pnet1 = pe * kce1 - pi1;
    const pnet2 = pe * kce2 - pi2;
    const kpStr = Kp < 1 ? ` <span style="color:#e67e22">(K<sub>p</sub>=${Kp.toFixed(2)})</span>` : '';
    const peCls = pe > 0 ? 'pos' : 'neg';
    const pn1Cls = pnet1 > 0 ? 'pos' : 'neg';
    const pn2Cls = pnet2 > 0 ? 'pos' : 'neg';
    const wFk = WALL_PRESSURE_LABEL_TO_FK[f.name];
    const wFkAttr = wFk ? ` data-fk="${wFk}"` : '';
    ewb += `<tr${wFkAttr}>
      <td class="label-cell">${f.name}${kpStr}</td><td>${f.dist}</td>
      <td title="${kaCellTitle}">${f.Ka.toFixed(1)}</td><td>${f.Cpe.toFixed(2)}</td>
      <td>${kci1.toFixed(2)}</td><td>${kce1.toFixed(2)}</td><td>${Kp.toFixed(2)}</td>
      ${vsitCol}${qzCol}
      ${pcTd(pe.toFixed(2), peCls, buildPeExternalPayload('External wall — p<sub>e</sub>', qzRow, f.Cpe, f.Ka, Kp, pe))}
      ${pcTd(pnet1.toFixed(2), pn1Cls, buildPnetPayload(pe, kce1, pi1, pnet1, 1, qz, cpi1, kci1, Kv))}
      ${pcTd(pnet2.toFixed(2), pn2Cls, buildPnetPayload(pe, kce2, pi2, pnet2, 2, qz, cpi2, kci2, Kv))}
    </tr>`;
  });
  const extWalls = document.getElementById('ext-walls-body');
  if(extWalls) extWalls.innerHTML = ewb;

  // Clause 5.4.4 — dimension a (walls vs roofs)
  const a = Math.min(0.2 * Math.min(w, d), S.height);
  const hRef_roof = hRef;
  const a_roof = (hRef_roof/w >= 0.2 || hRef_roof/d >= 0.2) ? 0.2 * Math.min(w, d) : 2 * hRef_roof;
  // Table 5.6 — Local pressure factor Kl (face-specific zones to match model mesh)
  const rWall = S.height / Math.min(w, d);
  const isCladding = !!S.showPressureMap;
  const fmtA = (n)=>Number(n).toFixed(1);
  const fmtA2 = (n)=>Number(n).toFixed(1);
  const sideWallKlZonesFromTable56 = ()=>{
    // Use the same breakpoints and Kl logic as sidewallKlTable56()/buildKlWalls.
    // Distances are from windward corner along the wall strip axis.
    const edges = [0, 0.5 * a, a, 2 * a];
    const zones = [];
    for(let i = 0; i < edges.length - 1; i++){
      const dMid = (edges[i] + edges[i+1]) / 2;
      const z = sidewallKlTable56(dMid, a, rWall, isCladding || useTable56KlForLocalTables, isGlazing);
      zones.push({
        ref: z.klZone || 'Other',
        Kl: z.Kl,
        area: i === 0
          ? '≤ 0.25a² ('+fmtA2(0.25*a*a)+'m²)'
          : (i === 1 ? '0.25a² to a²' : 'a² to 4a²'),
        dist: i === 0
          ? '0 to 0.5a ('+fmtA(0.5*a)+'m)'
          : (i === 1
              ? '0.5a to a ('+fmtA(a)+'m)'
              : 'a to 2a ('+fmtA(2*a)+'m)')
      });
    }
    // Merge consecutive bins that evaluate to the same Table 5.6 zone.
    const merged = [];
    zones.forEach(z=>{
      const prev = merged[merged.length - 1];
      if(prev && prev.ref === z.ref && Math.abs(prev.Kl - z.Kl) < 1e-9){
        // Keep the broader area wording; distance spans are still explicit and ordered.
        prev.dist = prev.dist.split(' to ')[0] + ' to ' + z.dist.split(' to ')[1];
      } else {
        merged.push({...z});
      }
    });
    return merged;
  };
  const wallKlZonesByFace = (faceLabel)=>{
    if(!useTable56KlForLocalTables && !isGlazing){
      return [{ref:'Other', area:'All', Kl:1.0, dist:'All'}];
    }
    if(faceLabel === 'Windward'){
      return [{ref:'WA1', area:'≤ a² ('+a.toFixed(1)+'×'+a.toFixed(1)+'m)', Kl:1.5, dist:'0 to a ('+a.toFixed(1)+'m)'}];
    }
    if(faceLabel === 'Leeward'){
      return [{ref:'Other', area:'All', Kl:1.0, dist:'All'}];
    }
    return sideWallKlZonesFromTable56();
  };
  let lwb = '';
  ['Windward','Leeward','Side Wall L','Side Wall R'].forEach(name=>{
    const wd = wallData.find(f=>f.name===name);
    if(!wd) return;
    const lwFk = WALL_PRESSURE_LABEL_TO_FK[name];
    const wdFace = lwFk ? S.R.faces[lwFk] : null;
    const Cpe = wdFace && Number.isFinite(wdFace.Cp_e) ? wdFace.Cp_e : wd.Cpe;
    const isKpFace = wd.applyKp;
    const qzLocal = name === 'Windward' && hRef > 25 ? qzWindwardForLpz : qzLpz;
    const lwFkAttr = lwFk ? ` data-fk="${lwFk}"` : '';
    const faceKlZones = wallKlZonesByFace(name);
    faceKlZones.forEach(kl=>{
      const rowFaceCase1 = wdFace ? {...wdFace, Cp_i:cpi1Lpz} : {Cp_e:Cpe, Cp_i:cpi1Lpz};
      const rowFaceCase2 = wdFace ? {...wdFace, Cp_i:cpi2Lpz} : {Cp_e:Cpe, Cp_i:cpi2Lpz};
      const terms1 = localPressureTerms(rowFaceCase1, kl.Kl, isKpFace, qzLocal, cpi1Lpz);
      const terms2 = localPressureTerms(rowFaceCase2, kl.Kl, isKpFace, qzLocal, cpi2Lpz);
      const Kp = terms1.Kp;
      const Cshpe = qzLocal > 0 ? terms1.pe / qzLocal : (Cpe * kl.Kl * Kp);
      const pe = terms1.pe;
      const pnet1 = pNetTable55Case1(terms1.pe, kce1, terms1.pi);
      const pnet2 = pNetTable55Case1(terms2.pe, kce2, terms2.pi);
      const csCls = Cshpe > 0 ? 'pos' : 'neg';
      const peCls = pe > 0 ? 'pos' : 'neg';
      const pn1Cls = pnet1 > 0 ? 'pos' : 'neg';
      const pn2Cls = pnet2 > 0 ? 'pos' : 'neg';
      const zoneAttr = kl.ref ? ` data-zone="${String(kl.ref).replace(/"/g, '&quot;')}"` : '';
      lwb += `<tr${lwFkAttr}${zoneAttr}>
        <td class="label-cell">${name}</td><td>${kl.dist}</td><td>${kl.ref}</td><td>${kl.area}</td>
        <td>${kl.Kl.toFixed(1)}</td>
        <td>${kci1.toFixed(2)}</td><td>${kce1.toFixed(2)}</td><td>${Kp.toFixed(2)}</td>
        ${vsitCol}${qzCol}
        ${pcTd(Cshpe.toFixed(2), csCls, buildCshpePayload(Cpe, kl.Kl, Kp, Cshpe))}
        ${pcTd(pe.toFixed(2), peCls, buildPeLocalPayload(qzLocal, Cshpe, pe))}
        ${pcTd(pnet1.toFixed(2), pn1Cls, buildPnetPayload(pe, kce1, terms1.pi, pnet1, 1, qzLocal, cpi1Lpz, kci1, Kv))}
        ${pcTd(pnet2.toFixed(2), pn2Cls, buildPnetPayload(pe, kce2, terms2.pi, pnet2, 2, qzLocal, cpi2Lpz, kci2, Kv))}
      </tr>`;
    });
  });
  const localWalls = document.getElementById('local-walls-body');
  if(localWalls) localWalls.innerHTML = lwb;

  let roofData;
  const krStr = Kr < 1 ? ` <span style="color:#9b59b6">(K<sub>r</sub>=${Kr.toFixed(2)})</span>` : '';

  if(crosswindZones && crosswindZones.length){
    roofData = [];
    const rTag = (isAlongRidge || monoHighWindward) ? 'T5.3(A) R' : 'T5.3(A)';
    crosswindZones.forEach(z => {
      const cMin = z.Cpe_min * Kr;
      const cMax = z.Cpe_max < 0 ? z.Cpe_max * Kr : z.Cpe_max;
      roofData.push({name:'R crosswind (min)', dist:z.dist, Ka:Ka, Cpe:cMin, clause:rTag});
      roofData.push({name:'R crosswind (max)', dist:z.dist, Ka:Ka, Cpe:cMax, clause:rTag});
    });
  } else if(rCrosswindSlopeR_TableC_UI && isAlongRidge){
    roofData = [
      {name:'R crosswind (hip)', dist:'Fig 5.2 — full slope', Ka:Ka, Cpe:CpRL_kr_R, clause:'T5.3(C) R'}
    ];
  } else {
    const uTag = useTableA_UD ? 'T5.3(A)' : 'T5.3(B)';
    const dTag = useTableA_UD ? 'T5.3(A)' : 'T5.3(C)';
    roofData = [
      {name:'Upwind (min)',  dist:'0 to d/2 ('+(d/2).toFixed(1)+'m)', Ka:Ka, Cpe:CpRW_kr_min, clause:uTag},
      {name:'Upwind (max)',  dist:'0 to d/2 ('+(d/2).toFixed(1)+'m)', Ka:Ka, Cpe:CpRW_kr_max, clause:uTag},
      {name:'Downwind',      dist:'d/2 to d ('+(d).toFixed(1)+'m)',   Ka:Ka, Cpe:CpRL_kr_D,      clause:dTag}
    ];
  }

  let erb = '';
  roofData.forEach(f=>{
    const Kp = (f.Cpe < 0) ? KpUniform : 1.0;
    const pe = qz * f.Cpe * f.Ka * Kp;
    const pnet1 = pe * kce1 - pi1;
    const pnet2 = pe * kce2 - pi2;
    const kpStr = Kp < 1 ? ` <span style="color:#e67e22">(K<sub>p</sub>=${Kp.toFixed(2)})</span>` : '';
    const peCls = pe > 0 ? 'pos' : 'neg';
    const pn1Cls = pnet1 > 0 ? 'pos' : 'neg';
    const pn2Cls = pnet2 > 0 ? 'pos' : 'neg';
    const rFkAttr = pressureTableFkAttr(roofExternalRowToFaceKeys(f.name));
    erb += `<tr${rFkAttr}>
      <td class="label-cell">${f.name}${kpStr}${krStr}</td><td>${f.dist}</td>
      <td title="${kaCellTitle}">${f.Ka.toFixed(1)}</td><td>${f.Cpe.toFixed(2)}</td>
      <td>${kci1.toFixed(2)}</td><td>${kce1.toFixed(2)}</td><td>${Kp.toFixed(2)}</td>
      ${vsitCol}${qzCol}
      ${pcTd(pe.toFixed(2), peCls, buildPeExternalPayload('External roof — p<sub>e</sub>', qz, f.Cpe, f.Ka, Kp, pe))}
      ${pcTd(pnet1.toFixed(2), pn1Cls, buildPnetPayload(pe, kce1, pi1, pnet1, 1, qz, cpi1, kci1, Kv))}
      ${pcTd(pnet2.toFixed(2), pn2Cls, buildPnetPayload(pe, kce2, pi2, pnet2, 2, qz, cpi2, kci2, Kv))}
    </tr>`;
  });
  const extRoof = document.getElementById('ext-roof-body');
  if(extRoof) extRoof.innerHTML = erb;

  // Table 5.6 — Local pressure factor Kl for roof
  // Generic path: derive rows from rendered roof local zones (faceMap) so all roof
  // shapes (flat/gable/hip/monoslope) stay in sync with the model.
  const roofZoneMeta = (ref, ar)=>{
    switch(ref){
      case 'RC1': return {dist:'< a from two edges', area:'≤ a² ('+ar.toFixed(1)+'×'+ar.toFixed(1)+'m)'};
      case 'RC2': return {dist:'< a from ridge & edge', area:'≤ a² ('+ar.toFixed(1)+'×'+ar.toFixed(1)+'m)'};
      case 'RA4': return {dist:'< 0.5a from ridge', area:'≤ a²'};
      case 'RA3': return {dist:'0.5a to a from ridge', area:'≤ a²'};
      case 'RA2': return {dist:'0 to a ('+ar.toFixed(1)+'m)', area:'≤ a² ('+ar.toFixed(1)+'×'+ar.toFixed(1)+'m)'};
      case 'RA1': return {dist:'a to 2a ('+(2*ar).toFixed(1)+'m)', area:'a² to 4a²'};
      case 'MR':  return {dist:'> a', area:'> a²'};
      default:    return {dist:'All', area:'All'};
    }
  };
  const roofSurfaceLabel = (faceKey, fallbackName)=>{
    if(faceKey === 'roof_hip_l') return 'Hip face L';
    if(faceKey === 'roof_hip_r') return 'Hip face R';
    if(faceKey === 'roof_lw') return 'Leeward';
    if(faceKey === 'roof_ww') return 'Upwind';
    return fallbackName || 'Roof';
  };
  const roofZoneOrder = {RC1:1,RC2:2,RA4:3,RA3:4,RA2:5,RA1:6,MR:7,Other:8};
  const collectRoofLocalRowsFromModel = ()=>{
    const out = [];
    const seen = new Set();
    const roofKeys = new Set(['roof_ww','roof_lw','roof_hip_l','roof_hip_r']);
    faceMap.forEach((data)=>{
      if(!data || !roofKeys.has(data.key)) return;
      const Kl = Number.isFinite(data.Kl) ? data.Kl : klFromKlZoneName(data.klZone);
      if(!(Kl > 0)) return;
      const ref = data.klZone || klZoneName(data.key, Kl) || 'Other';
      const rowKey = [data.key, ref, Kl.toFixed(3)].join('|');
      if(seen.has(rowKey)) return;
      seen.add(rowKey);
      out.push({
        faceKey: data.key,
        label: roofSurfaceLabel(data.key, data.name),
        ref,
        Kl,
        Cpe: Number.isFinite(data.Cp_e) ? data.Cp_e : (S.R.faces[data.key]?.Cp_e ?? 0)
      });
    });
    out.sort((a,b)=>{
      const kOrd = {'roof_ww':0,'roof_lw':1,'roof_hip_l':2,'roof_hip_r':3};
      const dk = (kOrd[a.faceKey] ?? 9) - (kOrd[b.faceKey] ?? 9);
      if(dk !== 0) return dk;
      return (roofZoneOrder[a.ref] ?? 99) - (roofZoneOrder[b.ref] ?? 99);
    });
    return out;
  };

  let lrbArr = [];
  const roofRows = collectRoofLocalRowsFromModel();
  if(roofRows.length){
    roofRows.forEach(row=>{
      const rlFkAttr = pressureTableFkAttr(row.faceKey);
      const roofFace = S.R.faces[row.faceKey];
      const rowFaceCase1 = roofFace ? {...roofFace, Cp_e:row.Cpe, Cp_i:cpi1Lpz} : {Cp_e:row.Cpe, Cp_i:cpi1Lpz};
      const rowFaceCase2 = roofFace ? {...roofFace, Cp_e:row.Cpe, Cp_i:cpi2Lpz} : {Cp_e:row.Cpe, Cp_i:cpi2Lpz};
      const terms1 = localPressureTerms(rowFaceCase1, row.Kl, true, qzLpz, cpi1Lpz);
      const terms2 = localPressureTerms(rowFaceCase2, row.Kl, true, qzLpz, cpi2Lpz);
      const Kp = terms1.Kp;
      const Cshpe = qzLpz > 0 ? terms1.pe / qzLpz : (row.Cpe * row.Kl * Kp);
      const pe = terms1.pe;
      const pn1 = pNetTable55Case1(terms1.pe, kce1, terms1.pi);
      const pn2 = pNetTable55Case1(terms2.pe, kce2, terms2.pi);
      const csCls = Cshpe > 0 ? 'pos' : 'neg';
      const peCls = pe > 0 ? 'pos' : 'neg';
      const pn1Cls = pn1 > 0 ? 'pos' : 'neg';
      const pn2Cls = pn2 > 0 ? 'pos' : 'neg';
      const meta = roofZoneMeta(row.ref, a_roof);
      const zoneAttr = row.ref ? ` data-zone="${String(row.ref).replace(/"/g, '&quot;')}"` : '';
      lrbArr.push(`<tr${rlFkAttr}${zoneAttr}>
        <td class="label-cell">${row.label}</td><td>${meta.dist}</td><td>${row.ref}</td><td>${meta.area}</td>
        <td>${row.Kl.toFixed(1)}</td>
        <td>${kci1.toFixed(2)}</td><td>${kce1.toFixed(2)}</td><td>${Kp.toFixed(2)}</td>
        ${vsitCol}${qzCol}
        ${pcTd(Cshpe.toFixed(2), csCls, buildCshpePayload(row.Cpe, row.Kl, Kp, Cshpe))}
        ${pcTd(pe.toFixed(2), peCls, buildPeLocalPayload(qzLpz, Cshpe, pe))}
        ${pcTd(pn1.toFixed(2), pn1Cls, buildPnetPayload(pe, kce1, terms1.pi, pn1, 1, qzLpz, cpi1Lpz, kci1, Kv))}
        ${pcTd(pn2.toFixed(2), pn2Cls, buildPnetPayload(pe, kce2, terms2.pi, pn2, 2, qzLpz, cpi2Lpz, kci2, Kv))}
      </tr>`);
    });
  }
  const localRoof = document.getElementById('local-roof-body');
  if(localRoof) localRoof.innerHTML = lrbArr.join('');
}

// ═══════════════════════════════════════════════
//   DETAILED REPORT SUBTAB — Full report preview (PDF behind paywall)
// ═══════════════════════════════════════════════
async function updateDocPreview(){
  const R = S.R;
  if(!R.faces) return;
  const el = document.getElementById('doc-preview');
  if(!el) return;

  let maxVsit=0,maxDir='';
  S.Vsit_dir.forEach((v,i)=>{if(v>maxVsit){maxVsit=v;maxDir=S.dirs[i]}});

  // Temporary: keep PDF preview unlocked while iterating on report changes.
  const isPro = true;
  let html = `
  <style>
    .rpt{font-family:'Segoe UI',Arial,sans-serif;color:#282828;line-height:1.5;max-width:800px;margin:0 auto}
    .rpt h1{color:#1a5276;font-size:24px;border-bottom:3px solid #1a5276;padding-bottom:6px;margin:24px 0 12px}
    .rpt h2{color:#1a5276;font-size:16px;background:#ebf5fb;padding:6px 10px;border-left:4px solid #2980b9;margin:18px 0 8px}
    .rpt table{border-collapse:collapse;width:100%;margin:8px 0;font-size:12px}
    .rpt th{background:#1a5276;color:white;padding:6px 8px;text-align:left;font-size:11px}
    .rpt td{border:1px solid #ddd;padding:5px 8px;font-size:11px}
    .rpt tr:nth-child(even){background:#f8fafe}
    .rpt .clause{color:#6464b4;font-style:italic;font-size:10px}
    .rpt .result{color:#1a5276;font-weight:700}
    .rpt .param-grid{display:grid;grid-template-columns:200px auto;gap:2px 12px;margin:6px 0}
    .rpt .param-grid dt{color:#666;font-size:11px}
    .rpt .param-grid dd{font-weight:600;font-size:11px;margin:0}
    .rpt .pro-cta{background:linear-gradient(135deg,#1a5276,#2980b9);color:white;padding:24px;border-radius:12px;text-align:center;margin:24px 0}
    .rpt .pro-cta h3{color:white;font-size:18px;margin-bottom:8px;text-transform:none;letter-spacing:0}
    .rpt .pro-cta p{color:rgba(255,255,255,.85);font-size:13px;margin-bottom:12px}
    .rpt .pro-cta button{background:white;color:#1a5276;border:none;padding:10px 28px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer}
    .rpt .pro-cta button:hover{opacity:.9}
    .rpt .blurred{filter:blur(5px);user-select:none;pointer-events:none}
    .rpt .locked-overlay{position:relative}
    .rpt .locked-overlay::after{content:'🔒 Pro Feature';position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);background:rgba(26,82,118,.9);color:white;padding:8px 20px;border-radius:8px;font-size:13px;font-weight:700;z-index:10}
    .rpt .pdf-preview-wrap{margin-top:16px;border:1px solid #d8e1e8;border-radius:10px;overflow:hidden;background:#fff}
    .rpt .pdf-preview-head{background:#f2f7fb;color:#1a5276;font-size:12px;font-weight:700;padding:8px 12px;border-bottom:1px solid #d8e1e8}
    .rpt .pdf-preview-frame{width:100%;height:1200px;border:0;display:block;background:#fff}
    .rpt .pdf-preview-loading{padding:18px 12px;font-size:12px;color:#666}
  </style>
  <div class="rpt">
  <h1>WIND LOAD ANALYSIS — DETAILED REPORT</h1>
  <p style="color:#666;font-size:12px">AS/NZS 1170.2:2021<br>
  Generated: ${new Date().toLocaleString()}</p>

  <h2>Project Overview</h2>
  <dl class="param-grid">
    <dt>Location</dt><dd>${S.address||S.lat.toFixed(4)+', '+S.lng.toFixed(4)}</dd>
    <dt>Wind Region</dt><dd>${S.region}</dd>
    <dt>Importance Level</dt><dd>${S.importance}</dd>
    <dt>Design life</dt><dd>${S.life} years</dd>
    <dt>Design ARI (ULS wind)</dt><dd>${S.ari} years (1170.0 Table 3.3)</dd>
    <dt>V<sub>R</sub> source</dt><dd>${S.vrManual?'Manual override':'Table 3.1'}</dd>
  </dl>

  <h2>Building</h2>
  <dl class="param-grid">
    <dt>Dimensions (W × D × H)</dt><dd>${S.width}m × ${S.depth}m × ${S.height}m</dd>
    <dt>Roof</dt><dd>${S.roofType} @ ${S.pitch}°</dd>
    <dt>Terrain Category</dt><dd>TC${S.terrainCat}</dd>
  </dl>

  <h2>Key Results</h2>
  <dl class="param-grid">
    <dt>Regional Wind Speed V<sub>R</sub></dt><dd class="result">${R.VR} m/s</dd>
    <dt>Site Wind Speed V<sub>sit</sub></dt><dd class="result">${R.Vsit.toFixed(1)} m/s</dd>
    <dt>Governing Direction</dt><dd>${maxDir}</dd>
    <dt>Design Pressure q<sub>z</sub></dt><dd class="result">${R.qz.toFixed(3)} kPa</dd>
  </dl>`;

  if(!isPro){
    html += `
    <div class="pro-cta">
      <h3>⚡ Unlock Full Professional Report</h3>
      <p>This local build includes the full PDF report with clause references, pressure coefficients, Table 5.6 local (K<sub>l</sub>) pressures, all cardinal-direction design pressures, and 3D visualisation.</p>
    </div>
    <h2>Detailed Report Contents (Pro)</h2>
    <div class="locked-overlay" style="position:relative">
    <div class="blurred">
    <table>
      <thead><tr><th>Face</th><th>C<sub>p,e</sub></th><th>p (kPa)</th><th>Force (kN)</th><th>Clause</th></tr></thead>
      <tbody>`;
    for(let k in R.faces){
      const f=R.faces[k];
      html+=`<tr><td>${f.name}</td><td>X.XX</td><td>X.XXX</td><td>XX.X</td><td>Table X.X</td></tr>`;
    }
    html+=`</tbody></table>
    </div></div>`;
  } else {
    // Pro users see the actual generated PDF embedded (1:1 with export).
    html += `
    <h2>Full Detailed Report (PDF Preview)</h2>
    <div class="pdf-preview-wrap">
      <div id="pdf-preview-loading" class="pdf-preview-loading">Generating full report preview...</div>
      <iframe id="pdf-preview-frame" class="pdf-preview-frame" title="Detailed PDF report preview" style="display:none"></iframe>
    </div>
    <p style="font-size:11px;color:#666;margin-top:12px">Use browser zoom controls for reading comfort, or click Export PDF Report to download.</p>`;
  }

  html += `</div>`;
  el.innerHTML = html;

  if(isPro){
    try{
      const blob = await generatePDF({ save: false, returnBlob: true, silent: true });
      const frame = document.getElementById('pdf-preview-frame');
      const loading = document.getElementById('pdf-preview-loading');
      if(frame){
        if(detailedPdfPreviewUrl){
          URL.revokeObjectURL(detailedPdfPreviewUrl);
          detailedPdfPreviewUrl = null;
        }
        detailedPdfPreviewUrl = URL.createObjectURL(blob);
        frame.src = detailedPdfPreviewUrl;
        frame.style.display = 'block';
      }
      if(loading) loading.style.display = 'none';
    } catch(err){
      const loading = document.getElementById('pdf-preview-loading');
      if(loading) loading.textContent = 'Could not generate PDF preview. Please try Export PDF Report.';
      console.warn('PDF preview render failed:', err);
    }
  }
}

// ═══════════════════════════════════════════════
//   MOUSE / TOUCH INTERACTION
// ═══════════════════════════════════════════════
function clearFaceInfoHideTimer(){
  if(faceInfoHideTimer){
    clearTimeout(faceInfoHideTimer);
    faceInfoHideTimer = null;
  }
}

function hideFaceInfoTooltip(){
  clearFaceInfoHideTimer();
  faceInfoAnchorUuid = null;
  const info = document.getElementById('face-info');
  if(info) info.style.display = 'none';
  lastUploadHoverUuid = null;
  const fiUploadRow = document.getElementById('fi-upload-diag-row');
  const fiOverrideRow = document.getElementById('fi-override-row');
  if(fiUploadRow) fiUploadRow.style.display = 'none';
  if(fiOverrideRow) fiOverrideRow.style.display = 'none';
  const fiRetrainRow = document.getElementById('fi-retrain-row');
  if(fiRetrainRow) fiRetrainRow.style.display = 'none';
  const fiZones = document.getElementById('fi-zones');
  if(fiZones){ fiZones.style.display = 'none'; fiZones.innerHTML = ''; fiZones.setAttribute('aria-hidden', 'true'); }
  if(hoveredFace){ hoveredFace.material.emissive&&hoveredFace.material.emissive.set(0); hoveredFace=null; }
}

/** Hide after pointer leaves mesh — delayed so you can move onto the floating panel (ray may miss briefly). */
function scheduleFaceInfoHide(){
  clearFaceInfoHideTimer();
  faceInfoHideTimer = setTimeout(()=>{
    faceInfoHideTimer = null;
    if(faceInfoPointerInside) return;
    hideFaceInfoTooltip();
  }, 420);
}

function onPointerMove(e){
  const r=renderer.domElement.getBoundingClientRect();
  mouse.x=((e.clientX-r.left)/r.width)*2-1;
  mouse.y=-((e.clientY-r.top)/r.height)*2+1;

  raycaster.setFromCamera(mouse,camera);
  // Raycast against parametric or uploaded model depending on which is active
  const rayTargets = (uploadedModelGroup && uploadedModelVisible && !parametricVisible)
    ? uploadedModelGroup.children : grpBuild.children;
  const hits=raycaster.intersectObjects(rayTargets,true).filter(i=>faceMap.has(i.object.uuid));
  const info=document.getElementById('face-info');

  if(hits.length){
    clearFaceInfoHideTimer();
    const obj=hits[0].object, data=faceMap.get(obj.uuid);
    if(data){
      hoveredFace&&hoveredFace!==obj&&hoveredFace.material.emissive&&hoveredFace.material.emissive.set(0);
      hoveredFace=obj;
      if(obj.material.emissive)obj.material.emissive.set(0x333355);
      info.style.display='block';
      const uid = obj.uuid;
      if(faceInfoAnchorUuid !== uid){
        faceInfoAnchorUuid = uid;
        info.style.left=(e.clientX+15)+'px';info.style.top=(e.clientY-10)+'px';
      }
      document.getElementById('fi-title').textContent=data.name;
      document.getElementById('fi-cpe').textContent=data.Cp_e.toFixed(2);
      const fiZones = document.getElementById('fi-zones');
      if(fiZones){
        if(data.zones && data.zones.length > 1){
          fiZones.style.display = 'block';
          fiZones.setAttribute('aria-hidden', 'false');
          const zh = (data.clause && String(data.clause).includes('5.2(C)')) ? 'Table 5.2(C) bands'
            : (data.clause && String(data.clause).includes('5.3')) ? 'Table 5.3 bands'
            : 'Distance zones';
          fiZones.innerHTML = `<div class="fi-zones-h">${zh}</div>` + data.zones.map(z=>{
            const c = Number.isFinite(z.Cpe) ? z.Cpe.toFixed(2) : '—';
            const a = Number.isFinite(z.area) ? z.area.toFixed(1) : '—';
            return `<div class="fi-zone-row"><span>${z.dist||'—'}</span><span>Cp,e ${c} · ${a} m²</span></div>`;
          }).join('');
        } else {
          fiZones.style.display = 'none';
          fiZones.setAttribute('aria-hidden', 'true');
          fiZones.innerHTML = '';
        }
      }
      let klVal = (data.Kl != null && Number.isFinite(data.Kl))
        ? data.Kl
        : (data.klZone ? klFromKlZoneName(data.klZone) : null);
      document.getElementById('fi-kl').textContent = klVal != null ? klVal.toFixed(1) : '—';
      // Show Kp if applicable (roofs & side walls with negative Cpe)
      const isRoofOrSide = data.key && (data.key.startsWith('roof') || data.key.startsWith('side'));
      const kpVal = (isRoofOrSide && data.Cp_e < 0) ? getKpValue() : 1.0;
      document.getElementById('fi-kp').textContent = S.Kp === 'auto' ? 'auto' : kpVal.toFixed(2);
      // Show Kr if applicable (roof faces only)
      const isRoof = data.key && data.key.startsWith('roof');
      const krVal = (isRoof && S.R.Kr !== undefined) ? S.R.Kr : 1.0;
      document.getElementById('fi-kr').textContent = krVal < 1 ? krVal.toFixed(2) : '-';
      document.getElementById('fi-cpi').textContent=data.Cp_i.toFixed(2);
      document.getElementById('fi-qz').textContent=S.R.qz.toFixed(3);
      const pDisp = (data.pMod != null && Number.isFinite(data.pMod)) ? data.pMod : data.p;
      document.getElementById('fi-p').textContent = pDisp.toFixed(3);
      const fiPCase1 = document.getElementById('fi-p-case1');
      const fiPCase2 = document.getElementById('fi-p-case2');
      const fiCpiCase2 = document.getElementById('fi-cpi-case2');
      if(fiPCase1) fiPCase1.textContent = Number.isFinite(data.p_case1) ? data.p_case1.toFixed(3) : pDisp.toFixed(3);
      if(fiPCase2) fiPCase2.textContent = Number.isFinite(data.p_case2) ? data.p_case2.toFixed(3) : '—';
      if(fiCpiCase2) fiCpiCase2.textContent = Number.isFinite(data.Cp_i_alt) ? data.Cp_i_alt.toFixed(2) : '—';
      document.getElementById('fi-area').textContent=data.area.toFixed(1);
      document.getElementById('fi-force').textContent=data.force.toFixed(1);
      document.getElementById('fi-clause').textContent='AS/NZS 1170.2 '+data.clause;
      const uploadMode = uploadedModelGroup && uploadedModelVisible && !parametricVisible;
      const fiUploadRow = document.getElementById('fi-upload-diag-row');
      const fiUploadDiag = document.getElementById('fi-upload-diag');
      const fiOverrideRow = document.getElementById('fi-override-row');
      if(uploadMode){
        lastUploadHoverUuid = obj.uuid;
        const dd = uploadClassDiag.get(obj.uuid);
        if(fiUploadRow && fiUploadDiag){
          fiUploadRow.style.display = dd ? 'block' : 'none';
          if(dd){
            const ak = dd.analyticKey || '—';
            const ovr = dd.postOverride ? ' · ' + dd.postOverride : '';
            // Phase 9: surface face-key confidence (Phase 8c output) in the
            // upload diagnostic line. Low values flag walls near the 45° tie
            // line or roofs near the flat-pitch threshold; users can review
            // and override these before pressure calc.
            const fkc = (typeof dd.faceKeyConfidence === 'number') ? dd.faceKeyConfidence : null;
            const confTag = fkc != null
              ? ' · fk=' + (fkc * 100).toFixed(0) + '%' + (fkc < 0.15 ? ' (low)' : '')
              : '';
            // Phase 13: if the mesh sits in an AS/NZS edge-distance band, show
            // the zone label (e.g. "0 to 1h") so users can see why the heatmap
            // colour differs from neighbours sharing the same face-key.
            const zoneTag = (Number.isFinite(dd.zoneIndex) && dd.zoneLabel)
              ? ' · zone ' + dd.zoneLabel
              : '';
            fiUploadDiag.textContent = dd.mlpClass + ' p=' + (dd.mlpP * 100).toFixed(0) + '% · analytic ' + ak + confTag + zoneTag + (dd.hullNear ? ' · near hull' : '') + ovr;
          }
        }
        if(fiOverrideRow){
          fiOverrideRow.style.display = 'flex';
          const sel = document.getElementById('fi-override-select');
          if(sel){
            if(!sel.dataset.cwOpts){
              sel.dataset.cwOpts = '1';
              const keys = ['windward','leeward','sidewall1','sidewall2','roof_ww','roof_lw','roof_cw','roof_hip_l','roof_hip_r'];
              sel.innerHTML = keys.map(k => '<option value="'+k+'">'+k+'</option>').join('');
            }
            if(data.key) sel.value = data.key;
          }
        }
      } else {
        lastUploadHoverUuid = null;
        if(fiUploadRow) fiUploadRow.style.display = 'none';
        if(fiOverrideRow) fiOverrideRow.style.display = 'none';
      }
      // Phase 7a: keep the retrain row in sync with the override row visibility.
      if(typeof updateRetrainUiRow === 'function') updateRetrainUiRow();
    }
  } else {
    if(faceInfoPointerInside) return;
    scheduleFaceInfoHide();
  }
}

function onPointerClick(e){
  raycaster.setFromCamera(mouse,camera);
  const rayTargets = (uploadedModelGroup && uploadedModelVisible && !parametricVisible)
    ? uploadedModelGroup.children : grpBuild.children;
  const hits=raycaster.intersectObjects(rayTargets,true).filter(i=>faceMap.has(i.object.uuid));
  if(hits.length){
    const data=faceMap.get(hits[0].object.uuid);
    if(data) toast(data.name+': p = '+data.p.toFixed(3)+' kPa, F = '+data.force.toFixed(1)+' kN');
  }
}

// ═══════════════════════════════════════════════
//   TOGGLE FUNCTIONS
// ═══════════════════════════════════════════════
function toggleFeature(key,btnId){
  S[key]=!S[key];document.getElementById(btnId).classList.toggle('active',S[key]);
  if(key==='showPressureMap'){
    const legend=document.getElementById('legend');
    if(legend) legend.style.display=S.showPressureMap?'':'none';
    if(S.analysisLocked){
      recalcPressures();
    } else {
      calc();
      rebuild();
      refreshDirectionalWindUI();
      recalcPressures();
    }
    return;
  }
  if(key==='showParticles'&&!S.showParticles&&particleSys){scene.remove(particleSys);particleSys.geometry.dispose();particleSys.material.dispose();particleSys=null;return}
  rebuild();
}
function setViewMode(m){
  S.viewMode=m;
  ['solid','transparent','wireframe'].forEach(v=>document.getElementById('btn-'+v).classList.toggle('active',v===m));
  rebuild();
}
function toggleShadows(){
  S.showShadows=!S.showShadows;document.getElementById('btn-shadows').classList.toggle('active',S.showShadows);
  renderer.shadowMap.enabled=S.showShadows;
  scene.traverse(o=>{if(o.isMesh){o.castShadow=S.showShadows;o.receiveShadow=S.showShadows}});
}
function toggleGrid(){
  S.showGrid=!S.showGrid;document.getElementById('btn-grid').classList.toggle('active',S.showGrid);
  gridHelper.visible=S.showGrid;
}
function toggleDarkMode(){
  S.darkMode=!S.darkMode;
  document.body.classList.toggle('light',!S.darkMode);
  document.getElementById('btn-darkmode').textContent=S.darkMode?'🌙 Dark':'☀️ Light';
  setSceneBg();
}
function resetCamera(){
  if(controls){
    controls.minDistance = 4;
    controls.maxDistance = 160;
  }
  if(camera){
    camera.near = 0.1;
    camera.far = 500;
    camera.updateProjectionMatrix();
  }
  camera.position.set(32,22,38);controls.target.set(0,S.height/2,0);controls.update();
}

/** Orbit limits must track model size: default minDistance 4m hides sub‑4m IFCs after scaling. */
function fitOrbitControlsToModelExtent(extentM){
  if(!controls) return;
  const e = Math.max(extentM, 1e-9);
  controls.minDistance = Math.max(0.02, Math.min(e * 0.06, 85));
  controls.maxDistance = Math.max(220, e * 35);
}

// ═══════════════════════════════════════════════
//   EXPORT FUNCTIONS — PDF UNLOCKED
// ═══════════════════════════════════════════════
function takeScreenshot(){
  renderer.render(scene,camera);
  const a=document.createElement('a');
  a.download='wind-analysis-'+Date.now()+'.png';
  a.href=renderer.domElement.toDataURL('image/png');a.click();
  toast('Screenshot saved!');
}

function exportExcel(){
  const R=S.R;
  const hStyle='style="background:#2a3a5c;color:#fff;font-weight:bold;padding:6px 12px;border:1px solid #1a2a4c;text-align:center;"';
  const pStyle='style="background:#1e2d4a;color:#8cf;font-weight:bold;padding:6px 12px;border:1px solid #1a2a4c;"';
  const cStyle='style="padding:4px 10px;border:1px solid #ccc;text-align:center;"';
  const lStyle='style="padding:4px 10px;border:1px solid #ccc;text-align:left;font-weight:bold;background:#f0f4f8;"';
  const vStyle='style="padding:4px 10px;border:1px solid #ccc;text-align:right;"';

  let html='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">';
  html+='<head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>';
  html+='<x:ExcelWorksheet><x:Name>Wind Pressures</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>';
  html+='<x:ExcelWorksheet><x:Name>Parameters</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>';
  html+='<x:ExcelWorksheet><x:Name>Direction Multipliers</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>';
  html+='</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>';

  // Sheet 1: Wind Pressures
  html+='<table border="1" cellpadding="4" cellspacing="0">';
  html+=`<tr><td colspan="8" style="background:#1a2a4c;color:#fff;font-size:16px;font-weight:bold;padding:10px;text-align:center;">Face Pressures (AS/NZS 1170.2:2021)</td></tr>`;
  html+=`<tr><td ${hStyle}>Face</td><td ${hStyle}>Cp,e</td><td ${hStyle}>Cp,i</td><td ${hStyle}>qz (kPa)</td><td ${hStyle}>p (kPa)</td><td ${hStyle}>Area (m²)</td><td ${hStyle}>Force (kN)</td><td ${hStyle}>Clause</td></tr>`;
  for(let k in R.faces){
    const f=R.faces[k];
    html+=`<tr><td ${lStyle}>${f.name}</td><td ${cStyle}>${f.Cp_e.toFixed(3)}</td><td ${cStyle}>${f.Cp_i.toFixed(3)}</td><td ${cStyle}>${R.qz.toFixed(3)}</td><td ${cStyle}>${f.p.toFixed(3)}</td><td ${cStyle}>${f.area.toFixed(1)}</td><td ${cStyle}>${f.force.toFixed(1)}</td><td ${cStyle}>AS/NZS 1170.2 ${f.clause}</td></tr>`;
  }
  html+='</table>';

  // Sheet 2: Parameters
  html+='<br><table border="1" cellpadding="4" cellspacing="0">';
  html+=`<tr><td colspan="2" style="background:#1a2a4c;color:#fff;font-size:16px;font-weight:bold;padding:10px;text-align:center;">Site & Building Parameters</td></tr>`;
  html+=`<tr><td ${pStyle}>Parameter</td><td ${pStyle}>Value</td></tr>`;
  const params=[
    ['Regional Wind Speed (V_R)',R.VR+' m/s'],
    ['Site Wind Speed (V_sit,β)',R.Vsit.toFixed(1)+' m/s'],
    ['Terrain/Height Multiplier (Mz,cat)',R.Mz.toFixed(3)],
    ['Design Wind Pressure (qz)',R.qz.toFixed(3)+' kPa'],
    ['Building Width',S.width+' m'],
    ['Building Depth',S.depth+' m'],
    ['Building Height',S.height+' m'],
    ['Roof Pitch',S.pitch+'°'],
    ['Terrain Category','TC'+S.terrainCat],
    ['Wind Region',S.region],
    ['Latitude',S.lat.toFixed(4)],
    ['Longitude',S.lng.toFixed(4)],
    ['Importance Level',String(S.importance)],
    ['Design life',(S.life != null ? S.life : 50)+' years'],
    ['Design ARI (ULS wind)',(S.ari != null ? S.ari : '-')+' years']
  ];
  params.forEach(([label,val])=>{
    html+=`<tr><td ${lStyle}>${label}</td><td ${vStyle}>${val}</td></tr>`;
  });
  html+='</table>';

  // Sheet 3: 8-Direction Multipliers
  html+='<br><table border="1" cellpadding="4" cellspacing="0">';
  html+=`<tr><td colspan="6" style="background:#1a2a4c;color:#fff;font-size:16px;font-weight:bold;padding:10px;text-align:center;">Directional Multipliers</td></tr>`;
  html+=`<tr><td ${hStyle}>Direction</td><td ${hStyle}>Md</td><td ${hStyle}>Mz,cat</td><td ${hStyle}>Ms</td><td ${hStyle}>Mt</td><td ${hStyle}>V_sit,β (m/s)</td></tr>`;
  for(let i=0;i<8;i++){
    html+=`<tr><td ${lStyle}>${S.dirs[i]}</td><td ${cStyle}>${effectiveMd(i).toFixed(2)}</td><td ${cStyle}>${formatMzcatForReport(i)}</td><td ${cStyle}>${S.Ms[i]}</td><td ${cStyle}>${S.Mt[i]}</td><td ${cStyle}>${S.Vsit_dir[i].toFixed(1)}</td></tr>`;
  }
  html+='</table>';

  html+='</body></html>';
  const blob=new Blob([html],{type:'application/vnd.ms-excel'});
  const a=document.createElement('a');a.download='wind-analysis-'+Date.now()+'.xls';a.href=URL.createObjectURL(blob);a.click();
  toast('Excel file exported!');
}

function exportPDF(){
  // Temporary: bypass PDF plan gate while report changes are in progress.
  generatePDF();
}

/** Table 5.6 wall Kl strips — same zones as buildKlWalls; PDF export only. */
function collectWallKlPdfRowsForPdf(){
  const w=S.width, d=S.depth, h=S.height;
  if(!S.R||!S.R.faces) return [];
  const fm=getWindFaceMap();
  const a = Math.min(0.2 * Math.min(w, d), h);
  const rB = h / Math.min(w, d);
  const isCladding = S.showPressureMap;
  const isGlazing = false;
  const wwZones = (isCladding || isGlazing)
    ? [{dist:Infinity, Kl:1.5, name:'WA1'}]
    : [{dist:Infinity, Kl:1.0, name:'Other'}];
  const lwZones = [{dist:Infinity, Kl:1.0, name:'Other'}];
  let swZones;
  if(rB <= 1){
    swZones = (isGlazing || isCladding)
      ? [{dist:0.5*a, Kl:2.0, name:'SA2'},{dist:a, Kl:1.5, name:'SA1'},{dist:Infinity, Kl:1.0, name:'Other'}]
      : [{dist:Infinity, Kl:1.0, name:'Other'}];
  } else {
    swZones = (isGlazing || isCladding)
      ? [{dist:0.5*a, Kl:3.0, name:'SA5'},{dist:a, Kl:2.0, name:'SA4'},{dist:Infinity, Kl:1.5, name:'SA3'}]
      : [{dist:Infinity, Kl:1.0, name:'Other'}];
  }
  const phys = [
    {key:'front', lbl:'Front (+Z)', wallW:w, zStrip:false, highNeighbor:'right'},
    {key:'back', lbl:'Back (-Z)', wallW:w, zStrip:false, highNeighbor:'right'},
    {key:'left', lbl:'Left (-X)', wallW:d, zStrip:true, highNeighbor:'front'},
    {key:'right', lbl:'Right (+X)', wallW:d, zStrip:true, highNeighbor:'front'},
  ];
  const rows = [];
  for(const pf of phys){
    const role = fm[pf.key];
    const faceData = S.R.faces[role];
    if(!faceData) continue;

    if(role === 'windward' || role === 'leeward'){
      const zones = role === 'windward' ? wwZones : lwZones;
      if(S.roofType === 'monoslope' && (pf.key === 'left' || pf.key === 'right')){
        const { pd, monoRise: mr } = getMonoWallParams();
        const strips = mr > 0.001 ? getKlStrips(pd, zones) : getKlStrips(pf.wallW, zones);
        strips.forEach(strip=>{
          if(strip.x2-strip.x1<1e-6) return;
          const klP = localPnet(faceData, strip.Kl, false);
          const zn = strip.name || klZoneName(role, strip.Kl);
          rows.push([pf.lbl, faceData.name, zn, strip.Kl.toFixed(1), faceData.Cp_e.toFixed(2), klP.toFixed(3)]);
        });
      } else {
        const strips = getKlStrips(pf.wallW, zones);
        strips.forEach(strip=>{
          if(strip.x2-strip.x1<0.01) return;
          const klP = localPnet(faceData, strip.Kl, false);
          const zn = strip.name || klZoneName(role, strip.Kl);
          rows.push([pf.lbl, faceData.name, zn, strip.Kl.toFixed(1), faceData.Cp_e.toFixed(2), klP.toFixed(3)]);
        });
      }
    } else {
      const fromHighEnd = (fm[pf.highNeighbor] === 'windward');
      if(S.roofType === 'monoslope' && (pf.key === 'left' || pf.key === 'right')){
        const { pd } = getMonoWallParams();
        const strips = getKlStripsOneSided(pd, swZones, fromHighEnd);
        strips.forEach(strip=>{
          if(strip.x2-strip.x1<1e-6) return;
          const klP = localPnet(faceData, strip.Kl, true);
          const zn = 'WA1 & ' + (strip.name || 'Other');
          rows.push([pf.lbl, faceData.name, zn, strip.Kl.toFixed(1), faceData.Cp_e.toFixed(2), klP.toFixed(3)]);
        });
      } else {
        const strips = getKlStripsOneSided(pf.wallW, swZones, fromHighEnd);
        strips.forEach(strip=>{
          if(strip.x2-strip.x1<0.01) return;
          const klP = localPnet(faceData, strip.Kl, true);
          const zn = strip.name || 'Other';
          rows.push([pf.lbl, faceData.name, zn, strip.Kl.toFixed(1), faceData.Cp_e.toFixed(2), klP.toFixed(3)]);
        });
      }
    }
  }
  return rows;
}

/** Table 5.6 roof Kl — corner / edge / interior; matches buildKlGableRoof factors for cladding. */
function collectRoofKlPdfRowsForPdf(){
  if(!S.R||!S.R.faces) return [];
  const isCladding = S.showPressureMap;
  let cornerKl, edgeKl, interiorKl;
  if(isCladding){
    cornerKl = 3.0; edgeKl = 1.5; interiorKl = 1.0;
  } else {
    cornerKl = edgeKl = interiorKl = 1.0;
  }
  const pitchDeg = S.pitch || 0;
  const useRC2 = pitchDeg >= 10;
  const rows = [];
  const add = (faceKey, zoneLabel, kl)=>{
    const f = S.R.faces[faceKey];
    if(!f) return;
    const p = localPnet(f, kl, true);
    rows.push([zoneLabel, f.name, kl.toFixed(1), f.Cp_e.toFixed(2), p.toFixed(3)]);
  };
  const rcEave = useRC2 ? 'RC1 eave corner' : 'RC1 corner';
  const rcRidge = 'RC2 ridge corner';
  add('roof_ww', rcEave, cornerKl);
  if(useRC2) add('roof_ww', rcRidge, cornerKl);
  add('roof_ww', 'RA edge / ridge band', edgeKl);
  add('roof_ww', 'MR interior', interiorKl);
  add('roof_lw', rcEave, cornerKl);
  if(useRC2) add('roof_lw', rcRidge, cornerKl);
  add('roof_lw', 'RA edge / ridge band', edgeKl);
  add('roof_lw', 'MR interior', interiorKl);
  if(S.roofType === 'hip'){
    if(S.R.faces.roof_hip_l){
      add('roof_hip_l', 'Hip end RC corner', cornerKl);
      add('roof_hip_l', 'Hip end MR interior', interiorKl);
    }
    if(S.R.faces.roof_hip_r){
      add('roof_hip_r', 'Hip end RC corner', cornerKl);
      add('roof_hip_r', 'Hip end MR interior', interiorKl);
    }
  }
  return rows;
}

/**
 * Phase 23: scrape the rendered LPZ HTML tables (local-walls-body /
 * local-roof-body) into a 2D array of strings. Each TD's textContent gives the
 * already-formatted value (Surface, Distance, Ref, Area, K_l, K_c,i, K_c,e,
 * K_p, V_sit, q_z, C_shp,e, p_e, p_net(+), p_net(−)). The DOM is the source of
 * truth for these tables — re-running recalcPressures() for a different
 * direction repopulates them, so we snapshot once per direction.
 */
function snapshotLpzTableRows(bodyId){
  if(typeof document === 'undefined') return [];
  const body = document.getElementById(bodyId);
  if(!body) return [];
  const rows = [];
  body.querySelectorAll('tr').forEach(tr => {
    const cells = Array.from(tr.querySelectorAll('td')).map(td => (td.textContent || '').trim());
    if(cells.length) rows.push(cells);
  });
  return rows;
}

/**
 * Phase 22: pure helper. Given one direction's captured face data plus the
 * internal pressure cases for that direction, return the per-face rows we want
 * to draw in the CHECKWIND-style per-direction detail tables. Pure: no DOM
 * access, no globals — easy to unit test.
 *
 * Returned rows: [name, Cp_e, Cp_i (design), p_e, p_net(+), p_net(-), p_design, area, force]
 */
function buildDirectionFaceRows(facesEntry, faceKeyOrder, qz, cpi1, cpi2){
  const out = [];
  if(!facesEntry || !Array.isArray(faceKeyOrder)) return out;
  const Q = Number.isFinite(qz) ? qz : 0;
  const C1 = Number.isFinite(cpi1) ? cpi1 : 0;
  const C2 = Number.isFinite(cpi2) ? cpi2 : 0;
  for(const k of faceKeyOrder){
    const f = facesEntry[k];
    if(!f) continue;
    const cpe = Number.isFinite(f.Cp_e) ? f.Cp_e : 0;
    const cpiD = Number.isFinite(f.Cp_i) ? f.Cp_i : 0;
    const pe = (Number.isFinite(f.peExternalAvg) ? f.peExternalAvg : (Q * cpe));
    const pNet1 = pe - Q * C1;
    const pNet2 = pe - Q * C2;
    const pDes = Number.isFinite(f.p) ? f.p : 0;
    const area = Number.isFinite(f.area) ? f.area : 0;
    const force = Number.isFinite(f.force) ? f.force : 0;
    out.push([
      f.name || k,
      cpe.toFixed(2),
      cpiD.toFixed(2),
      pe.toFixed(3),
      pNet1.toFixed(3),
      pNet2.toFixed(3),
      pDes.toFixed(3),
      area.toFixed(1),
      force.toFixed(1)
    ]);
  }
  return out;
}

async function generatePDF(options = {}){
  const saveFile = options.save !== false;
  const returnBlob = options.returnBlob === true;
  const silent = options.silent === true;
  if(!S.R||!S.R.faces){toast('Run a calculation first');return}

  const savedWindAngle = S.windAngle;
  const faceKeyOrder = ['windward','leeward','sidewall1','sidewall2','roof_ww','roof_lw','roof_cw'];
  if(S.R.faces.roof_hip_l) faceKeyOrder.push('roof_hip_l','roof_hip_r');
  // Phase 22: capture per-direction multipliers, internal pressure cases, and a
  // full Cp_e/Cp_i/area/force breakdown for every face. The original capture
  // only retained {p, name}; we now retain enough to render a CHECKWIND-style
  // per-direction full pressure analysis later in Section 17.
  const cardinalData = [];
  for(let i=0;i<8;i++){
    S.windAngle = i * 45;
    calc();
    // Phase 23: drive the Local Pressure Zone tables for this direction so we
    // can DOM-scrape them. recalcPressures() reads S.windAngle and rewrites
    // local-walls-body / local-roof-body in place.
    if(typeof recalcPressures === 'function') {
      try { recalcPressures(); } catch(e){ console.warn('recalcPressures failed for dir '+i, e); }
    }
    const F = S.R.faces;
    const cpiCases = (typeof getCpiCasesForDesign === 'function') ? getCpiCasesForDesign() : { cpi1: 0, cpi2: 0, clause: '' };
    const entry = {
      dir: S.dirs[i], angle: i * 45,
      qz: S.R.qz,
      Vsit: S.Vsit_dir[i],
      Md: (typeof effectiveMd === 'function') ? effectiveMd(i) : (S.Md ? S.Md[i] : 1),
      MzcatStr: (typeof formatMzcatForReport === 'function') ? formatMzcatForReport(i) : '',
      Ms: S.Ms[i],
      Mt: S.Mt[i],
      cpi1: cpiCases.cpi1,
      cpi2: cpiCases.cpi2,
      cpiClause: cpiCases.clause,
      faces: {},
      // Phase 23: snapshot the rendered LPZ tables for this direction so we
      // can render an exact replica per direction in Section 17b.
      lpzWalls: snapshotLpzTableRows('local-walls-body'),
      lpzRoof:  snapshotLpzTableRows('local-roof-body')
    };
    faceKeyOrder.forEach(k => {
      if(F[k]){
        entry.faces[k] = {
          name: F[k].name,
          p: F[k].p,
          Cp_e: F[k].Cp_e,
          Cp_i: F[k].Cp_i,
          area: F[k].area,
          force: F[k].force,
          clause: F[k].clause,
          peExternalAvg: F[k].peExternalAvg
        };
      }
    });
    cardinalData.push(entry);
  }
  S.windAngle = savedWindAngle;
  calc();
  if(typeof recalcPressures === 'function') {
    try { recalcPressures(); } catch(e){ console.warn('recalcPressures restore failed', e); }
  }
  const R=S.R;
  const wallKlPdfRows = collectWallKlPdfRowsForPdf();
  const roofKlPdfRows = collectRoofKlPdfRowsForPdf();

  const{jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'portrait', unit:'mm', format:'a4'});

  // ═══════════════════════════════════════════════
  //   PDF CONSTANTS & HELPERS
  // ═══════════════════════════════════════════════
  const PW=210, PH=297, ML=15, MR=15, MT=20, MB=25;
  const CW=PW-ML-MR; // content width
  let pg=1, y=MT;

  // Colours
  const COL_PRIMARY=[26,82,118];   // dark blue
  const COL_ACCENT=[41,128,185];   // medium blue
  const COL_LIGHT_BG=[235,245,251];// light blue bg
  const COL_TEXT=[40,40,40];
  const COL_MUTED=[120,120,120];
  const COL_GREEN=[39,174,96];
  const COL_RED=[192,57,43];
  const COL_CLAUSE=[100,100,180];

  // Sanitize Unicode chars that jsPDF standard fonts cannot render.
  // Latin-1 Supplement (U+00A0-00FF) and WinAnsi specials are preserved
  // because jsPDF renders them natively (°, ², ³, ×, ±, —, –, etc).
  function pdfSafe(s){
    if(typeof s !== 'string') return String(s);
    return s
      // Greek letters (NOT in WinAnsi/Latin-1 — must replace)
      .replace(/\u03B8/g, 'th')
      .replace(/\u03B2/g, 'B')
      .replace(/\u0394H/g, 'dH')
      .replace(/\u0394/g, 'D')
      .replace(/\u03C1/g, 'rho')
      .replace(/\u03B1/g, 'a')
      // Math symbols NOT in WinAnsi
      .replace(/\u2264/g, '<=')
      .replace(/\u2265/g, '>=')
      .replace(/\u2248/g, '~=')
      .replace(/\u2212/g, '-')
      // Strip remaining chars outside jsPDF renderable range
      // Keep: ASCII printable, Latin-1 Supplement, WinAnsi specials, whitespace
      .replace(/[^\x20-\x7E\xA0-\xFF\u2013\u2014\u2018-\u201D\u2022\u2026\n\r\t]/g, '');
  }

  // Monkey-patch doc.text and doc.splitTextToSize so ALL text is sanitized,
  // including cover page, TOC, notes, and references that bypass helpers.
  const _origDocText = doc.text.bind(doc);
  doc.text = function(text, ...rest) {
    if (typeof text === 'string') text = pdfSafe(text);
    else if (Array.isArray(text)) text = text.map(t => typeof t === 'string' ? pdfSafe(t) : t);
    return _origDocText(text, ...rest);
  };
  const _origSplitText = doc.splitTextToSize.bind(doc);
  doc.splitTextToSize = function(text, ...rest) {
    if (typeof text === 'string') text = pdfSafe(text);
    return _origSplitText(text, ...rest);
  };

  function checkPage(need){
    if(y+need>PH-MB){doc.addPage();pg++;y=MT;addHeader();return true}
    return false;
  }
  function addHeader(){
    doc.setFillColor(...COL_PRIMARY);
    doc.rect(0,0,PW,12,'F');
    doc.setFontSize(7);doc.setTextColor(255,255,255);
    doc.text('WIND LOAD ANALYSIS -- AS/NZS 1170.2:2021',ML,8);
    doc.text('Wind Analysis',PW-MR-30,8);
    // Footer
    doc.setFillColor(240,240,240);
    doc.rect(0,PH-15,PW,15,'F');
    doc.setFontSize(7);doc.setTextColor(...COL_MUTED);
    doc.text('Generated: '+new Date().toLocaleString(),ML,PH-7);
    doc.text('Page '+pg,PW-MR-10,PH-7);
    doc.text('CONFIDENTIAL',PW/2-12,PH-7);
    y=MT+2;
  }
  // Polish: capture the real page number every time a numbered section starts
  // so the Table of Contents can be rewritten with accurate page references
  // after the body has been laid out.
  const sectionPages = {};
  function sectionTitle(title,clauseRef){
    checkPage(20);
    title=pdfSafe(title); clauseRef=clauseRef?pdfSafe(clauseRef):'';
    // Record the page where this section starts. Match the leading "N." or
    // "Nb." (e.g. "1.", "17b.") and store the current page.
    const m = title.match(/^\s*([0-9]+[a-zA-Z]?)\./);
    if(m && !sectionPages[m[1]]) sectionPages[m[1]] = pg;
    doc.setFillColor(...COL_PRIMARY);
    doc.rect(ML,y,CW,8,'F');
    doc.setFontSize(10);doc.setTextColor(255,255,255);doc.setFont(undefined,'bold');
    doc.text(title,ML+3,y+5.5);
    if(clauseRef){
      doc.setFontSize(7);
      doc.text(clauseRef,ML+CW-doc.getTextWidth(clauseRef)-3,y+5.5);
    }
    doc.setFont(undefined,'normal');
    y+=12;
  }
  function subTitle(title,clauseRef){
    checkPage(14);
    title=pdfSafe(title); clauseRef=clauseRef?pdfSafe(clauseRef):'';
    doc.setFillColor(...COL_LIGHT_BG);
    doc.rect(ML,y,CW,7,'F');
    doc.setDrawColor(...COL_ACCENT);
    doc.line(ML,y+7,ML+CW,y+7);
    doc.setFontSize(9);doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
    doc.text(title,ML+3,y+5);
    if(clauseRef){
      doc.setFontSize(7);doc.setTextColor(...COL_CLAUSE);doc.setFont(undefined,'italic');
      doc.text(clauseRef,ML+CW-doc.getTextWidth(clauseRef)-3,y+5);
    }
    doc.setFont(undefined,'normal');doc.setTextColor(...COL_TEXT);
    y+=10;
  }
  function clauseNote(text){
    text=pdfSafe(text);
    doc.setFontSize(7);doc.setTextColor(...COL_CLAUSE);doc.setFont(undefined,'italic');
    doc.text(text,ML+3,y);
    doc.setFont(undefined,'normal');doc.setTextColor(...COL_TEXT);
    y+=4;
  }
  function paramRow(label,value,clause,indent){
    checkPage(6);
    label=pdfSafe(label); value=pdfSafe(String(value)); clause=clause?pdfSafe(clause):'';
    const x=ML+(indent||3);
    doc.setFontSize(8.5);doc.setTextColor(...COL_TEXT);
    doc.text(label,x,y);
    doc.setFont(undefined,'bold');
    doc.text(value,ML+90,y);
    doc.setFont(undefined,'normal');
    if(clause){
      doc.setFontSize(6.5);doc.setTextColor(...COL_CLAUSE);doc.setFont(undefined,'italic');
      doc.text(clause,ML+135,y);
      doc.setFont(undefined,'normal');doc.setTextColor(...COL_TEXT);
    }
    y+=5.5;
  }
  function calcLine(formula,result,clause){
    checkPage(7);
    formula=pdfSafe(formula); result=pdfSafe(result); clause=clause?pdfSafe(clause):'';
    doc.setFontSize(8);doc.setTextColor(...COL_MUTED);
    doc.text(formula,ML+6,y);
    doc.setFontSize(9);doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
    doc.text('= '+result,ML+120,y);
    doc.setFont(undefined,'normal');
    if(clause){
      doc.setFontSize(6.5);doc.setTextColor(...COL_CLAUSE);doc.setFont(undefined,'italic');
      doc.text(clause,ML+155,y);
      doc.setFont(undefined,'normal');
    }
    y+=6;
  }
  function drawTable(headers,rows,colWidths,opts){
    opts=opts||{};
    // Sanitize all text for jsPDF font compatibility
    headers=headers.map(h=>pdfSafe(h));
    rows=rows.map(r=>r.map(c=>pdfSafe(c)));
    checkPage(10+rows.length*6);
    doc.autoTable({
      startY:y,
      margin:{left:ML,right:MR},
      head:[headers],
      body:rows,
      columnStyles: colWidths ? Object.fromEntries(colWidths.map((w,i)=>
        [i,{cellWidth:w}])) : {},
      styles:{fontSize:7.5,cellPadding:1.5,lineColor:[200,200,200],lineWidth:0.2,textColor:COL_TEXT},
      headStyles:{fillColor:COL_PRIMARY,textColor:[255,255,255],fontStyle:'bold',fontSize:7.5},
      alternateRowStyles:{fillColor:[248,250,252]},
      didParseCell:function(data){
        if(opts.highlightCol!==undefined && data.column.index===opts.highlightCol && data.section==='body'){
          data.cell.styles.fontStyle='bold';
          data.cell.styles.textColor=COL_PRIMARY;
        }
      },
      theme:'grid'
    });
    y=doc.lastAutoTable.finalY+4;
  }

  // ═══════════════════════════════════════════════
  //   PAGE 1 — COVER PAGE
  // ═══════════════════════════════════════════════
  // Blue header band
  doc.setFillColor(...COL_PRIMARY);
  doc.rect(0,0,PW,85,'F');

  // Diagonal accent
  doc.setFillColor(...COL_ACCENT);
  doc.triangle(0,75,PW,60,PW,85,'F');

  doc.setFontSize(32);doc.setTextColor(255,255,255);doc.setFont(undefined,'bold');
  doc.text('WIND LOAD',ML,38);
  doc.text('ANALYSIS REPORT',ML,52);
  doc.setFontSize(12);doc.setFont(undefined,'normal');
  doc.text('AS/NZS 1170.2:2021 Structural design actions — Wind actions',ML,68);

  // Project info box
  y=100;
  doc.setFillColor(248,250,252);
  doc.roundedRect(ML,y-5,CW,60,3,3,'F');
  doc.setDrawColor(...COL_ACCENT);
  doc.roundedRect(ML,y-5,CW,60,3,3,'S');

  doc.setFontSize(10);doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
  doc.text('PROJECT DETAILS',ML+5,y+3);y+=10;
  doc.setFont(undefined,'normal');doc.setFontSize(9);doc.setTextColor(...COL_TEXT);

  // Polish: hemisphere-aware coords (S.lat is signed; tag N/S and E/W from sign)
  // and fix the Design Life precedence bug — `?.value+' years' || '50 years'`
  // evaluated as `('undefined years') || '50 years'`, so the fallback never fired.
  const _latSign = (S.lat < 0) ? 'S' : 'N';
  const _lngSign = (S.lng < 0) ? 'W' : 'E';
  const _life = (document.getElementById('inp-life')?.value) || '50';
  const coverFields=[
    ['Location',S.address||S.lat.toFixed(5)+', '+S.lng.toFixed(5)],
    ['Coordinates',Math.abs(S.lat).toFixed(6)+'° '+_latSign+', '+Math.abs(S.lng).toFixed(6)+'° '+_lngSign],
    ['Elevation',S.elevation||'N/A'],
    ['Wind Region',S.region],
    ['Importance Level',S.importance],
    ['Design Life', _life + ' years'],
    ['Standard','AS/NZS 1170.2:2021'],
  ];
  coverFields.forEach(([k,v])=>{
    doc.setTextColor(...COL_MUTED);doc.text(k+':',ML+8,y);
    doc.setTextColor(...COL_TEXT);doc.setFont(undefined,'bold');
    doc.text(String(v),ML+55,y);doc.setFont(undefined,'normal');
    y+=6.5;
  });

  // Building quick summary
  y=175;
  doc.setFillColor(248,250,252);
  doc.roundedRect(ML,y-5,CW/2-2,40,3,3,'F');
  doc.setDrawColor(...COL_ACCENT);
  doc.roundedRect(ML,y-5,CW/2-2,40,3,3,'S');

  doc.setFontSize(9);doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
  doc.text('BUILDING',ML+5,y+3);y+=8;doc.setFont(undefined,'normal');
  doc.setTextColor(...COL_TEXT);doc.setFontSize(8);
  [['Width',S.width+' m'],['Depth',S.depth+' m'],['Eave Height',S.height+' m'],
   ['Roof',S.roofType+' @ '+S.pitch+'°']].forEach(([k,v])=>{
    doc.text(k+': '+v,ML+8,y);y+=5.5;
  });

  y=175;
  const x2=ML+CW/2+2;
  doc.setFillColor(248,250,252);
  doc.roundedRect(x2,y-5,CW/2-2,40,3,3,'F');
  doc.setDrawColor(...COL_ACCENT);
  doc.roundedRect(x2,y-5,CW/2-2,40,3,3,'S');

  doc.setFontSize(9);doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
  doc.text('WIND SUMMARY',x2+5,y+3);y+=8;doc.setFont(undefined,'normal');
  doc.setTextColor(...COL_TEXT);doc.setFontSize(8);
  [['Region V_R',R.VR+' m/s'],['Site V_sit',R.Vsit.toFixed(1)+' m/s'],
   ['q_z',R.qz.toFixed(3)+' kPa'],['Terrain','TC'+S.terrainCat]].forEach(([k,v])=>{
    doc.text(k+': '+v,x2+5,y);y+=5.5;
  });

  // Disclaimer at bottom
  doc.setFontSize(7);doc.setTextColor(...COL_MUTED);
  doc.text('This report was generated by this software. All calculations are in accordance with',ML,PH-35);
  doc.text('AS/NZS 1170.2:2021. The engineer of record must verify all inputs, assumptions and results.',ML,PH-30);
  doc.text('Report Date: '+new Date().toLocaleDateString()+' | '+new Date().toLocaleTimeString(),ML,PH-22);

  // ═══════════════════════════════════════════════
  //   PAGE 2 — TABLE OF CONTENTS
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=5;

  // Polish: dynamic TOC. We render once with placeholder page numbers using
  // sections (number + title), then re-render this same area at the end of
  // generatePDF once `sectionPages` has been populated by every sectionTitle()
  // call. The initial pass reserves the page so the body lays out predictably.
  const tocSections = [
    ['1','Site Information & Location'],
    ['2','3D Visualization'],
    ['3','Regional Wind Speed (V_R)'],
    ['4','Wind Direction Multiplier (M_d)'],
    ['5','Terrain/Height Multiplier (M_z,cat)'],
    ['6','Shielding Multiplier (M_s)'],
    ['7','Topographic Multiplier (M_t)'],
    ['8','Site Wind Speed (V_sit,β)'],
    ['9','Design Wind Speed (V_des,θ)'],
    ['10','Dynamic Response Factor (C_dyn)'],
    ['11','Aerodynamic Shape Factor (C_fig)'],
    ['12','External Pressure Coefficients — Walls'],
    ['13','External Pressure Coefficients — Roof'],
    ['14','Internal Pressure Coefficients'],
    ['15','Local Pressure Factors (K_l)'],
    ['16','Design Wind Pressures Summary'],
    ['17','Cardinal Directions — Wind Speeds & Pressures'],
    ['17b','Per-Direction Detailed Pressures'],
    ['18','Topographic Long Sections (Cardinal)'],
    ['19','Notes & Disclaimers'],
  ];
  const TOC_PAGE_NUMBER = pg; // capture the page that holds the TOC
  function drawToc(yStart){
    let yy = yStart;
    doc.setFontSize(16);doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
    doc.text('TABLE OF CONTENTS',ML,yy);yy+=12;doc.setFont(undefined,'normal');
    tocSections.forEach(([num,title])=>{
      const pageRef = sectionPages[num] != null ? String(sectionPages[num]) : '—';
      doc.setFontSize(9);
      doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
      doc.text(num+'.',ML+3,yy);
      doc.setTextColor(...COL_TEXT);doc.setFont(undefined,'normal');
      doc.text(title,ML+12,yy);
      // dotted leader between title and page number
      doc.setTextColor(...COL_MUTED);
      const titleW = doc.getTextWidth(title);
      const dotsX0 = ML + 12 + titleW + 2;
      const dotsX1 = ML + CW - 8;
      if(dotsX1 > dotsX0){
        const w = doc.getTextWidth('. ');
        const n = Math.max(0, Math.floor((dotsX1 - dotsX0)/w));
        if(n > 0) doc.text('. '.repeat(n), dotsX0, yy);
      }
      doc.text(pageRef,ML+CW-5,yy,{align:'right'});
      yy+=7;
    });
    return yy;
  }
  // First-pass: placeholders. Will be redrawn at the end with real page numbers.
  drawToc(y);

  // ═══════════════════════════════════════════════
  //   PAGE 3 — SITE INFORMATION
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('1. SITE INFORMATION','AS/NZS 1170.2:2021');
  paramRow('Project Address',S.address||'Not specified');
  paramRow('Latitude',S.lat.toFixed(6)+'°');
  paramRow('Longitude',S.lng.toFixed(6)+'°');
  paramRow('Site Elevation',S.elevation||'N/A');
  paramRow('Wind Region',S.region,'Clause 3.2, Table 3.1');
  paramRow('Climate Change','Not applicable','Clause 3.2');
  paramRow('Importance Level',S.importance,'AS/NZS 1170.0 Table 3.1');
  paramRow('Design Life',
    (document.getElementById('inp-life')?.value||'50')+' years',
    'AS/NZS 1170.0 Clause 3.2');
  paramRow('Annual probability of exceedance (ULS wind)','1/'+S.ari,'AS/NZS 1170.0 Table 3.3');
  y+=3;

  // --- Map capture embedded in Site Information ---
  try {
    const mapEl = document.getElementById('leaflet-map');
    if(mapEl && typeof html2canvas !== 'undefined'){
      const mapCanvas = await html2canvas(mapEl, {useCORS:true, logging:false, scale:1.5, allowTaint:true});
      const mapImg = mapCanvas.toDataURL('image/png');
      checkPage(75);
      subTitle('Site Map','');
      doc.addImage(mapImg,'PNG',ML,y,CW,CW*0.5);
      y += CW*0.5 + 4;
      clauseNote('Map showing site location (pin) and surrounding terrain context.');
      y += 2;
    }
  } catch(e){
    console.warn('Map capture failed for PDF:', e);
    clauseNote('[Map capture unavailable — CORS or rendering issue]');
    y += 3;
  }

  // ═══════════════════════════════════════════════
  //   3D VISUALIZATION
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('2. 3D VISUALIZATION','');
  y+=2;
  try{
    renderer.render(scene,camera);
    const img=renderer.domElement.toDataURL('image/png');
    if(y+90>PH-MB){doc.addPage();pg++;addHeader();y+=3}
    doc.addImage(img,'PNG',ML,y,CW,CW*0.5625);
    y+=CW*0.5625+5;
  }catch(e){
    doc.setFontSize(8);doc.setTextColor(...COL_MUTED);
    doc.text('[3D visualization capture failed]',ML+3,y);y+=8;
  }
  clauseNote('3D model showing building geometry, wind direction, and pressure distribution (heatmap).');
  y+=4;

  doc.addPage();pg++;
  addHeader();y+=3;

  // ═══════════════════════════════════════════════
  //   REGIONAL WIND SPEED
  // ═══════════════════════════════════════════════
  sectionTitle('3. REGIONAL WIND SPEED (V_R)','Clause 3.2');
  clauseNote('The regional 3-second gust wind speed for the given region and return period.');
  y+=2;
  paramRow('Wind Region',S.region,'Table 3.1');
  paramRow('Average recurrence interval (design, ULS wind)',S.ari+' years','AS/NZS 1170.0 Table 3.3');
  paramRow('Regional wind speed V_R',R.VR+' m/s','AS/NZS 1170.2 Table 3.1'+(S.vrManual?' (manual override)':''));
  if(S.svcVr){
    paramRow('Serviceability V_R (SLS)',S.svcVr+' m/s','Table 3.1');
  }
  y+=2;
  clauseNote('V_R is the regional 3-s gust wind speed for the selected region and return period (Clause 3.2).');
  clauseNote('Return period follows AS/NZS 1170.0 Table 3.3 from importance level and design life; V_R from AS/NZS 1170.2 Table 3.1 (interpolated between tabulated R where needed).');
  y+=4;

  // ═══════════════════════════════════════════════
  //   WIND DIRECTION MULTIPLIER
  // ═══════════════════════════════════════════════
  sectionTitle('4. WIND DIRECTION MULTIPLIER (M_d)','Clause 3.3');
  clauseNote('M_d accounts for the directional probability of strong winds for each cardinal direction.');
  clauseNote('For non-cyclonic regions: M_d values per Table 3.2. For cyclonic: M_d = 1.0 for all directions.');
  y+=2;

  const dirAngles=[0,45,90,135,180,225,270,315];
  const dir8Full = ['North','North East','East','South East','South','South West','West','North West'];
  const dir16Short = ['N','NNE','NE','ENE','E','ESE','SE','SSE','S','SSW','SW','WSW','W','WNW','NW','NNW'];
  const dir16Full = ['North','North North East','North East','East North East','East','East South East','South East','South South East','South','South South West','South West','West South West','West','West North West','North West','North North West'];
  const normDeg = (deg) => ((deg % 360) + 360) % 360;
  const dirLongLabel = (idx) => `${S.dirs[idx]} (${dir8Full[idx]})`;
  const windwardFaceAngle = (idx) => normDeg(dirAngles[idx] - 22.5);
  const windwardFaceIdx = (idx) => Math.round(windwardFaceAngle(idx) / 22.5) % 16;
  const windwardFaceShortLabel = (idx) => `${dir16Short[windwardFaceIdx(idx)]} (${windwardFaceAngle(idx).toFixed(1)}°)`;
  const windwardFaceFullLabel = (idx) => `${dir16Full[windwardFaceIdx(idx)]} (θ = ${windwardFaceAngle(idx).toFixed(1)}°)`;
  const mdRows=S.dirs.map((d,i)=>[dirLongLabel(i),dirAngles[i]+'°',effectiveMd(i).toFixed(2)]);
  drawTable(
    ['Direction','Angle (θ)','M_d'],
    mdRows,
    [30,40,30],
    {highlightCol:2}
  );
  clauseNote('Reference: Table 3.2 — Wind direction multipliers for Region '+S.region);
  if(mdOverrideKlDesign()){
    clauseNote('M_d = 1.0 with Local Pressure Zones on for cladding design in Regions B2, C, and D (Clause 3.3(b)): Table 3.2 does not apply.');
  }
  y+=2;

  // ═══════════════════════════════════════════════
  //   TERRAIN/HEIGHT MULTIPLIER
  // ═══════════════════════════════════════════════
  sectionTitle('5. TERRAIN/HEIGHT MULTIPLIER (M_z,cat)','Clause 4.2');
  clauseNote('M_z,cat accounts for the variation of wind speed with height above ground and terrain roughness.');
  clauseNote('Values are obtained from Table 4.1 by interpolation for intermediate heights.');
  y+=2;

  paramRow('Reference Height (z)',R.h.toFixed(2)+' m','Clause 4.2.1');
  paramRow('Terrain Category','TC'+S.terrainCat,'Clause 4.2.1, Table 4.1');
  y+=2;

  const mzcatRows=S.dirs.map((d,i)=>[dirLongLabel(i),dirAngles[i]+'°',formatTcForReport(i),formatMzcatForReport(i)]);
  drawTable(
    ['Direction','Angle (θ)','Terrain Cat.','M_z,cat'],
    mzcatRows,
    [25,30,35,30],
    {highlightCol:3}
  );

  clauseNote('Terrain categories: TC1=Open, TC2=Few obstructions, TC2.5=Suburban, TC3=Industrial, TC4=Dense urban');
  clauseNote('Reference: Table 4.1(A) through 4.1(D) — Terrain/height multiplier, M_z,cat');
  y+=4;

  // ═══════════════════════════════════════════════
  //   SHIELDING MULTIPLIER
  // ═══════════════════════════════════════════════
  checkPage(50);
  sectionTitle('6. SHIELDING MULTIPLIER (M_s)','Clause 4.3');
  clauseNote('M_s accounts for the shielding effect of upwind buildings and structures.');
  clauseNote('M_s = 1.0 indicates no shielding. M_s < 1.0 indicates shielding benefit from surrounding structures.');
  y+=2;

  const msRows=S.dirs.map((d,i)=>[dirLongLabel(i),dirAngles[i]+'°',S.Ms[i].toFixed(2)]);
  drawTable(
    ['Direction','Angle (θ)','M_s'],
    msRows,
    [30,40,30],
    {highlightCol:2}
  );
  clauseNote('Shielding parameters are determined per Clause 4.3 and Table 4.3 of AS/NZS 1170.2:2021.');
  clauseNote('Where shielding is uncertain, M_s = 1.0 shall be adopted (conservative).');
  y+=4;

  // ═══════════════════════════════════════════════
  //   TOPOGRAPHIC MULTIPLIER
  // ═══════════════════════════════════════════════
  checkPage(50);
  sectionTitle('7. TOPOGRAPHIC MULTIPLIER (M_t)','Clause 4.4');
  clauseNote('M_t accounts for the increase in wind speed due to topographic features (hills, ridges, escarpments).');
  clauseNote('M_t ≥ 1.0. For flat terrain or sites not near topographic features, M_t = 1.0.');
  y+=2;

  const mtRows=S.dirs.map((d,i)=>[dirLongLabel(i),dirAngles[i]+'°',S.Mt[i].toFixed(2)]);
  drawTable(
    ['Direction','Angle (θ)','M_t'],
    mtRows,
    [30,40,30],
    {highlightCol:2}
  );
  clauseNote('Topographic multiplier determined per Clause 4.4 and Figures 4.4(a)-(d).');
  clauseNote('M_t combines hill-shape M_h (elevation) and lee M_lee per regional rules (max, product, or linear — see code). M_lee is listed separately on the Mlee tab for traceability.');
  y+=4;

  // Section 18 (TOPOGRAPHIC LONG SECTIONS) was originally rendered here, but
  // its TOC entry places it after 17b. The block has been moved to its
  // numerically-correct position (just before NOTES & DISCLAIMERS) so the
  // PDF reads in the order the TOC promises.

  // ═══════════════════════════════════════════════
  //   SITE WIND SPEED  (PAGE BREAK)
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('8. SITE WIND SPEED (V_sit,β)','Clause 2.2');
  clauseNote('The site wind speed for each cardinal direction, β:');
  const hasAnyLee = S.Mlee.some(m => m > 1);
  clauseNote('V_sit,β = V_R × M_d × M_z,cat × M_s × M_t   (Equation 2.2). M_t combines hill (Cl 4.4.2) and lee (Cl 4.4.3) per regional rules — not multiplied by M_lee again.');
  if(hasAnyLee) clauseNote('M_lee column below is for traceability; it is already included in M_t.');
  y+=3;

  const vsitRows=S.dirs.map((d,i)=>{
    const vsit=S.Vsit_dir[i];
    const row = [d,dirAngles[i]+'°',
      effectiveMd(i).toFixed(2), formatMzcatForReport(i), S.Ms[i].toFixed(2), S.Mt[i].toFixed(2)];
    if(hasAnyLee) row.push((S.Mlee[i]||1).toFixed(2));
    row.push(vsit.toFixed(1)+' m/s');
    return row;
  });
  const vsitHdrs = ['Dir','θ','M_d','M_z,cat','M_s','M_t'];
  const vsitWds = [15,15,20,23,20,20];
  if(hasAnyLee){ vsitHdrs.push('M_lee'); vsitWds.push(18); }
  vsitHdrs.push('V_sit,β (m/s)');
  vsitWds.push(hasAnyLee ? 26 : 32);
  drawTable(vsitHdrs, vsitRows, vsitWds, {highlightCol:vsitHdrs.length-1});

  // Show governing direction
  let maxVsit=0, maxDir='';
  S.Vsit_dir.forEach((v,i)=>{if(v>maxVsit){maxVsit=v;maxDir=S.dirs[i]}});
  y+=2;
  paramRow('Governing Direction',maxDir+' ('+maxVsit.toFixed(1)+' m/s)','Clause 2.2');
  y+=2;

  // Calculation walkthrough for governing
  subTitle('Calculation — Governing Direction','Eq. 2.2');
  calcLine(
    'V_sit,β = V_R × M_d × M_z,cat × M_s × M_t',
    '',''
  );
  calcLine(
    'V_sit,β = '+R.VR+' × '+R.Md+' × '+R.Mz.toFixed(3)+' × '+R.Ms+' × '+R.Mt,
    R.Vsit.toFixed(1)+' m/s',
    'Eq. 2.2'
  );
  y+=4;

  // ═══════════════════════════════════════════════
  //   DESIGN WIND SPEED
  // ═══════════════════════════════════════════════
  sectionTitle('9. DESIGN WIND SPEED (V_des,θ)','Clause 2.3');
  clauseNote('V_des,θ = max(V_sit,β) for directions within a 45° sector either side of θ');
  clauseNote('The design wind speed is the maximum site wind speed from the relevant sector (Clause 2.3).');
  y+=3;

  paramRow('Design Wind Speed V_des,θ',Math.max(...S.Vsit_dir).toFixed(1)+' m/s','Clause 2.3');
  paramRow('Governing Sector',maxDir+' ± 45°','Clause 2.3');
  y+=4;

  // ═══════════════════════════════════════════════
  //   DYNAMIC RESPONSE FACTOR
  // ═══════════════════════════════════════════════
  sectionTitle('10. DYNAMIC RESPONSE FACTOR (C_dyn)','Clause 6.1');
  clauseNote('For buildings with natural frequency ≥ 1 Hz (most low-to-mid-rise buildings):');
  y+=2;

  const bNatFreq = 46.0/R.h; // approximate natural frequency
  paramRow('Approx. Building Height',R.h.toFixed(2)+' m');
  paramRow('Approx. Natural Frequency (f_n)',bNatFreq.toFixed(2)+' Hz','Clause 6.1');
  if(bNatFreq>=1.0){
    paramRow('C_dyn','1.0 (building may be treated as static)','Clause 6.1');
    clauseNote('Since f_n ≥ 1.0 Hz, C_dyn = 1.0 per Clause 6.1. No dynamic analysis required.');
  } else {
    paramRow('C_dyn','Requires detailed dynamic analysis','Clause 6.2');
    clauseNote('WARNING: f_n < 1.0 Hz — building requires dynamic analysis per Clause 6.2.');
  }
  y+=4;

  // ═══════════════════════════════════════════════
  //   AERODYNAMIC SHAPE FACTOR
  // ═══════════════════════════════════════════════
  sectionTitle('11. AERODYNAMIC SHAPE FACTOR (C_fig)','Clause 5.2');
  clauseNote('C_fig = C_p,e × K_a × K_c,e × K_l × K_p   for external pressures  (Equation 5.2(1))');
  clauseNote('C_fig = C_p,i × K_c,i                        for internal pressures  (Equation 5.2(2))');
  y+=3;

  paramRow('K_a (Area Reduction Factor)','1.0','Table 5.4');
  const kce1Pdf = parseFloat(document.getElementById('kce1-val')?.value) || 0.8;
  const kci1Pdf = parseFloat(document.getElementById('kci1-val')?.value) || 1.0;
  const kcCase = document.getElementById('kc-design-case')?.value || 'f';
  const kcLabel = KC_DESIGN_CASES[kcCase] ? `Case (${kcCase})` : 'Custom';
  paramRow('K_c,e (Combination Factor - External)', kce1Pdf.toFixed(1), `Table 5.5 ${kcLabel}`);
  paramRow('K_c,i (Combination Factor - Internal)', kci1Pdf.toFixed(1), `Table 5.5 ${kcLabel}`);
  paramRow('K_l (Local Pressure Factor)','Varies — see Section 14','Table 5.6');
  paramRow('K_p (Porous Cladding Factor)','1.0','Clause 5.4.4');
  y+=4;

  // ═══════════════════════════════════════════════
  //   DESIGN WIND PRESSURE (qz)
  // ═══════════════════════════════════════════════
  subTitle('Design Wind Pressure','Clause 2.4.1');
  clauseNote('p = 0.5 × ρ_air × (V_des,θ)² × C_fig × C_dyn   (Equation 2.4(1))');
  clauseNote('where ρ_air = 1.2 kg/m³');
  y+=2;

  calcLine(
    'q_z = 0.5 × 1.2 × V_sit² / 1000',
    '',''
  );
  calcLine(
    'q_z = 0.5 × 1.2 × '+R.Vsit.toFixed(1)+'² / 1000',
    R.qz.toFixed(3)+' kPa',
    'Eq. 2.4(1)'
  );
  y+=6;

  // ═══════════════════════════════════════════════
  //   EXTERNAL PRESSURE COEFFICIENTS — WALLS
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('12. EXTERNAL PRESSURE COEFFICIENTS — WALLS','Clause 5.4, Table 5.2');
  clauseNote('External wall pressure coefficients per Table 5.2(A), 5.2(B), and 5.2(C) of AS/NZS 1170.2:2021.');
  y+=3;

  // Building geometry info
  const rHW = S.height / Math.min(R.effW, R.effD);
  subTitle('Building Geometry','Clause 5.4.1');
  paramRow('Building Width (b)',R.effW+' m');
  paramRow('Building Depth (d)',R.effD+' m');
  paramRow('Eave Height (h)',S.height+' m');
  paramRow('Reference Height (z)',R.h.toFixed(2)+' m');
  paramRow('h/d ratio',rHW.toFixed(3),'Table 5.2(A)');
  y+=3;

  subTitle('Wall Pressure Coefficients','Table 5.2');
  const sw1=R.faces.sidewall1, sw2=R.faces.sidewall2;
  const swZoneNote=(f)=>f.zones?f.zones.map(z=>z.dist+': '+z.Cpe.toFixed(2)).join(', '):'—';
  const ww = R.faces.windward;
  const wwZoned = ww.zones && ww.zones.length > 0;
  const wwNotes = wwZoned
    ? 'q_z varies with height; C_p,e = +0.8 (h > 25 m)'
    : (R.h > 25 ? '+0.8 (h > 25 m; wind at z = h)' : '+0.7 (h ≤ 25 m; wind at z = h)');
  const wallPressRows=[
    ['Windward Wall', ww.Cp_e.toFixed(2), wwZoned ? 'Zoned (AGL)' : '0 to h', wwNotes, 'Table 5.2(A)'],
    ['Leeward Wall',R.faces.leeward.Cp_e.toFixed(2),'0 to d','Varies with h/d','Table 5.2(B)'],
    ['Side Wall L','Zoned',swZoneNote(sw1),'From windward edge','Table 5.2(C)'],
    ['Side Wall R','Zoned',swZoneNote(sw2),'From windward edge','Table 5.2(C)'],
  ];
  drawTable(
    ['Surface','C_p,e','Distance','Notes','Reference'],
    wallPressRows,
    [35,20,30,45,30]
  );

  clauseNote('Windward wall: Table 5.2(A) — C_p,e = +0.7 when h ≤ 25 m (wind speed at z = h); C_p,e = +0.8 when h > 25 m with q_z varying with height on the windward face.');
  clauseNote('Leeward wall: C_p,e varies with h/d ratio per Table 5.2(B). h/d ≤ 0.25: -0.3; 0.25 < h/d ≤ 1: interpolate; h/d > 1: -0.5.');
  clauseNote('Side walls: C_p,e varies with distance from windward edge per Table 5.2(C).');
  y+=3;

  // Net wall pressures
  const qz=R.qz;
  subTitle('Net Wall Pressures','Clause 2.4.1');

  const cpiPdfDesign = getCpiCasesForDesign();
  const Cpi1Pdf = cpiPdfDesign.cpi1;
  const Cpi2Pdf = cpiPdfDesign.cpi2;

  const wallNetRows=[];
  for(let k of['windward','leeward','sidewall1','sidewall2']){
    const f=R.faces[k];
    const pe = (f.peExternalAvg != null && Number.isFinite(f.peExternalAvg))
      ? f.peExternalAvg
      : qz * f.Cp_e;
    const pnet1=(pe - qz*Cpi1Pdf).toFixed(3);
    const pnet2=(pe - qz*Cpi2Pdf).toFixed(3);
    wallNetRows.push([
      f.name,f.Cp_e.toFixed(2),(pe).toFixed(3),
      Cpi1Pdf.toFixed(2),Cpi2Pdf.toFixed(2),pnet1,pnet2,
      f.p.toFixed(3),f.force.toFixed(1)
    ]);
  }
  drawTable(
    ['Surface','C_p,e','p_e (kPa)','C_p,i(+)','C_p,i(-)','p_net(+) kPa','p_net(-) kPa','p_design kPa','Force kN'],
    wallNetRows,
    [25,15,18,16,16,22,22,22,18],
    {highlightCol:7}
  );
  y+=2;
  if(R.faces.windward.zones && R.faces.windward.zones.length>0){
    subTitle('Windward Wall Zone Breakdown','Table 5.2(A)');
    const wwZoneRows = R.faces.windward.zones.map(z=>[
      'Windward Wall', z.dist,
      z.Cpe.toFixed(2), z.qz.toFixed(3), z.area.toFixed(1), z.p.toFixed(3), z.force.toFixed(1)
    ]);
    drawTable(['Surface','Zone (AGL)','C_p,e','q_z kPa','Area m²','p kPa','Force kN'], wwZoneRows, [25,28,12,14,16,14,14]);
    y+=2;
  }
  // Sidewall zone breakdown (Table 5.2(C))
  if(sw1.zones && sw1.zones.length>0){
    subTitle('Side Wall Zone Breakdown','Table 5.2(C)');
    const swZoneRows=[];
    ['Side Wall L','Side Wall R'].forEach((name,i)=>{
      const f=R.faces['sidewall'+(i+1)];
      if(f.zones) f.zones.forEach(z=>{
        swZoneRows.push([name,z.dist,z.Cpe.toFixed(2),z.area.toFixed(1),(z.p).toFixed(3),z.force.toFixed(1)]);
      });
    });
    if(swZoneRows.length>0){
      drawTable(['Surface','Zone','C_p,e','Area m²','p kPa','Force kN'],swZoneRows,[25,18,15,18,18,18]);
      y+=2;
    }
  }

  // ═══════════════════════════════════════════════
  //   EXTERNAL PRESSURE COEFFICIENTS — ROOF
  // ═══════════════════════════════════════════════
  checkPage(60);
  sectionTitle('13. EXTERNAL PRESSURE COEFFICIENTS — ROOF','Clause 5.4, Table 5.3');
  clauseNote('Roof pressure coefficients depend on roof type, pitch angle, and distance from windward edge.');
  y+=2;

  subTitle('Roof Geometry','');
  paramRow('Roof Type',S.roofType.charAt(0).toUpperCase()+S.roofType.slice(1));
  paramRow('Roof Pitch',S.pitch+'°','Table 5.3(A),(B),(C)');
  paramRow('Parapet Height',S.parapet+' m','Clause 5.4.2');
  if(R.Kr < 1) paramRow('Kr (Parapet Reduction)',R.Kr.toFixed(2),'Table 5.7');
  paramRow('Overhang',S.overhang+' m','Clause 5.4.3');
  y+=3;

  subTitle('Roof Pressure Coefficients','Table 5.3');
  const roofPressRows=[
    [R.faces.roof_ww.name,R.faces.roof_ww.Cp_e.toFixed(2),'0 to h/2',
      S.pitch<=10?'Negative (flat/low pitch)':'Varies with pitch angle','Table 5.3(A)'],
    [R.faces.roof_lw.name,R.faces.roof_lw.Cp_e.toFixed(2),'h/2 to d',
      'Suction (downwind region)','Table 5.3(B)'],
  ];
  drawTable(
    ['Surface','C_p,e','Zone','Notes','Reference'],
    roofPressRows,
    [35,20,25,50,30]
  );

  // Net roof pressures
  subTitle('Net Roof Pressures','Clause 2.4.1');
  const roofNetRows=[];
  const roofFaceKeys = ['roof_ww','roof_lw','roof_cw'];
  if(R.faces.roof_hip_l) roofFaceKeys.push('roof_hip_l','roof_hip_r');
  for(let k of roofFaceKeys){
    const f=R.faces[k];
    if(!f) continue;
    const pe=qz*f.Cp_e;
    const pnet1=(pe - qz*Cpi1Pdf).toFixed(3);
    const pnet2=(pe - qz*Cpi2Pdf).toFixed(3);
    roofNetRows.push([
      f.name,f.Cp_e.toFixed(2),(pe).toFixed(3),
      pnet1,pnet2,f.p.toFixed(3),f.area.toFixed(1),f.force.toFixed(1)
    ]);
  }
  drawTable(
    ['Surface','C_p,e','p_e (kPa)','p_net(+) kPa','p_net(-) kPa','p_design kPa','Area m²','Force kN'],
    roofNetRows,
    [30,18,22,24,24,22,20,20],
    {highlightCol:5}
  );
  y+=2;
  clauseNote('For gable roofs with α ≤ 10°: C_p,e = -0.9 (windward 0-h/2) and -0.5 (leeward) per Table 5.3(A).');
  clauseNote('For pitch > 25°: Windward roof C_p,e may be positive (pressure) per Table 5.3(A).');
  y+=4;

  // ═══════════════════════════════════════════════
  //   INTERNAL PRESSURE COEFFICIENTS
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('14. INTERNAL PRESSURE COEFFICIENTS','Clause 5.3');
  clauseNote('Internal pressure depends on the permeability of the building envelope and the dominant opening condition.');
  y+=2;

  subTitle('Building Permeability','Clause 5.3');
  paramRow('Windward Openings',S.openWW+'%','Clause 5.3');
  paramRow('Leeward Openings',S.openLW+'%','Clause 5.3');
  paramRow('Side Openings',S.openSW+'%','Clause 5.3');
  y+=3;

  const totOpen = S.openWW+S.openLW+S.openSW*2;
  const cpiResultPdf = getCpiCasesForDesign();

  paramRow('Total Permeability',totOpen.toFixed(1)+'%');
  paramRow('Opening Condition',cpiResultPdf.clause);
  y+=3;

  subTitle('Internal Pressure Values','Table 5.1');
  drawTable(
    ['Condition','C_p,i (Positive)','C_p,i (Negative)','Reference'],
    [
      ['Case 1 — Opening on windward','+0.7','—','Table 5.1(A)'],
      ['Case 2 — Opening on leeward','—','-0.65','Table 5.1(A)'],
      ['Case 3 — Effectively sealed','+0.0','-0.2','Table 5.1(A)'],
      ['Case 4 — No dominant opening','+0.0','-0.3','Table 5.1(B)'],
      ['ADOPTED VALUE',Cpi1Pdf.toFixed(2),Cpi2Pdf.toFixed(2),'As per analysis'],
    ],
    [45,32,32,35]
  );

  clauseNote('Both positive and negative C_p,i values must be considered to determine the worst-case net pressure.');
  clauseNote('Reference: Clause 5.3 and Tables 5.1(A) & 5.1(B) of AS/NZS 1170.2:2021.');
  y+=4;

  // ═══════════════════════════════════════════════
  //   LOCAL PRESSURE FACTORS
  // ═══════════════════════════════════════════════
  sectionTitle('15. LOCAL PRESSURE FACTORS (K_l)','Clause 5.4.4, Table 5.6');
  clauseNote('Net design pressure p with local factor: p_net = p_e K_c,e − p_i (Table 5.5), with p_e,local = q_z C_p,e K_l K_p.');
  clauseNote('Wall and roof strips match Table 5.6 zones used in the Local Pressure Zones view (cladding on: full K_l; off: K_l = 1).');
  y+=3;

  const a_w=Math.min(0.2*Math.min(S.width,S.depth),S.R.h);
  const a_r=(S.R.h/S.width>=0.2||S.R.h/S.depth>=0.2)?0.2*Math.min(S.width,S.depth):2*S.R.h;
  paramRow('Dimension "a" (walls)',a_w.toFixed(2)+' m','min(0.2b, 0.2d, h) — Cl 5.4.4');
  paramRow('Dimension "a" (roofs)',a_r.toFixed(2)+' m','min(0.2b, 0.2d) or 2h — Cl 5.4.4');
  y+=2;

  subTitle('Walls — all Table 5.6 strips','');
  if(wallKlPdfRows.length){
    drawTable(
      ['Building face','Wind role','Zone','K_l','C_p,e','p_net (kPa)'],
      wallKlPdfRows,
      [22,28,22,14,16,22]
    );
  } else {
    clauseNote('[Wall K_l rows unavailable — run calculation]');
    y+=2;
  }

  checkPage(40);
  subTitle('Roofs — corner, edge/ridge, and interior K_l','Table 5.6');
  if(roofKlPdfRows.length){
    drawTable(
      ['Table 5.6 zone','Surface','K_l','C_p,e','p_net (kPa)'],
      roofKlPdfRows,
      [38,36,14,16,22]
    );
  } else {
    clauseNote('[Roof K_l rows unavailable]');
    y+=2;
  }
  clauseNote('Roof zone labels follow the 3D local-pressure mesh (RC = roof corner, RA = edge/ridge band, MR = interior).');
  clauseNote('Reference: Table 5.6 — Local pressure factor, K_l.');
  y+=4;

  // ═══════════════════════════════════════════════
  //   DESIGN PRESSURES SUMMARY
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('16. DESIGN WIND PRESSURES — SUMMARY','Clause 2.4');
  clauseNote('Design pressure: p = q_z × C_fig × C_dyn  where C_fig = (C_p,e − C_p,i) × K_a × K_c');
  clauseNote('Importance is incorporated via AS/NZS 1170.0 Table 3.3 (annual probability of exceedance) and regional V_R from AS/NZS 1170.2 Table 3.1 — not an additional factor on p.');
  y+=3;

  subTitle('Summary of All Face Pressures','');
  const summaryRows=[];
  for(let k in R.faces){
    const f=R.faces[k];
    summaryRows.push([
      f.name,
      f.Cp_e.toFixed(2),
      f.Cp_i.toFixed(2),
      (f.Cp_e-f.Cp_i).toFixed(2),
      R.qz.toFixed(3),
      '—',
      f.p.toFixed(3),
      f.area.toFixed(1),
      f.force.toFixed(1),
      'AS/NZS 1170.2 '+f.clause
    ]);
  }
  drawTable(
    ['Face','C_p,e','C_p,i','C_fig','q_z kPa','I (wind)','p (kPa)','Area m²','Force kN','Clause'],
    summaryRows,
    [26,14,14,14,17,10,17,17,17,30],
    {highlightCol:6}
  );

  // Calculation walkthrough for critical face
  let maxP=0,maxFace=null;
  for(let k in R.faces){if(Math.abs(R.faces[k].p)>maxP){maxP=Math.abs(R.faces[k].p);maxFace=R.faces[k]}}

  if(maxFace){
    y+=2;
    subTitle('Worked Example — '+maxFace.name+' (Critical Face)','Eq. 2.4(1)');
    calcLine('p = q_z × (C_p,e − C_p,i)','','');
    calcLine(
      'p = '+R.qz.toFixed(3)+' × ('+maxFace.Cp_e.toFixed(2)+' − '+maxFace.Cp_i.toFixed(2)+')',
      maxFace.p.toFixed(3)+' kPa','Eq. 2.4(1)'
    );
    calcLine(
      'F = |p| × A = |'+maxFace.p.toFixed(3)+'| × '+maxFace.area.toFixed(1),
      maxFace.force.toFixed(1)+' kN','Force'
    );
  }
  y+=4;

  // Total base shear
  let totalForceX=0, totalForceZ=0;
  totalForceX = Math.abs(R.faces.windward.force) + Math.abs(R.faces.leeward.force);
  totalForceZ = Math.abs(R.faces.roof_ww.force) + Math.abs(R.faces.roof_lw.force);
  if(R.faces.roof_hip_l && R.faces.roof_hip_r){
    totalForceZ += Math.abs(R.faces.roof_hip_l.force) + Math.abs(R.faces.roof_hip_r.force);
  }

  subTitle('Total Building Forces','');
  paramRow('Base Shear (Windward → Leeward)',totalForceX.toFixed(1)+' kN');
  paramRow('Roof Uplift (Net)',totalForceZ.toFixed(1)+' kN');
  paramRow('Total Lateral',totalForceX.toFixed(1)+' kN','Along-wind');
  y+=4;

  // ═══════════════════════════════════════════════
  //   CARDINAL DIRECTIONS — SPEEDS & FACE PRESSURES
  // ═══════════════════════════════════════════════
  sectionTitle('17. CARDINAL DIRECTIONS — WIND SPEEDS & DESIGN PRESSURES','Clause 2.2, 2.4');
  clauseNote('Wind from each compass sector (θ = direction wind blows from). Multipliers and V_sit,β match Section 7.');
  clauseNote('Windward building-face orientation for each direction is also shown (e.g., N wind corresponds to windward face NNW at θ = 337.5°).');
  clauseNote('Face labels are wind-relative roles for that direction (WW = windward wall, SW1/SW2 = side walls, etc.), not fixed building north/south.');
  y+=2;

  const faceShortHdr = { windward:'WW', leeward:'LW', sidewall1:'SW1', sidewall2:'SW2', roof_ww:'Roof U', roof_lw:'Roof D', roof_cw:'Roof R', roof_hip_l:'Hip L', roof_hip_r:'Hip R' };
  const dirMultRows = cardinalData.map((cd, i) => [
    dirLongLabel(i), windwardFaceShortLabel(i),
    R.VR.toFixed(0),
    effectiveMd(i).toFixed(2), formatMzcatForReport(i), S.Ms[i].toFixed(2), S.Mt[i].toFixed(2),
    S.Vsit_dir[i].toFixed(1), cd.qz.toFixed(3)
  ]);
  drawTable(
    ['Direction', 'WW face θ', 'V_R', 'M_d', 'M_z,cat', 'M_s', 'M_t', 'V_sit,β m/s', 'q_z kPa'],
    dirMultRows,
    [22, 22, 12, 14, 18, 12, 14, 22, 18],
    { highlightCol: 8 }
  );
  y += 2;

  checkPage(45);
  subTitle('Design pressure p (kPa) by wind-relative face','Eq. 2.4(1)');
  const pHdr = ['Direction', 'WW face θ'];
  faceKeyOrder.forEach(k => { if(cardinalData[0].faces[k]) pHdr.push(faceShortHdr[k] || k); });
  const pRows = cardinalData.map((cd, i) => {
    const row = [dirLongLabel(i), windwardFaceShortLabel(i)];
    faceKeyOrder.forEach(k => {
      if(cd.faces[k]) row.push(cd.faces[k].p.toFixed(3));
    });
    return row;
  });
  const pColW = pHdr.map((_, idx) => (idx < 2 ? 20 : 14));
  drawTable(pHdr, pRows, pColW, { highlightCol: pHdr.length - 1 });
  y += 2;

  checkPage(18);
  doc.setFontSize(8);
  doc.setTextColor(...COL_MUTED);
  for(let i=0;i<8;i++){
    doc.text(`${dirLongLabel(i)}: windward building face ${windwardFaceFullLabel(i)}`, ML + 2, y);
    y += 4.2;
  }
  y += 1;

  // Phase 22: CHECKWIND-style per-direction full pressure analysis. The summary
  // tables above give the side-by-side view; here we devote one block per
  // direction so each direction's design pressures, internal pressure cases,
  // and forces are documented end-to-end (matching the example CHECKWIND
  // report layout where every windward-direction has its own page-section).
  doc.addPage();pg++;
  addHeader();y+=3;
  sectionTitle('17b. PER-DIRECTION DETAILED PRESSURES','Clause 2.4, 5.3, 5.4');
  clauseNote('For each cardinal wind direction (θ = direction wind blows from), the tables below give every face\'s C_p,e, internal pressure case, net pressures, design pressure and force, and the Table 5.6 Local Pressure Zone breakdown for walls and roof.');
  clauseNote('Face roles (WW/LW/SW1/SW2/Roof) are wind-relative for that direction; the building geometry is fixed but face roles rotate with wind.');
  clauseNote('Each direction is laid out on its own page for clarity.');
  y += 2;

  for(let i=0;i<8;i++){
    const cd = cardinalData[i];
    if(!cd) continue;
    // Phase 23: each direction starts on a fresh page so the per-direction
    // multipliers, face pressures, and LPZ tables stay together. Skip the
    // page break for the very first direction (we just opened a page above).
    if(i > 0){
      doc.addPage();pg++;
      addHeader();y+=3;
    }
    subTitle(`Direction ${dirLongLabel(i)} — windward face ${windwardFaceFullLabel(i)}`,'Eq. 2.4(1)');

    paramRow('V_sit,β', cd.Vsit.toFixed(1)+' m/s','Eq. 2.2');
    paramRow('q_z', cd.qz.toFixed(3)+' kPa','Eq. 2.4(1)');
    paramRow('M_d',
      Number.isFinite(cd.Md) ? cd.Md.toFixed(2) : String(cd.Md),
      'Table 3.2');
    paramRow('M_z,cat', String(cd.MzcatStr || ''),'Table 4.1');
    paramRow('M_s',
      Number.isFinite(cd.Ms) ? cd.Ms.toFixed(2) : String(cd.Ms),
      'Clause 4.3');
    paramRow('M_t',
      Number.isFinite(cd.Mt) ? cd.Mt.toFixed(2) : String(cd.Mt),
      'Clause 4.4');
    paramRow('Internal pressure cases',
      `C_p,i(+) = ${Number.isFinite(cd.cpi1) ? cd.cpi1.toFixed(2) : '—'}, C_p,i(−) = ${Number.isFinite(cd.cpi2) ? cd.cpi2.toFixed(2) : '—'}`,
      cd.cpiClause || 'Table 5.1');
    y += 1;

    const dirRows = buildDirectionFaceRows(cd.faces, faceKeyOrder, cd.qz, cd.cpi1, cd.cpi2);
    if(dirRows.length){
      drawTable(
        ['Surface','C_p,e','C_p,i','p_e kPa','p_net(+) kPa','p_net(−) kPa','p_design kPa','Area m²','Force kN'],
        dirRows,
        [28, 14, 14, 18, 22, 22, 22, 18, 18],
        { highlightCol: 6 }
      );
    } else {
      clauseNote('[no faces captured for this direction]');
    }
    y += 3;

    // Phase 23: per-direction Local Pressure Zone tables. Same column layout as
    // the on-screen LOCAL (WALLS) and LOCAL (ROOF) tabs — Surface, Distance
    // from edge, Reference, Area, K_l, K_c,i, K_c,e, K_p, V_sit,β, q_z,
    // C_shp,e, p_e, p_net(+), p_net(−).
    const lpzHdrs = ['Surface','Distance','Ref','Area','K_l','K_c,i','K_c,e','K_p','V_sit,β','q_z','C_shp,e','p_e','p_net(+)','p_net(−)'];
    const lpzCols = [18, 28, 12, 24, 9, 9, 9, 9, 10, 11, 11, 10, 10, 10];
    if(cd.lpzWalls && cd.lpzWalls.length){
      checkPage(50);
      subTitle('Local Pressure Zones — Walls','Cl 5.4.4, Table 5.6');
      drawTable(lpzHdrs, cd.lpzWalls, lpzCols, { highlightCol: 13 });
      y += 2;
    }
    if(cd.lpzRoof && cd.lpzRoof.length){
      checkPage(50);
      subTitle('Local Pressure Zones — Roof','Cl 5.4.4, Table 5.6');
      drawTable(lpzHdrs, cd.lpzRoof, lpzCols, { highlightCol: 13 });
      y += 2;
    }
    y += 2;
  }

  // Wind rose removed from the PDF report — directional V_sit,β is fully
  // documented in the per-direction tables above (Sections 4-8 and 17/17b).

  // ═══════════════════════════════════════════════
  //   TOPOGRAPHIC LONG SECTIONS (CARDINAL DIRECTIONS)
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('18. TOPOGRAPHIC LONG SECTIONS (CARDINAL DIRECTIONS)','Clause 4.4.2');
  clauseNote('Long sections below are included for each cardinal direction and use the same elevation/profile pipeline as the Topography tab.');
  y += 2;

  const topoProfiles = Array.isArray(S.detectedProfiles) ? S.detectedProfiles : [];
  const topoDetails = Array.isArray(S.mhDetailsSub) ? S.mhDetailsSub : [];
  const topoNSub = topoProfiles.length;
  const topoSiteElev = Number.isFinite(S.detectedSiteElev) ? S.detectedSiteElev : (Number(S.elevation) || 0);
  const topoBuildH = Number.isFinite(R.h) ? R.h : Number(S.height || 0);

  if(topoNSub > 0 && topoDetails.length === topoNSub){
    const topoCanvasWidth = 1000;
    const topoCanvasHeight = 220;
    const topoImgHeight = (CW * topoCanvasHeight) / topoCanvasWidth;
    for(let i=0; i<8; i++){
      const subIdx = Math.round((i * topoNSub) / 8) % topoNSub;
      const prof = topoProfiles[subIdx];
      const det = topoDetails[subIdx];
      const dirLabel = `${dirLongLabel(i)} (${dirAngles[i]}°)`;

      checkPage(topoImgHeight + 16);
      subTitle(`Topographic Long Section — ${dirLabel}`,'');

      try{
        const topoCanvas = document.createElement('canvas');
        drawTopoCanvas(
          topoCanvas,
          prof,
          det,
          topoSiteElev,
          topoBuildH,
          subIdx,
          topoNSub,
          topoProfiles,
          { fixedWidth: topoCanvasWidth }
        );
        const topoImg = topoCanvas.toDataURL('image/png');
        doc.addImage(topoImg, 'PNG', ML, y, CW, topoImgHeight);
        y += topoImgHeight + 3;
      } catch(err){
        doc.setFontSize(8);
        doc.setTextColor(...COL_MUTED);
        doc.text(`[Topographic long section unavailable for ${dirLabel}]`, ML + 3, y + 4);
        y += 8;
      }
    }
  } else {
    doc.setFontSize(8);
    doc.setTextColor(...COL_MUTED);
    doc.text('Topographic long sections not available — run site detection/topography sampling first.', ML + 3, y + 4);
    y += 8;
  }
  y += 2;

  // ═══════════════════════════════════════════════
  //   NOTES & DISCLAIMERS
  // ═══════════════════════════════════════════════
  doc.addPage();pg++;
  addHeader();y+=3;

  sectionTitle('19. NOTES & DISCLAIMERS','');
  y+=2;

  const notes=[
    'This report has been generated in accordance with AS/NZS 1170.2:2021 "Structural design actions — Wind actions".',
    'All wind multipliers should be verified by the engineer of record for the specific site conditions.',
    'Wind Region: As classified per Figure 3.1 and Table 3.1 of AS/NZS 1170.2:2021.',
    'Terrain Category: Should be assessed per Clause 4.2 considering the actual upwind terrain for each direction. Terrain may vary by direction.',
    'Shielding: Where shielding from adjacent buildings is relied upon, the permanence of shielding buildings must be considered (Clause 4.3.4).',
    'Topographic Multiplier: Sites on or near hills, ridges, or escarpments require assessment per Clause 4.4 and Figures 4.4(a)-(d).',
    'Internal Pressure: The internal pressure coefficient depends on the permeability condition of the building and any dominant openings (Clause 5.3).',
    'Local Pressures: Higher local pressures at edges, corners, and ridges must be used for the design of cladding and fixings (Clause 5.4.4, Table 5.6).',
    'Dynamic Response: Buildings with natural frequency less than 1.0 Hz require a dynamic analysis per Section 6.',
    'This software provides calculations as a design aid only. The engineer of record is responsible for verification of all inputs, assumptions, and results.',
    'Combination factors (K_c) must be applied where internal and external pressures act simultaneously (Table 5.5).',
    'For cyclonic regions, additional requirements of Clause 3.2 apply including M_d = 1.0 for all directions.',
  ];

  doc.setFontSize(8);doc.setTextColor(...COL_TEXT);
  notes.forEach((note,i)=>{
    checkPage(12);
    doc.setTextColor(...COL_PRIMARY);doc.setFont(undefined,'bold');
    doc.text((i+1)+'.',ML+3,y);
    doc.setFont(undefined,'normal');doc.setTextColor(...COL_TEXT);
    const lines=doc.splitTextToSize(note,CW-15);
    doc.text(lines,ML+10,y);
    y+=lines.length*4.5+2;
  });

  y+=6;
  subTitle('Reference Standards','');
  const refs=[
    'AS/NZS 1170.2:2021 — Structural design actions — Wind actions',
    'AS/NZS 1170.0:2002 — Structural design actions — General principles',
    'NZS 3604:2011 — Timber-framed buildings (for wind zone classification)',
    'AS/NZS 1170.2:2021 Supplement 1 — Commentary',
  ];
  refs.forEach(ref=>{
    checkPage(8);
    doc.setFontSize(8);doc.setTextColor(...COL_TEXT);
    doc.text('• '+ref,ML+5,y);y+=5.5;
  });

  y+=8;
  doc.setDrawColor(...COL_ACCENT);
  doc.line(ML,y,ML+CW,y);y+=5;
  doc.setFontSize(8);doc.setTextColor(...COL_MUTED);
  doc.text('END OF REPORT',PW/2-12,y);
  y+=5;
  doc.text('Standard: AS/NZS 1170.2:2021',ML,y);
  y+=4;
  doc.text('Report generated: '+new Date().toISOString(),ML,y);

  // ═══════════════════════════════════════════════
  //   POLISH PASS — DYNAMIC TOC + PAGE X of Y FOOTER
  // ═══════════════════════════════════════════════
  // Total pages is now known. Rewrite the TOC on its reserved page using the
  // captured `sectionPages`, and stamp every page's footer with "Page X of Y".
  const TOTAL_PAGES = pg;
  try {
    // 1) Repaint TOC on the reserved page with real section pages.
    if(TOC_PAGE_NUMBER){
      doc.setPage(TOC_PAGE_NUMBER);
      // Clear the body area below the header (header occupies 0..12mm).
      // Leave the footer band intact (PH-15..PH).
      doc.setFillColor(255,255,255);
      doc.rect(0, 12, PW, PH-15-12, 'F');
      // Re-render the TOC content at the same Y the first pass used: addHeader
      // sets y = MT+2, then "y+=5" advances to MT+7 before drawToc(y) runs.
      drawToc(MT + 7);
    }
    // 2) Stamp "Page X of Y" on every page's footer (skip the cover, which
    // has no footer band).
    for(let p=2; p<=TOTAL_PAGES; p++){
      doc.setPage(p);
      // Repaint just the page-number cell so we don't disturb header/cover.
      doc.setFillColor(240,240,240);
      doc.rect(PW-MR-22, PH-12, 22, 9, 'F');
      doc.setFontSize(7);doc.setTextColor(...COL_MUTED);
      doc.text('Page '+p+' of '+TOTAL_PAGES, PW-MR-2, PH-7, {align:'right'});
    }
    // Restore page state so anything appended later still works.
    doc.setPage(TOTAL_PAGES);
  } catch(e){
    console.warn('TOC / page-footer rewrite failed:', e);
  }

  // ═══════════════════════════════════════════════
  //   SAVE
  // ═══════════════════════════════════════════════
  if(returnBlob){
    return doc.output('blob');
  }
  if(saveFile){
    doc.save('Wind_Report_'+S.region+'_'+new Date().toISOString().slice(0,10)+'.pdf');
    if(!silent) toast('Detailed PDF report exported — '+pg+' pages!');
  }
  return doc;
}

function exportAllProjects(){
  const data = JSON.stringify({
    version:'v5',
    building:{width:S.width,depth:S.depth,height:S.height,pitch:S.pitch,
      roofType:S.roofType},
    wind:{speed:S.windSpeed,terrain:S.terrainCat,importance:S.importance,
      life:S.life,ari:S.ari,vrManual:S.vrManual,
      loadCase:S.loadCase,region:S.region,angle:S.windAngle},
    location:{lat:S.lat,lng:S.lng,address:S.address},
    directions:{Md:S.Md,Mzcat:S.Mzcat,Ms:S.Ms,Mt:S.Mt,Vsit:S.Vsit_dir},
    results:S.R
  }, null, 2);
  const blob = new Blob([data],{type:'application/json'});
  const a = document.createElement('a');a.download='wind-project.json';
  a.href=URL.createObjectURL(blob); a.click();
  toast('Project exported as JSON');
}

function printReport(){
  updateDocPreview();
  const content = document.getElementById('doc-preview').innerHTML;
  const win = window.open('','','width=900,height=700');
  win.document.write(`<html><head><title>Wind Report</title>
    <style>body{font-family:Segoe UI,sans-serif;padding:20px}table{border-collapse:collapse;width:100%;margin:8px 0}
    th,td{border:1px solid #ddd;padding:6px 8px;font-size:11px}th{background:#f0f4f8;font-weight:700}
    h2{color:#1a5276;border-bottom:2px solid #1a5276;padding-bottom:4px}</style></head>
    <body>${content}</body></html>`);
  win.document.close();
  win.print();
}

// ═══════════════════════════════════════════════
//   ACCOUNT MENU
// ═══════════════════════════════════════════════
function toggleAccountMenu(){
  const m = document.getElementById('account-menu');
  if(m.classList.contains('show')){
    m.classList.remove('show');
  } else {
    m.classList.add('show');
  }
}

// Close account menu when clicking outside
document.addEventListener('click', function(e){
  const menu = document.getElementById('account-menu');
  const btn = document.getElementById('btn-account');
  if(menu && btn && !menu.contains(e.target) && !btn.contains(e.target)){
    menu.classList.remove('show');
  }
});

// ═══════════════════════════════════════════════
//   TOAST
// ═══════════════════════════════════════════════
function toast(msg){
  const t=document.getElementById('toast');
  if(!t) return;
  t.textContent=msg;t.classList.add('show');
  clearTimeout(t._tid);t._tid=setTimeout(()=>t.classList.remove('show'),3500);
}

// ═══════════════════════════════════════════════
//   ANIMATION LOOP
// ═══════════════════════════════════════════════
function animate(){
  requestAnimationFrame(animate);
  clock+=.016;
  controls.update();

  tickParticles();
  renderer.render(scene,camera);
}

// ═══════════════════════════════════════════════
//   INIT DIRECTION TABLE ON LOAD
// ═══════════════════════════════════════════════
setTimeout(()=>{
  refreshDirectionalWindUI();
}, 500);


// ═══════════════════════════════════════════════════════════════════
//   FIREBASE AUTHENTICATION & STRIPE PAYMENT INTEGRATION
// ═══════════════════════════════════════════════════════════════════

function initAuthListener(){
  currentUser = null;
  S.user.signedIn = false;
  S.user.name = 'Local Workspace';
  S.user.plan = 'local';
  userSubscription = { active: true, plan: 'local', teamInvitee: false };

  document.getElementById('acct-name').textContent = 'Local Workspace';
  document.getElementById('acct-plan').textContent = 'Local Workspace';
  document.getElementById('btn-signin-menu').style.display = 'none';
  document.getElementById('btn-signout-menu').style.display = 'none';
  document.getElementById('btn-dashboard-menu').style.display = 'none';
  updatePlanUI();
  closeAuthOverlay();
  try{ refreshDirectionalWindUI(); }catch(e){}
  try{ updateIfcUploadSigninHint(); }catch(e){}
}

async function checkSubscriptionStatus(){
  userSubscription = { active: true, plan: 'local', teamInvitee: false, members: [], maxMembers: 1 };
  S.user.plan = userSubscription.plan;
  updatePlanUI();
  if(typeof updateDocPreview === 'function') try { updateDocPreview(); } catch(e){}
}

function updatePlanUI(){
  const label = PLAN_LABELS[userSubscription.plan] || 'Local Workspace';
  document.getElementById('acct-plan').textContent = label;
  const show = (id, cond) => { const el = document.getElementById(id); if(el) el.style.display = cond ? 'block' : 'none'; };
  show('btn-saved-projects', true);
  show('btn-shared-projects', false);
  show('btn-share-project-menu', false);
  show('btn-activity-log', false);
  show('btn-api-keys', false);
  show('btn-templates', false);
  const upgradeBtn = document.getElementById('btn-upgrade');
  if(upgradeBtn) upgradeBtn.style.display = 'none';
  const billingBtn = document.getElementById('btn-manage-billing');
  if(billingBtn) billingBtn.style.display = 'none';
  updateIfcAiAssistUI();
}

// ═══ Auth Overlay ═══
function showAuthOverlay(){
  document.getElementById('auth-overlay').classList.add('show');
  document.getElementById('account-menu').classList.remove('show');
}

function closeAuthOverlay(){
  document.getElementById('auth-overlay').classList.remove('show');
}

// ═══ Google Sign In ═══
async function signInWithGoogle(){
  showAuthOverlay();
}

// ═══ Microsoft Sign In ═══
async function signInWithMicrosoft(){
  showAuthOverlay();
}

// ═══ Sign Out ═══
async function signOutUser(){
  document.getElementById('account-menu').classList.remove('show');
  toast('This local build does not use cloud sign-in.');
}

// ═══ Payment Modal ═══
function openPaymentModal(){
  document.getElementById('account-menu').classList.remove('show');
  toast('Paid plans are removed in this local build.');
}

function closePaymentModal(){}

// ═══ Stripe Checkout ═══
async function startCheckout(plan){
  toast('Checkout is disabled in this local build.');
  return;
}

async function __unused_startCheckout(plan){
  toast('Checkout is removed in this local build.');
}

function openEnterpriseEnquiry(){
  toast('Enterprise/cloud sales flows are removed in this local build.');
}

async function submitEnterpriseEnquiry(e){
  if(e && typeof e.preventDefault === 'function') e.preventDefault();
  toast('Enterprise/cloud sales flows are removed in this local build.');
}

// ═══ Billing Portal ═══
async function openBillingPortal(){
  toast('Billing is disabled in this local build.');
  return;
}

async function __unused_openBillingPortal(){
  toast('Billing is removed in this local build.');
}

// ═══ Check for checkout result in URL ═══
function checkCheckoutResult(){
  const params = new URLSearchParams(window.location.search);
  const checkout = params.get('checkout');

  if(checkout === 'success'){
    toast('Checkout parameters ignored in the local build.');
    window.history.replaceState({}, '', window.location.pathname);
  } else if(checkout === 'cancelled'){
    toast('Checkout is disabled in the local build.');
    window.history.replaceState({}, '', window.location.pathname);
  }
}

function pollSubscriptionStatus(attempt){
  const maxAttempts = 8;
  const delays = [1000, 2000, 2000, 3000, 3000, 5000, 5000, 5000]; // total ~26 seconds
  if(attempt >= maxAttempts) return;

  setTimeout(async ()=>{
    if(!currentUser) return;
    await checkSubscriptionStatus();
    try{ refreshDirectionalWindUI(); }catch(e){}
    if(userSubscription.active){
      toast('✅ Plan activated: ' + (PLAN_LABELS[userSubscription.plan] || userSubscription.plan));
      closePaymentModal();
    } else {
      // Not active yet — try again
      pollSubscriptionStatus(attempt + 1);
    }
  }, delays[attempt] || 3000);
}

// ═══════════════════════════════════════════════════
//   SHARED PROJECTS (Enterprise & Team)
// ═══════════════════════════════════════════════════
async function openSharedProjects(){
  toast('Shared projects are not implemented in this local build yet.');
}

function showSharedProjectsOverlay(projects){
  // Remove existing overlay if present
  let overlay = document.getElementById('shared-projects-overlay');
  if(overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'shared-projects-overlay';
  overlay.className = 'modal-overlay show';
  overlay.style.zIndex = '8500';

  const planName = userSubscription.plan;
  const maxMembers = PLAN_MAX_MEMBERS[planName] || 1;
  const members = userSubscription.members || [];
  const showTeamMgmt = hasSharedProjects();

  let projRows = '';
  if(projects.length === 0){
    const emptyMsg = showTeamMgmt
      ? 'No shared projects yet. Share a project from the Detailed Report tab.'
      : 'No projects shared with you yet. Ask your team owner to share a project from the Detailed Report tab.';
    projRows = '<tr><td colspan="5" style="text-align:center;color:var(--text3);padding:20px">' + emptyMsg + '</td></tr>';
  } else {
    projects.forEach(p => {
      const lastEditor = p.lastEditedBy || p.owner || 'Unknown';
      const lastEditorShort = lastEditor.split('@')[0];
      projRows += `<tr>
        <td style="font-weight:600">${p.name || 'Untitled'}</td>
        <td style="font-size:11px">${(p.owner||'').split('@')[0] || 'Unknown'}</td>
        <td style="font-size:11px">${lastEditorShort}</td>
        <td style="font-size:10px;white-space:nowrap">${new Date(p.updated || Date.now()).toLocaleString()}</td>
        <td style="white-space:nowrap">
          <button onclick="loadSharedProject('${p.id}')" style="background:var(--primary);color:var(--bg);border:none;padding:4px 10px;border-radius:4px;font-size:10px;cursor:pointer;margin-right:4px">Open</button>
          <button onclick="viewProjectHistory('${p.id}')" style="background:var(--surface3);border:1px solid var(--border);color:var(--text);padding:4px 10px;border-radius:4px;font-size:10px;cursor:pointer" title="View edit history">History</button>
        </td>
      </tr>`;
    });
  }

  const teamSection = showTeamMgmt
    ? `<p style="color:var(--text2);font-size:12px;margin-bottom:4px">Plan: <b>${PLAN_LABELS[planName]}</b> | ${members.length}/${maxMembers} members</p>
      <div style="margin:12px 0">
        <h4 style="color:var(--text2);margin-bottom:6px">Team Members</h4>
        <div id="shared-members-list" style="font-size:12px;color:var(--text);margin-bottom:8px">
          ${members.length ? members.map(m => '<span style="background:var(--surface2);padding:2px 8px;border-radius:4px;margin:2px;display:inline-block">'+m+'</span>').join('') : '<span style="color:var(--text3)">No members yet</span>'}
        </div>
        ${maxMembers > 1 ? '<button onclick="inviteTeamMember()" style="background:var(--surface3);border:1px solid var(--border);color:var(--text);padding:5px 12px;border-radius:4px;font-size:11px;cursor:pointer">+ Invite Member</button>' : ''}
      </div>`
    : `<p style="color:var(--text2);font-size:12px;margin-bottom:12px">You were invited to collaborate on team projects. Only the team owner can invite people or change the team plan.</p>`;

  overlay.innerHTML = `
    <div class="modal-card" style="max-width:800px">
      <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">X</button>
      <h2 style="color:var(--primary);margin-bottom:12px">Shared Projects</h2>
      ${teamSection}
      <div style="border-top:1px solid var(--border);padding-top:12px">
        <h4 style="color:var(--text2);margin-bottom:8px">Projects</h4>
        <table class="result-table" style="width:100%">
          <thead><tr><th>Project</th><th>Owner</th><th>Last Edited By</th><th>Updated</th><th></th></tr></thead>
          <tbody>${projRows}</tbody>
        </table>
      </div>
    </div>`;

  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if(e.target === overlay) overlay.remove(); });
}

async function inviteTeamMember(){
  toast('Team invites are not available in this local build.');
}

async function loadSharedProject(projectId){
  toast('Shared projects are not implemented in this local build yet.');
}

// Save updates back to a shared project
async function saveSharedProject(){
  toast('Shared project sync is not implemented in this local build yet.');
}

// View edit history for a shared project
async function viewProjectHistory(projectId){
  toast('Project history is not implemented in this local build yet.');
}

async function shareCurrentProject(){
  toast('Project sharing is not available in this local build yet.');
}

// ═══════════════════════════════════════════════════
//   SAVE / LOAD PROJECTS (Pro+)
// ═══════════════════════════════════════════════════
let currentProjectId = null;
let currentProjectName = 'Untitled Project';
let currentProjectShared = false; // whether loaded project is a shared team project

function getProjectState(){
  readKvOverrideFromUi();
  return {
    width: S.width, depth: S.depth, height: S.height, pitch: S.pitch,
    roofType: S.roofType, parapet: S.parapet, overhang: S.overhang,
    windSpeed: S.windSpeed, terrainCat: S.terrainCat, importance: S.importance,
    loadCase: S.loadCase, windAngle: S.windAngle, Kp: S.Kp,
    openWW: S.openWW, openLW: S.openLW, openSW: S.openSW,
    lat: S.lat, lng: S.lng, address: S.address, elevation: S.elevation,
    region: S.region, ari: S.ari, svcVr: S.svcVr, life: S.life, vrManual: S.vrManual,
    mapBuildingAngle: S.mapBuildingAngle,
    Md: S.Md, Mzcat: S.Mzcat, Ms: S.Ms, Mt: S.Mt, Mt_hill: S.Mt_hill, Mlee: S.Mlee,
    TC_dir: S.TC_dir,
    terrainRecalcCtx: S.terrainRecalcCtx || null,
    detectedBuildingsList: S.detectedBuildingsList || null,
    kcDesignCase: document.getElementById('kc-design-case')?.value || 'f',
    kci: document.getElementById('kci1-val')?.value || '1.0',
    kce: document.getElementById('kce1-val')?.value || '0.9',
    kvOverride: S.kvOverride,
    volKvManual: (S.kvOverride && Number.isFinite(S.volKvManual) && S.volKvManual > 0) ? S.volKvManual : null
  };
}

/** Firestore stores detectedBuildingsList / terrainRecalcCtx as JSON strings (nested arrays are invalid). */
function restoreFirestoreStateJsonFields(target){
  if(!target) return;
  if(typeof target.detectedBuildingsList === 'string'){
    try { target.detectedBuildingsList = JSON.parse(target.detectedBuildingsList); }
    catch(e){ target.detectedBuildingsList = null; }
  }
  if(typeof target.terrainRecalcCtx === 'string'){
    try { target.terrainRecalcCtx = JSON.parse(target.terrainRecalcCtx); }
    catch(e){ target.terrainRecalcCtx = null; }
  }
}

/** After Object.assign(S, state): old saves omit volKvManual — reset so it does not leak from prior session. */
function mergeProjectStateDefaults(state){
  if(state && !Object.prototype.hasOwnProperty.call(state, 'volKvManual'))
    S.volKvManual = NaN;
  delete S.buildingHeightOverrides;
  restoreFirestoreStateJsonFields(S);
}

// Push current S state values into UI input elements
function pushStateToUI(){
  dirPolarClosePopover();
  if(S.volKvManual != null && (!Number.isFinite(S.volKvManual) || S.volKvManual <= 0))
    S.volKvManual = NaN;
  const setV = (id, v) => { const el = document.getElementById(id); if(el) el.value = v; };
  setV('inp-width', S.width);
  setV('inp-depth', S.depth);
  setV('inp-height', S.height);
  setV('inp-pitch', S.pitch);
  setV('inp-rooftype', S.roofType);
  syncRoofTypeButtons(S.roofType || 'gable');
  setV('inp-parapet', S.parapet);
  setV('inp-overhang', S.overhang);
  setV('inp-windspeed', S.windSpeed);
  if(S.svcVr != null) setV('inp-svc-vr', S.svcVr);
  setV('inp-terrain', S.terrainCat);
  setV('inp-importance', S.importance);
  setV('inp-life', S.life != null ? S.life : 50);
  syncImportanceSegButtons();
  syncDesignLifeSegButtons();
  setV('inp-loadcase', S.loadCase);
  setV('inp-kp', S.Kp);
  [['ww', S.openWW], ['lw', S.openLW], ['sw', S.openSW]].forEach(([face, v])=>{
    const n = v != null && Number.isFinite(Number(v)) ? Number(v) : 0;
    setV('inp-open-'+face, n);
    setV('inp-open-'+face+'-range', n);
    const valEl = document.getElementById('val-open-'+face);
    if(valEl) valEl.textContent = formatOpeningPctLabel(n);
  });
  setV('inp-region', S.region);
  const regDisp = document.getElementById('region-display');
  if(regDisp && S.region != null) regDisp.textContent = S.region;
  setV('inp-lat', S.lat);
  setV('inp-lng', S.lng);
  setV('inp-winddir', S.windAngle);
  const wdLbl = document.getElementById('val-winddir');
  if(wdLbl) wdLbl.textContent = Math.round(S.windAngle);

  if(S.kcDesignCase) setV('kc-design-case', S.kcDesignCase);
  if(S.kci || S.kci1) setV('kci1-val', S.kci || S.kci1);
  if(S.kce || S.kce1) setV('kce1-val', S.kce || S.kce1);

  const kvT = document.getElementById('kv-override-toggle');
  const inpVol = document.getElementById('inp-vol-kv');
  if(kvT) kvT.checked = !!S.kvOverride;
  if(S.kvManual != null && Number.isFinite(Number(S.kvManual))){
    const km = Number(S.kvManual);
    if(km >= 0.5 && km <= 1.5 && S.kvOverride){
      S.kvOverride = false;
      if(kvT) kvT.checked = false;
    } else if(km > 10 && S.kvOverride){
      S.volKvManual = km;
    }
    delete S.kvManual;
  }
  const vg = S.width * S.depth * S.height;
  if(inpVol){
    if(S.kvOverride && Number.isFinite(S.volKvManual) && S.volKvManual > 0){
      inpVol.value = String(Number(S.volKvManual.toFixed(1)));
      inpVol.readOnly = false;
    } else if(S.kvOverride && vg > 0){
      inpVol.readOnly = false;
      if(!inpVol.value.trim()) inpVol.value = String(Number(vg.toFixed(1)));
    } else {
      inpVol.readOnly = true;
      if(vg > 0) inpVol.value = String(Number(vg.toFixed(1)));
    }
  }
  readKvOverrideFromUi();
}

// Browser fallback cache (used only if local server persistence is unavailable)
function saveProjectBrowserFallback(name, state){
  const projects = JSON.parse(localStorage.getItem('sw_projects') || '[]');
  const existing = currentProjectId ? projects.findIndex(p => p.id === currentProjectId) : -1;
  const now = new Date().toISOString();
  if(existing >= 0){
    projects[existing].name = name;
    projects[existing].state = state;
    projects[existing].updated = now;
    projects[existing].location = state.address || '';
  } else {
    const id = 'local_' + Date.now();
    projects.push({ id, name, state, location: state.address || '', created: now, updated: now });
    currentProjectId = id;
  }
  localStorage.setItem('sw_projects', JSON.stringify(projects));
}

function getBrowserFallbackProjects(){
  return JSON.parse(localStorage.getItem('sw_projects') || '[]');
}

function deleteBrowserFallbackProject(id){
  const projects = getBrowserFallbackProjects().filter(p => p.id !== id);
  localStorage.setItem('sw_projects', JSON.stringify(projects));
}

async function saveProjectLocal(name, state){
  try {
    const response = await fetch('/api/save-project', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, projectId: currentProjectId, state })
    });
    const data = await response.json();
    if(!response.ok) throw new Error(data.error || 'Local save failed');
    currentProjectId = data.projectId;
    saveProjectBrowserFallback(name, state);
    return { source: 'server', projectId: currentProjectId };
  } catch(err) {
    console.warn('Local server save failed, using browser fallback:', err);
    saveProjectBrowserFallback(name, state);
    return { source: 'browser-fallback', projectId: currentProjectId };
  }
}

async function getLocalProjects(){
  try {
    const response = await fetch('/api/list-projects', { method: 'GET' });
    const data = await response.json();
    if(!response.ok) throw new Error(data.error || 'Local project list failed');
    return (data.projects || []).map(function(p){
      return {
        ...p,
        location: p.location || p.state?.address || '',
        updated: p.updatedAt || p.updated || p.created || null
      };
    });
  } catch(err){
    console.warn('Local server list failed, using browser fallback:', err);
    return getBrowserFallbackProjects();
  }
}

async function loadLocalProject(projectId){
  try {
    const response = await fetch('/api/load-project?id=' + encodeURIComponent(projectId), { method: 'GET' });
    const data = await response.json();
    if(!response.ok) throw new Error(data.error || 'Failed to load local project');
    if(data.project) return data.project;
    return data;
  } catch(err){
    console.warn('Local server load failed, using browser fallback:', err);
    return getBrowserFallbackProjects().find(p => p.id === projectId) || null;
  }
}

async function deleteLocalProject(id){
  try {
    const response = await fetch('/api/delete-project', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ projectId: id })
    });
    if(!response.ok){
      const data = await response.json().catch(function(){ return {}; });
      throw new Error(data.error || 'Local delete failed');
    }
  } catch(err){
    console.warn('Local server delete failed, continuing with browser fallback:', err);
  }
  deleteBrowserFallbackProject(id);
}

async function saveProject(){
  const name = prompt('Project name:', currentProjectName);
  if(!name) return;
  currentProjectName = name;
  const state = getProjectState();

  const saved = await saveProjectLocal(name, state);
  toast(saved.source === 'server' ? 'Project saved locally: ' + name : 'Project saved in browser fallback: ' + name);

  currentProjectShared = false;
}

async function openSavedProjects(){
  showSavedProjectsOverlay(await getLocalProjects());
}

function showSavedProjectsOverlay(projects){
  let overlay = document.getElementById('saved-projects-overlay');
  if(overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'saved-projects-overlay';
  overlay.className = 'modal-overlay show';
  overlay.style.zIndex = '8500';

  let rows = '';
  if(projects.length === 0){
    rows = '<tr><td colspan="4" style="text-align:center;color:var(--text3);padding:20px">No saved projects yet. Click "Save Project" in the Detailed Report tab.</td></tr>';
  } else {
    projects.forEach(p => {
      rows += `<tr>
        <td style="font-weight:600">${p.name || 'Untitled'}</td>
        <td style="font-size:10px;color:var(--text2)">${p.location || '—'}</td>
        <td style="font-size:10px">${new Date(p.updated || p.created || Date.now()).toLocaleDateString()}</td>
        <td style="white-space:nowrap">
          <button onclick="loadSavedProject('${p.id}')" style="background:var(--primary);color:var(--bg);border:none;padding:4px 10px;border-radius:4px;font-size:10px;cursor:pointer;margin-right:4px">Open</button>
          <button onclick="deleteSavedProject('${p.id}','${(p.name||'').replace(/'/g,"\\'")}'); " style="background:var(--danger);color:#fff;border:none;padding:4px 10px;border-radius:4px;font-size:10px;cursor:pointer">Delete</button>
        </td>
      </tr>`;
    });
  }

  overlay.innerHTML = `
    <div class="modal-card" style="max-width:650px">
      <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">✕</button>
      <h2 style="color:var(--primary);margin-bottom:12px">📂 My Projects</h2>
      <p style="color:var(--text2);font-size:12px;margin-bottom:12px">${projects.length} saved project${projects.length!==1?'s':''}</p>
      <table class="result-table" style="width:100%">
        <thead><tr><th>Name</th><th>Location</th><th>Updated</th><th></th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;

  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if(e.target === overlay) overlay.remove(); });
}

async function loadSavedProject(projectId){
  const local = await loadLocalProject(projectId);
  if(local && local.state){
    Object.assign(S, local.state);
    mergeProjectStateDefaults(local.state);
    if(typeof S.vrManual !== 'boolean') S.vrManual = true;
    currentProjectId = projectId;
    currentProjectName = local.name || 'Untitled';
    pushStateToUI();
    onInput();
    toast('Loaded: ' + currentProjectName);
    const overlay = document.getElementById('saved-projects-overlay');
    if(overlay) overlay.remove();
    return;
  }
  toast('Project not found in local storage.');
}

async function deleteSavedProject(projectId, name){
  if(!confirm('Delete project "' + name + '"?')) return;

  await deleteLocalProject(projectId);

  toast('Deleted: ' + name);
  openSavedProjects(); // Refresh
}

// ═══════════════════════════════════════════════════
//   ACTIVITY LOG (Team/Enterprise)
// ═══════════════════════════════════════════════════
async function openActivityLog(){
  toast('Shared activity log is not available in this local build yet.');
}

function showActivityLogOverlay(entries){
  let overlay = document.getElementById('activity-log-overlay');
  if(overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'activity-log-overlay';
  overlay.className = 'modal-overlay show';
  overlay.style.zIndex = '8500';

  let rows = '';
  if(entries.length === 0){
    rows = '<tr><td colspan="5" style="text-align:center;color:var(--text3);padding:20px">No activity yet</td></tr>';
  } else {
    entries.forEach(e => {
      const actionIcons = { created: '+', updated: '~', deleted: 'x', shared: '>', invited: '+' };
      const actionColors = { created: 'var(--success,#27ae60)', updated: 'var(--primary)', deleted: 'var(--danger,#e74c3c)', shared: '#8e44ad', invited: '#2980b9' };
      const icon = actionIcons[e.action] || '-';
      const color = actionColors[e.action] || 'var(--text2)';
      const userShort = (e.user || 'Unknown').split('@')[0];
      const changes = e.changes ? e.changes.map(c => `<span style="font-size:10px;background:var(--surface2);padding:1px 6px;border-radius:3px;margin:1px;display:inline-block">${c}</span>`).join(' ') : '';
      rows += `<tr>
        <td style="white-space:nowrap;font-size:10px;color:var(--text3)">${new Date(e.timestamp).toLocaleString()}</td>
        <td style="font-weight:600;font-size:12px">${userShort}</td>
        <td><span style="display:inline-block;width:18px;height:18px;line-height:18px;text-align:center;border-radius:4px;background:${color};color:#fff;font-size:10px;font-weight:bold">${icon}</span> <span style="color:${color};font-weight:600;text-transform:capitalize;font-size:12px">${e.action || 'unknown'}</span></td>
        <td style="font-size:11px">${e.details || 'No details'}</td>
        <td style="font-size:10px">${changes}</td>
      </tr>`;
    });
  }

  overlay.innerHTML = `
    <div class="modal-card" style="max-width:850px">
      <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">X</button>
      <h2 style="color:var(--primary);margin-bottom:4px">Activity Log</h2>
      <p style="color:var(--text2);font-size:12px;margin-bottom:12px">Shared-project activity is not included in this local build.</p>
      <div style="max-height:450px;overflow-y:auto">
        <table class="result-table" style="width:100%">
          <thead><tr><th style="width:130px">Time</th><th style="width:80px">User</th><th style="width:100px">Action</th><th>Details</th><th>Changes</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>`;

  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if(e.target === overlay) overlay.remove(); });
}

// ═══════════════════════════════════════════════════
//   API KEY MANAGEMENT (Team/Enterprise)
// ═══════════════════════════════════════════════════
async function openApiKeys(){
  toast('API key management is not part of this local build yet.');
}

function showApiKeysOverlay(keys, usage){
  let overlay = document.getElementById('api-keys-overlay');
  if(overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'api-keys-overlay';
  overlay.className = 'modal-overlay show';
  overlay.style.zIndex = '8500';

  const plan = userSubscription.plan;
  const limit =
    typeof usage.limit === 'number' && usage.limit > 0
      ? usage.limit
      : plan === 'enterprise'
        ? 5000
        : 30;
  const used = usage.count || 0;
  const pct = Math.min((used/limit)*100, 100);

  let keyRows = '';
  if(keys.length === 0){
    keyRows = '<div style="color:var(--text3);padding:12px;text-align:center">No API keys yet</div>';
  } else {
    keys.forEach(k => {
      keyRows += `<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border)">
        <code style="flex:1;font-size:11px;background:var(--surface2);padding:4px 8px;border-radius:4px;color:var(--primary)">${k.key}</code>
        <span style="font-size:10px;color:var(--text2)">${k.name || 'Default'}</span>
        <span style="font-size:10px;color:var(--text3)">${new Date(k.created).toLocaleDateString()}</span>
        <button onclick="revokeApiKey('${k.id}')" style="background:var(--danger);color:#fff;border:none;padding:3px 8px;border-radius:4px;font-size:10px;cursor:pointer">Revoke</button>
      </div>`;
    });
  }

  overlay.innerHTML = `
    <div class="modal-card" style="max-width:650px">
      <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">✕</button>
      <h2 style="color:var(--primary);margin-bottom:12px">🔑 API Keys</h2>
      <div style="margin-bottom:16px">
        <div style="display:flex;justify-content:space-between;font-size:11px;color:var(--text2);margin-bottom:4px">
          <span>API Usage This Month</span>
          <span>${used.toLocaleString()} / ${limit.toLocaleString()} requests</span>
        </div>
        <div style="height:8px;background:var(--surface3);border-radius:4px;overflow:hidden">
          <div style="height:100%;width:${pct}%;background:${pct>80?'var(--danger)':'var(--primary)'};border-radius:4px;transition:width .3s"></div>
        </div>
      </div>
      <div style="margin-bottom:12px">
        <h4 style="color:var(--text2);margin-bottom:8px">Your Keys</h4>
        ${keyRows}
      </div>
      <button onclick="generateApiKey()" style="background:var(--primary);color:var(--bg);border:none;padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">+ Generate New Key</button>
      <div style="margin-top:16px;padding-top:12px;border-top:1px solid var(--border)">
        <h4 style="color:var(--text2);margin-bottom:6px">Quick Start</h4>
        <pre style="background:var(--surface2);padding:10px;border-radius:6px;font-size:11px;overflow-x:auto;color:var(--text)">curl -X POST https://example.local/api/v1/wind-pressure \\
  -H "Authorization: Bearer YOUR_API_KEY" \\
  -H "Content-Type: application/json" \\
  -d '{"lat":-36.85,"lng":174.76,"width":20,"depth":15,"height":6}'</pre>
      </div>
    </div>`;

  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if(e.target === overlay) overlay.remove(); });
}

async function generateApiKey(){
  toast('API key generation is not implemented in this local build yet.');
}

async function revokeApiKey(keyId){
  toast('API key revocation is not implemented in this local build yet.');
}

// ═══════════════════════════════════════════════════
//   TEMPLATES (Enterprise)
// ═══════════════════════════════════════════════════
async function openTemplates(){
  showTemplatesOverlay([]);
}

function showTemplatesOverlay(templates){
  let overlay = document.getElementById('templates-overlay');
  if(overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'templates-overlay';
  overlay.className = 'modal-overlay show';
  overlay.style.zIndex = '8500';

  let rows = '';
  // Built-in templates
  const builtIn = [
    { id: '_warehouse', name: 'Warehouse 30×60', desc: '30m×60m×8m, Gable 5°, TC2.5', preset: { width:30, depth:60, height:8, pitch:5, roofType:'gable', terrainCat:2.5 }},
    { id: '_office2', name: '2-Storey Office', desc: '15m×20m×7m, Hip 15°, TC3', preset: { width:15, depth:20, height:7, pitch:15, roofType:'hip', terrainCat:3 }},
    { id: '_shed', name: 'Farm Shed', desc: '12m×20m×4.5m, Monoslope 10°, TC1', preset: { width:12, depth:20, height:4.5, pitch:10, roofType:'monoslope', terrainCat:1 }},
    { id: '_retail', name: 'Retail Building', desc: '20m×30m×5m, Flat, TC3', preset: { width:20, depth:30, height:5, pitch:0, roofType:'flat', terrainCat:3 }},
  ];

  rows += '<h4 style="color:var(--text2);margin:8px 0 6px">Built-in Templates</h4>';
  builtIn.forEach(t => {
    rows += `<div style="display:flex;align-items:center;gap:8px;padding:8px;background:var(--surface);border:1px solid var(--border);border-radius:6px;margin-bottom:6px">
      <div style="flex:1"><b style="font-size:12px">${t.name}</b><br><span style="font-size:10px;color:var(--text2)">${t.desc}</span></div>
      <button onclick='applyTemplate(${JSON.stringify(t.preset)})' style="background:var(--primary);color:var(--bg);border:none;padding:5px 12px;border-radius:4px;font-size:11px;cursor:pointer">Apply</button>
    </div>`;
  });

  if(templates.length > 0){
    rows += '<h4 style="color:var(--text2);margin:12px 0 6px">Custom Templates</h4>';
    templates.forEach(t => {
      rows += `<div style="display:flex;align-items:center;gap:8px;padding:8px;background:var(--surface);border:1px solid var(--border);border-radius:6px;margin-bottom:6px">
        <div style="flex:1"><b style="font-size:12px">${t.name}</b><br><span style="font-size:10px;color:var(--text2)">${t.description || ''}</span></div>
        <button onclick='applyTemplate(${JSON.stringify(t.preset)})' style="background:var(--primary);color:var(--bg);border:none;padding:5px 12px;border-radius:4px;font-size:11px;cursor:pointer">Apply</button>
        <button onclick="deleteTemplate('${t.id}')" style="background:var(--danger);color:#fff;border:none;padding:5px 8px;border-radius:4px;font-size:10px;cursor:pointer">✕</button>
      </div>`;
    });
  }

  overlay.innerHTML = `
    <div class="modal-card" style="max-width:550px">
      <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">✕</button>
      <h2 style="color:var(--primary);margin-bottom:12px">📋 Templates</h2>
      <p style="color:var(--text2);font-size:12px;margin-bottom:12px">Quickly apply standard building configurations</p>
      <div style="max-height:400px;overflow-y:auto">${rows}</div>
      <div style="margin-top:12px;padding-top:12px;border-top:1px solid var(--border)">
        <button onclick="saveAsTemplate()" style="background:var(--surface3);border:1px solid var(--border);color:var(--text);padding:8px 16px;border-radius:6px;font-size:12px;cursor:pointer">+ Save Current as Template</button>
      </div>
    </div>`;

  document.body.appendChild(overlay);
  overlay.addEventListener('click', e => { if(e.target === overlay) overlay.remove(); });
}

function applyTemplate(preset){
  Object.assign(S, preset);
  mergeProjectStateDefaults(preset);
  onInput();
  toast('Template applied');
  // Close overlay
  const overlay = document.getElementById('templates-overlay');
  if(overlay) overlay.remove();
}

async function saveAsTemplate(){
  toast('Custom template saving is not implemented in this local build yet.');
}

async function deleteTemplate(templateId){
  toast('Custom templates are not implemented in this local build yet.');
}

// ═══════════════════════════════════════════════
//   3D MODEL UPLOAD & VIEWER
// ═══════════════════════════════════════════════

/** Single CDN base for web-ifc JS + WASM (keep in sync). */
const WEB_IFC_CDN = 'https://cdn.jsdelivr.net/npm/web-ifc@0.0.77/';

function modelFileExt(filename){
  const n = String(filename || '').trim().toLowerCase();
  const dot = n.lastIndexOf('.');
  if(dot < 0) return '';
  return n.slice(dot + 1);
}

function handleModelUpload(event){
  const file = event.target.files ? event.target.files[0] : event;
  if(!file) return;
  const input = event.target && event.target.files ? event.target : null;
  void handleModelFile(file).catch(function(err){
    console.error('handleModelFile:', err);
    showModelStatus('error', err && err.message ? err.message : 'Upload failed');
    toast('⚠ ' + (err && err.message ? err.message : 'Could not load model'));
  }).finally(function(){
    try{
      if(input && input.value) input.value = '';
    }catch(e){}
  });
}

async function handleModelFile(file){
  const ext = modelFileExt(file.name);
  const supported = ['glb','gltf','obj','fbx','ifc','rvt'];
  if(!supported.includes(ext)){
    toast('⚠ Unsupported format. Use: ' + supported.join(', '));
    return;
  }
  if(ext === 'rvt'){
    uploadedModelName = file.name;
    showModelStatus('loading', 'Checking file…');
    toast('Revit .rvt cannot be opened in the browser. Export IFC or GLB from Revit, then upload that file.');
    handleRvtUpload(file);
    return;
  }

  uploadedModelName = file.name;
  showModelStatus('loading', 'Loading ' + file.name + '…');

  if(ext === 'ifc'){
    await handleIfcUpload(file);
    return;
  }

  const reader = new FileReader();
  reader.onload = function(e){
    const buffer = e.target.result;
    try {
      if(ext === 'glb') loadGLB(buffer, file.name);
      else if(ext === 'gltf') loadGLTFJson(e.target.result, file.name);
      else if(ext === 'obj') loadOBJ(e.target.result, file.name);
      else if(ext === 'fbx') loadFBX(buffer, file.name);
    } catch(err){
      console.error('Model load error:', err);
      showModelStatus('error', 'Failed to load: ' + err.message);
      toast('⚠ Failed to load model');
    }
  };
  if(ext === 'obj') reader.readAsText(file);
  else if(ext === 'gltf') reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
}

// ── IFC Dropzone drag-and-drop ──
function initIfcDropzone(){
  const dz = document.getElementById('ifc-dropzone');
  if(!dz) return;
  ['dragenter','dragover'].forEach(evt => {
    dz.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); dz.classList.add('dragover'); });
  });
  ['dragleave','drop'].forEach(evt => {
    dz.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); dz.classList.remove('dragover'); });
  });
  dz.addEventListener('drop', e => {
    const files = e.dataTransfer.files;
    if(files.length > 0){
      void handleModelFile(files[0]).catch(function(err){
        console.error('handleModelFile (drop):', err);
        showModelStatus('error', err && err.message ? err.message : 'Upload failed');
        toast('⚠ ' + (err && err.message ? err.message : 'Could not load model'));
      });
    }
  });
}
setTimeout(initIfcDropzone, 300);
setTimeout(() => {
  initIfcAiAssistFromStorage();
  updateIfcAiAssistUI();
}, 320);

function loadGLB(buffer, name){
  const loader = new THREE.GLTFLoader();
  loader.parse(buffer, '', function(gltf){
    addUploadedModel(gltf.scene, name);
  }, function(err){
    console.error('GLTF parse error:', err);
    showModelStatus('error', 'Failed to parse GLB');
    toast('⚠ Failed to parse GLB file');
  });
}

function loadGLTFJson(text, name){
  const loader = new THREE.GLTFLoader();
  const json = JSON.parse(text);
  // Embedded GLTF (no external resources) — parse as buffer
  const blob = new Blob([text], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  loader.load(url, function(gltf){
    URL.revokeObjectURL(url);
    addUploadedModel(gltf.scene, name);
  }, undefined, function(err){
    URL.revokeObjectURL(url);
    console.error('GLTF load error:', err);
    showModelStatus('error', 'Failed to parse GLTF');
    toast('⚠ Failed to parse GLTF file');
  });
}

function loadOBJ(text, name){
  const loader = new THREE.OBJLoader();
  const obj = loader.parse(text);
  addUploadedModel(obj, name);
}

function loadFBX(buffer, name){
  if(THREE.FBXLoader){
    const loader = new THREE.FBXLoader();
    const fbx = loader.parse(buffer, '');
    addUploadedModel(fbx, name);
    return;
  }
  // Dynamically load fflate + FBXLoader
  showModelStatus('loading', 'Loading FBX parser…');
  const s1 = document.createElement('script');
  s1.src = 'https://cdn.jsdelivr.net/npm/fflate@0.6.9/umd/index.js';
  s1.onload = () => {
    const s2 = document.createElement('script');
    s2.src = 'https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/loaders/FBXLoader.js';
    s2.onload = () => {
      try {
        const loader = new THREE.FBXLoader();
        const fbx = loader.parse(buffer, '');
        addUploadedModel(fbx, name);
      } catch(err){
        console.error('FBX parse error:', err);
        showModelStatus('error', 'FBX parse failed — try GLB/OBJ export');
        toast('⚠ FBX parsing failed: ' + err.message);
      }
    };
    s2.onerror = () => {
      showModelStatus('error', 'FBX loader unavailable — try GLB/OBJ');
      toast('⚠ Could not load FBX parser. Export as .glb or .obj instead.');
    };
    document.head.appendChild(s2);
  };
  s1.onerror = () => {
    showModelStatus('error', 'FBX dependencies unavailable');
    toast('⚠ Could not load FBX dependencies. Export as .glb or .obj instead.');
  };
  document.head.appendChild(s1);
}

let webIfcScriptPromise = null;
async function ensureWebIfcScriptLoaded(){
  if(typeof WebIFC !== 'undefined') return;
  if(!webIfcScriptPromise){
    webIfcScriptPromise = new Promise(function(resolve, reject){
      const script = document.createElement('script');
      script.src = WEB_IFC_CDN + 'web-ifc-api-iife.js';
      script.setAttribute('data-web-ifc-api', '1');
      script.onload = function(){ resolve(); };
      script.onerror = function(){ reject(new Error('Could not load web-ifc from CDN')); };
      document.head.appendChild(script);
    });
  }
  await webIfcScriptPromise;
  if(typeof WebIFC === 'undefined') throw new Error('web-ifc loaded but WebIFC global is missing');
}

/** Must be async and awaited so handleModelUpload’s finally clears the file input only after parse finishes. */
async function handleIfcUpload(file){
  showModelStatus('loading', 'Loading IFC parser…');
  await ensureWebIfcScriptLoaded();
  await parseIfcFile(file);
}

/**
 * web-ifc vertex buffers are usually 6 floats (xyz + nxyz), but newer builds may use 8 (e.g. + UVs).
 * Mis-reading stride blows up bounding boxes → tiny scale → invisible meshes in the viewer.
 */
function inferIfcVertexStride(verts, indices){
  const n = verts.length;
  if(n < 6) return 6;
  const candidates = [];
  for(const s of [6, 8, 10, 12]){
    if(n % s === 0 && n >= s) candidates.push(s);
  }
  if(!candidates.length) return 6;
  let maxIdx = 0;
  if(indices && indices.length){
    for(let i = 0; i < indices.length; i++){
      const v = indices[i];
      if(v > maxIdx) maxIdx = v;
    }
    const ok = candidates.filter(s => maxIdx < n / s - 1e-6);
    if(ok.length) return ok.sort((a, b) => a - b)[0];
    // Index buffer may not match any stride (exotic layout) — prefer smallest stride so verts exist
    if(candidates.length) return candidates.sort((a, b) => a - b)[0];
  }
  if(n % 8 === 0 && n % 6 !== 0) return 8;
  if(n % 10 === 0 && n % 6 !== 0 && n % 8 !== 0) return 10;
  return 6;
}

function ifcPlacedMatrixToThree(flatT){
  const m = new THREE.Matrix4();
  if(!flatT || !flatT.length) return m.identity();
  if(flatT.length >= 16){
    m.fromArray(flatT);
  } else if(flatT.length >= 12){
    // 3×4 row-major: [ R | t ] with implicit last row 0,0,0,1 (common when not 16 floats).
    m.set(
      flatT[0], flatT[1], flatT[2], flatT[3],
      flatT[4], flatT[5], flatT[6], flatT[7],
      flatT[8], flatT[9], flatT[10], flatT[11],
      0, 0, 0, 1
    );
  } else {
    return m.identity();
  }
  const e = m.elements;
  for(let i = 0; i < 16; i++){
    if(!Number.isFinite(e[i])) return new THREE.Matrix4().identity();
  }
  return m;
}

/** Archicad/Revit often export near-black greys — invisible on dark UI + MeshStandard needs light. */
function brightenIfcAlbedo(color){
  const c = color.clone();
  c.r = Math.min(1, c.r); c.g = Math.min(1, c.g); c.b = Math.min(1, c.b);
  const lum = 0.2126 * c.r + 0.7152 * c.g + 0.0722 * c.b;
  if(lum < 0.22){
    c.lerp(new THREE.Color(0.58, 0.6, 0.63), 0.72);
  }
  return c;
}

/**
 * Proven path from 30.03.2026 archive: Init() + default OpenModel + LoadAllGeometry only,
 * fixed stride-6 vertices, index as BufferAttribute(itemSize 1), Matrix4.fromArray(placement).
 * Later refactors (COORDINATE_TO_ORIGIN first, geom.delete, inferred stride) broke visibility for some IFCs.
 */
function buildIfcGroupLikeArchive(ifcApi, modelID){
  const group = new THREE.Group();
  const meshes = ifcApi.LoadAllGeometry(modelID);
  for(let i = 0; i < meshes.size(); i++){
    const mesh = meshes.get(i);
    const placedGeometries = mesh.geometries;
    const ifcExpressID = (typeof mesh.expressID === 'number' && mesh.expressID > 0)
      ? mesh.expressID
      : null;

    for(let j = 0; j < placedGeometries.size(); j++){
      const pg = placedGeometries.get(j);
      const geomData = ifcApi.GetGeometry(modelID, pg.geometryExpressID);
      const verts = ifcApi.GetVertexArray(geomData.GetVertexData(), geomData.GetVertexDataSize());
      const indices = ifcApi.GetIndexArray(geomData.GetIndexData(), geomData.GetIndexDataSize());

      const vertCount = verts.length / 6;
      if(vertCount < 1 || !Number.isFinite(vertCount)) continue;

      const positions = new Float32Array(vertCount * 3);
      const normals = new Float32Array(vertCount * 3);
      for(let k = 0; k < vertCount; k++){
        positions[k * 3]     = verts[k * 6];
        positions[k * 3 + 1] = verts[k * 6 + 1];
        positions[k * 3 + 2] = verts[k * 6 + 2];
        normals[k * 3]       = verts[k * 6 + 3];
        normals[k * 3 + 1]   = verts[k * 6 + 4];
        normals[k * 3 + 2]   = verts[k * 6 + 5];
      }

      const geometry = new THREE.BufferGeometry();
      geometry.setAttribute('position', new THREE.BufferAttribute(positions, 3));
      geometry.setAttribute('normal', new THREE.BufferAttribute(normals, 3));
      if(indices && indices.length){
        geometry.setIndex(new THREE.BufferAttribute(indices, 1));
      } else {
        geometry.computeVertexNormals();
      }

      const color = new THREE.Color(pg.color.x, pg.color.y, pg.color.z);
      const mat = new THREE.MeshStandardMaterial({
        color, transparent: pg.color.w < 1, opacity: pg.color.w,
        side: THREE.DoubleSide, roughness: 0.6
      });

      const m = ifcPlacedMatrixToThree(pg.flatTransformation);
      const threeMesh = new THREE.Mesh(geometry, mat);
      threeMesh.applyMatrix4(m);
      if(ifcExpressID != null){
        threeMesh.userData.ifcExpressID = ifcExpressID;
        try {
          if(typeof ifcApi.GetLineType === 'function'){
            threeMesh.userData.ifcLineType = ifcApi.GetLineType(modelID, ifcExpressID);
          }
        } catch(e) { /* ignore */ }
      }
      group.add(threeMesh);
    }
  }
  return group;
}

function addIfcPlacedGeometryToGroup(ifcApi, modelID, pg, ifcExpressID, group){
  const geomData = ifcApi.GetGeometry(modelID, pg.geometryExpressID);
  if(!geomData){
    return;
  }
  const verts = ifcApi.GetVertexArray(geomData.GetVertexData(), geomData.GetVertexDataSize());
  const indices = ifcApi.GetIndexArray(geomData.GetIndexData(), geomData.GetIndexDataSize());

  // web-ifc ships xyz+nxnynz (stride 6) by default; infer only when length does not fit 6.
  let stride = 6;
  if(!verts || verts.length < 6){
    return;
  }
  if(verts.length % 6 !== 0){
    stride = inferIfcVertexStride(verts, indices);
  }
  const vertCount = Math.floor(verts.length / stride);
  if(vertCount < 1) return;

  const positions = new Float32Array(vertCount * 3);
  const normals = new Float32Array(vertCount * 3);
  for(let k = 0; k < vertCount; k++){
    const o = k * stride;
    positions[k * 3]     = verts[o];
    positions[k * 3 + 1] = verts[o + 1];
    positions[k * 3 + 2] = verts[o + 2];
    normals[k * 3]       = verts[o + 3];
    normals[k * 3 + 1]   = verts[o + 4];
    normals[k * 3 + 2]   = verts[o + 5];
  }

  const geometry = new THREE.BufferGeometry();
  geometry.setAttribute('position', new THREE.BufferAttribute(positions, 3));
  geometry.setAttribute('normal', new THREE.BufferAttribute(normals, 3));
  if(indices && indices.length){
    geometry.setIndex(indices);
  } else {
    geometry.computeVertexNormals();
  }

  const cx = pg.color && Number.isFinite(pg.color.x) ? pg.color.x : 0.7;
  const cy = pg.color && Number.isFinite(pg.color.y) ? pg.color.y : 0.7;
  const cz = pg.color && Number.isFinite(pg.color.z) ? pg.color.z : 0.7;
  const cw = pg.color && Number.isFinite(pg.color.w) ? pg.color.w : 1;
  const color = brightenIfcAlbedo(new THREE.Color(cx, cy, cz));
  const op = cw < 0.02 ? 1 : Math.min(1, Math.max(0.15, cw));
  const mat = new THREE.MeshStandardMaterial({
    color,
    emissive: color.clone().multiplyScalar(0.2),
    emissiveIntensity: 0.85,
    transparent: op < 0.999,
    opacity: op,
    side: THREE.DoubleSide,
    roughness: 0.65,
    metalness: 0.05,
    depthWrite: true
  });

  const m = ifcPlacedMatrixToThree(pg.flatTransformation);
  const threeMesh = new THREE.Mesh(geometry, mat);
  threeMesh.frustumCulled = false;
  threeMesh.applyMatrix4(m);
  if(ifcExpressID != null){
    threeMesh.userData.ifcExpressID = ifcExpressID;
    try {
      if(typeof ifcApi.GetLineType === 'function'){
        threeMesh.userData.ifcLineType = ifcApi.GetLineType(modelID, ifcExpressID);
      }
    } catch(e2) { /* ignore */ }
  }
  group.add(threeMesh);
}

/** Vector-like objects from web-ifc always expose .size() and .get(i); guard both patterns. */
function ifcVectorLength(v){
  if(!v) return 0;
  if(typeof v.size === 'function') return v.size();
  if(typeof v.size === 'number' && Number.isFinite(v.size)) return v.size;
  return 0;
}

function ifcVectorGet(v, i){
  if(!v || typeof v.get !== 'function') return null;
  return v.get(i);
}

/** Append one web-ifc FlatMesh (stream or LoadAllGeometry vector item) into a THREE.Group. */
function appendIfcFlatMesh(ifcApi, modelID, flatMesh, group){
  if(!flatMesh) return;
  const placedGeometries = flatMesh.geometries;
  const nPg = ifcVectorLength(placedGeometries);
  const rawId = flatMesh.expressID != null ? flatMesh.expressID : flatMesh.expressId;
  const ifcExpressID = (typeof rawId === 'number' && rawId > 0) ? rawId : null;
  for(let j = 0; j < nPg; j++){
    const pg = ifcVectorGet(placedGeometries, j);
    if(pg) addIfcPlacedGeometryToGroup(ifcApi, modelID, pg, ifcExpressID, group);
  }
}

function resolveWebIfcTypeCode(name){
  const W = typeof WebIFC !== 'undefined' ? WebIFC : null;
  if(!W || W[name] == null) return null;
  const v = W[name];
  if(typeof v === 'number' && Number.isFinite(v)) return v;
  if(typeof v === 'function'){
    try{
      const r = v();
      return typeof r === 'number' && Number.isFinite(r) ? r : null;
    }catch(e){
      return null;
    }
  }
  return null;
}

/**
 * Resolve after StreamAllMeshes returns using setTimeout(0), not queueMicrotask.
 * WASM / emscripten often schedules follow-up work as microtasks; resolving in the same
 * microtask queue can race and leave group empty before CloseModel (especially with MT WASM).
 */
function ifcRunStreamAllMeshes(ifcApi, modelID, group){
  return new Promise(function(resolve, reject){
    if(typeof ifcApi.StreamAllMeshes !== 'function'){ resolve(); return; }
    try{
      ifcApi.StreamAllMeshes(modelID, function(flatMesh, index, total){
        appendIfcFlatMesh(ifcApi, modelID, flatMesh, group);
      });
    }catch(e){
      reject(e);
      return;
    }
    setTimeout(resolve, 0);
  });
}

function ifcRunStreamAllMeshesWithTypes(ifcApi, modelID, types, group){
  return new Promise(function(resolve, reject){
    if(typeof ifcApi.StreamAllMeshesWithTypes !== 'function'){ resolve(); return; }
    try{
      ifcApi.StreamAllMeshesWithTypes(modelID, types, function(flatMesh, index, total){
        appendIfcFlatMesh(ifcApi, modelID, flatMesh, group);
      });
    }catch(e){
      reject(e);
      return;
    }
    setTimeout(resolve, 0);
  });
}

function waitTwoAnimationFrames(){
  return new Promise(function(resolve){
    requestAnimationFrame(function(){ requestAnimationFrame(resolve); });
  });
}

/** Fallback: stream only common building-element IFC types (some files omit generic StreamAllMeshes). */
async function appendIfcStreamedTypes(ifcApi, modelID, group){
  if(typeof ifcApi.StreamAllMeshesWithTypes !== 'function') return;
  const typeNames = [
    'IFCWALLSTANDARDCASE', 'IFCWALL', 'IFCSLAB', 'IFCWINDOW', 'IFCDOOR', 'IFCROOF', 'IFCCOLUMN', 'IFCBEAM',
    'IFCBUILDINGELEMENTPROXY', 'IFCPLATE', 'IFCMEMBER', 'IFCSTAIR', 'IFCRAILING', 'IFCCURTAINWALL', 'IFCCOVERING',
    'IFCFOOTING', 'IFCBUILDINGELEMENTPART', 'IFCCHIMNEY', 'IFCRAMP', 'IFCSHADINGDEVICE'
  ];
  const types = [];
  for(let i = 0; i < typeNames.length; i++){
    const c = resolveWebIfcTypeCode(typeNames[i]);
    if(c != null) types.push(c);
  }
  if(!types.length) return;
  try{
    await ifcRunStreamAllMeshesWithTypes(ifcApi, modelID, types, group);
  }catch(e){
    console.warn('IFC StreamAllMeshesWithTypes:', e);
  }
}

/**
 * Last resort: GetFlatMesh per IFC product express ID (slow; capped).
 */
function appendIfcFlatMeshesFromProducts(ifcApi, modelID, group){
  if(typeof ifcApi.GetLineIDsWithType !== 'function' || typeof ifcApi.GetFlatMesh !== 'function') return;
  const productType = resolveWebIfcTypeCode('IFCPRODUCT');
  if(productType == null) return;
  let ids;
  try{
    ids = ifcApi.GetLineIDsWithType(modelID, productType, true);
  }catch(e){
    return;
  }
  const n = ifcVectorLength(ids);
  const maxScan = 6000;
  for(let i = 0; i < Math.min(n, maxScan); i++){
    let eid;
    try{ eid = ifcVectorGet(ids, i); }catch(e){ break; }
    if(typeof eid !== 'number' || eid < 1) continue;
    let flatMesh = null;
    try{
      flatMesh = ifcApi.GetFlatMesh(modelID, eid);
    }catch(e2){
      continue;
    }
    if(flatMesh) appendIfcFlatMesh(ifcApi, modelID, flatMesh, group);
  }
}

/**
 * Prefer LoadAllGeometry (bulk) — same data as streaming but avoids stream-only quirks on some files.
 * Then StreamAllMeshes — typed stream — per-product GetFlatMesh.
 * Do not call flatMesh.delete() inside StreamAllMeshes (can corrupt WASM while iterating).
 */
function appendIfcLoadAllGeometry(ifcApi, modelID, group){
  if(typeof ifcApi.LoadAllGeometry !== 'function') return;
  let meshes;
  try{
    meshes = ifcApi.LoadAllGeometry(modelID);
  }catch(e){
    console.warn('IFC LoadAllGeometry:', e);
    return;
  }
  const nMesh = ifcVectorLength(meshes);
  for(let i = 0; i < nMesh; i++){
    const fm = ifcVectorGet(meshes, i);
    if(fm) appendIfcFlatMesh(ifcApi, modelID, fm, group);
  }
}

async function buildThreeGroupFromIfcModel(ifcApi, modelID){
  const group = new THREE.Group();
  appendIfcLoadAllGeometry(ifcApi, modelID, group);
  if(group.children.length === 0 && typeof ifcApi.StreamAllMeshes === 'function'){
    try{
      await ifcRunStreamAllMeshes(ifcApi, modelID, group);
    } catch(streamErr){
      console.warn('IFC StreamAllMeshes:', streamErr);
    }
    await waitTwoAnimationFrames();
  }
  if(group.children.length === 0){
    appendIfcLoadAllGeometry(ifcApi, modelID, group);
  }
  if(group.children.length === 0){
    await appendIfcStreamedTypes(ifcApi, modelID, group);
  }
  if(group.children.length === 0){
    appendIfcFlatMeshesFromProducts(ifcApi, modelID, group);
  }
  return group;
}

async function parseIfcFile(file){
  let ifcApi = null;
  let modelID = -1;
  try {
    showModelStatus('loading', 'Parsing IFC model…');
    const buffer = await file.arrayBuffer();
    const u8 = new Uint8Array(buffer);
    ifcApi = new WebIFC.IfcAPI();
    ifcApi.SetWasmPath(WEB_IFC_CDN, true);
    // Same as 30.03.2026 archive: default Init() lets web-ifc pick ST vs MT from crossOriginIsolated.
    await ifcApi.Init();

    function closeIfOpen(){
      if(ifcApi && typeof modelID === 'number' && modelID >= 0){
        try{ ifcApi.CloseModel(modelID); }catch(e){}
        modelID = -1;
      }
    }

    /**
     * Default OpenModel(u8) → COORDINATE_TO_ORIGIN false (web-ifc default). The old working build never
     * opened with true first; doing so can change tessellation and left the viewer empty for some files.
     */
    async function openAndBuild(settings){
      modelID = settings === undefined ? ifcApi.OpenModel(u8) : ifcApi.OpenModel(u8, settings);
      if(typeof modelID !== 'number' || modelID < 0){
        return null;
      }
      let group = buildIfcGroupLikeArchive(ifcApi, modelID);
      if(group.children.length === 0){
        group = await buildThreeGroupFromIfcModel(ifcApi, modelID);
      }
      closeIfOpen();
      return group;
    }

    let group = await openAndBuild(undefined);
    if(!group || group.children.length === 0){
      group = await openAndBuild({ COORDINATE_TO_ORIGIN: true });
      if(group && group.children.length > 0){
        toast('IFC loaded using world coordinates (re-centered in viewer).');
      }
    }

    if(!group || group.children.length === 0){
      throw new Error('No triangulated geometry in this IFC — export IFC with shape tessellation from Revit/Archicad, or use GLB/OBJ.');
    }
    addUploadedModel(group, file.name, { unlitFallback: true });
  } catch(err){
    console.error('IFC parse error:', err);
    showModelStatus('error', 'IFC parse failed — try GLB/OBJ export');
    toast('⚠ IFC parsing failed: ' + (err && err.message ? err.message : String(err)));
  } finally {
    if(ifcApi && typeof modelID === 'number' && modelID >= 0){
      try{ ifcApi.CloseModel(modelID); }catch(e){}
    }
  }
}

function handleRvtUpload(file){
  // Revit .rvt is a proprietary format that can't be parsed client-side.
  // Show the guide immediately — no network request needed.
  showRvtSetupGuide();
}

async function pollRvtTranslation(urn, filename, attempt=0){
  if(attempt > 60){
    showModelStatus('error', 'Translation timed out');
    toast('⚠ Autodesk translation timed out');
    return;
  }
  try {
    const response = await fetch('/api/rvt-status?urn=' + encodeURIComponent(urn));
    const data = await response.json();

    if(data.status === 'success' && data.objUrl){
      showModelStatus('loading', 'Downloading translated model…');
      loadTranslatedObj(data.objUrl, filename);
    } else if(data.status === 'failed'){
      showModelStatus('error', 'Autodesk translation failed');
      toast('⚠ Model translation failed');
    } else {
      // Still processing — poll again in 5 seconds
      const pct = data.progress || '…';
      showModelStatus('loading', `Translating: ${pct}`);
      setTimeout(() => pollRvtTranslation(urn, filename, attempt + 1), 5000);
    }
  } catch(err){
    showModelStatus('error', 'Status check failed');
    setTimeout(() => pollRvtTranslation(urn, filename, attempt + 1), 8000);
  }
}

async function loadTranslatedObj(url, filename){
  try {
    const response = await fetch(url);
    const text = await response.text();
    loadOBJ(text, filename.replace('.rvt','.obj'));
  } catch(err){
    showModelStatus('error', 'Failed to download translated model');
    toast('⚠ Failed to download model');
  }
}

function showRvtSetupGuide(){
  showModelStatus('', '');  // clear status bar — the modal is enough
  // Remove any existing overlay
  const existing = document.getElementById('rvt-setup-overlay');
  if(existing) existing.remove();

  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.id = 'rvt-setup-overlay';
  overlay.innerHTML = `
    <div class="modal-panel" style="max-width:560px">
      <h2>🏗️ Revit File Support</h2>

      <div style="background:var(--primary);background:linear-gradient(135deg,#3b82f6,#8b5cf6);padding:14px 16px;border-radius:8px;margin:12px 0;color:white">
        <h4 style="margin:0 0 6px;font-size:14px">✅ Recommended: Export from Revit</h4>
        <p style="margin:0;font-size:12px;opacity:0.95;line-height:1.6">
          In Revit, go to <strong>File → Export</strong> and save as one of these formats that load directly in your browser:
        </p>
        <ul style="margin:8px 0 0;padding-left:18px;font-size:12px;opacity:0.95;line-height:1.8">
          <li><strong>.IFC</strong> — Industry Foundation Classes (best for BIM)</li>
          <li><strong>.GLB</strong> — glTF Binary (best quality & performance)</li>
          <li><strong>.OBJ</strong> — Wavefront OBJ (widely supported)</li>
          <li><strong>.FBX</strong> — Autodesk FBX (3D exchange)</li>
        </ul>
      </div>

      <details style="margin:12px 0">
        <summary style="cursor:pointer;color:var(--text2);font-size:12px;user-select:none">Advanced: Direct .rvt upload via Autodesk APS</summary>
        <div style="background:var(--surface2);padding:12px;border-radius:8px;margin:8px 0;font-size:12px">
          <p style="margin:0 0 8px;color:var(--text2)">To upload .rvt files directly, you need Autodesk Platform Services (paid API):</p>
          <ol style="margin:0;padding-left:20px;line-height:1.8;color:var(--text2)">
            <li>Create an <a href="https://aps.autodesk.com" target="_blank" style="color:var(--primary)">Autodesk APS account</a></li>
            <li>Create an app to get Client ID &amp; Secret</li>
            <li>In Netlify Dashboard → Site Settings → Environment Variables, add:<br>
              <code style="background:var(--bg);padding:2px 6px;border-radius:3px;font-size:11px">APS_CLIENT_ID</code> and
              <code style="background:var(--bg);padding:2px 6px;border-radius:3px;font-size:11px">APS_CLIENT_SECRET</code></li>
            <li>Redeploy the site</li>
          </ol>
        </div>
      </details>

      <div style="text-align:right;margin-top:16px">
        <button onclick="document.getElementById('rvt-setup-overlay').remove(); showModelStatus('','');" style="padding:8px 20px;background:var(--primary);color:white;border:none;border-radius:6px;cursor:pointer;font-size:13px">Got it</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
}

// ═══════════════════════════════════════════════
//   UPLOADED MODEL — WIND ANALYSIS
// ═══════════════════════════════════════════════

function distPointToSegment2D(px, pz, x1, z1, x2, z2){
  const dx = x2 - x1, dz = z2 - z1;
  const len2 = dx * dx + dz * dz;
  if(len2 < 1e-18) return Math.hypot(px - x1, pz - z1);
  let t = ((px - x1) * dx + (pz - z1) * dz) / len2;
  t = Math.max(0, Math.min(1, t));
  const nx = x1 + t * dx, nz = z1 + t * dz;
  return Math.hypot(px - nx, pz - nz);
}

/** Min distance from (px,pz) to polygon edges (XZ footprint). */
function distPointToPolygonBoundaryXZ(px, pz, poly){
  if(!poly || poly.length < 2) return Infinity;
  let minD = Infinity;
  const n = poly.length;
  for(let i = 0; i < n; i++){
    const j = (i + 1) % n;
    const d = distPointToSegment2D(px, pz, poly[i][0], poly[i][1], poly[j][0], poly[j][1]);
    if(d < minD) minD = d;
  }
  return minD;
}

/** Andrew monotone chain — points are [x,z] in site metres. */
function convexHullXZ(points){
  if(!points || points.length < 3) return points ? points.slice() : [];
  const pts = points.slice().sort((a, b) => (a[0] !== b[0] ? a[0] - b[0] : a[1] - b[1]));
  function cross(o, a, b){
    return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);
  }
  const lower = [];
  for(let i = 0; i < pts.length; i++){
    while(lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], pts[i]) <= 0) lower.pop();
    lower.push(pts[i]);
  }
  const upper = [];
  for(let i = pts.length - 1; i >= 0; i--){
    while(upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], pts[i]) <= 0) upper.pop();
    upper.push(pts[i]);
  }
  upper.pop();
  lower.pop();
  return lower.concat(upper);
}

function collectMeshCentersXZFromGroup(group){
  const pts = [];
  if(!group) return pts;
  group.updateMatrixWorld(true);
  group.traverse(child => {
    if(!child.isMesh || !child.geometry) return;
    const box = new THREE.Box3().setFromObject(child);
    if(box.isEmpty()) return;
    const c = box.getCenter(new THREE.Vector3());
    pts.push([c.x, c.z]);
  });
  return pts;
}

function buildUploadFootprintHullContext(xzPts, bbox){
  const extX = bbox.max.x - bbox.min.x;
  const extZ = bbox.max.z - bbox.min.z;
  const maxXZ = Math.max(extX, extZ, 1e-9);
  if(!xzPts || xzPts.length < 3) return null;
  const hull = convexHullXZ(xzPts);
  if(hull.length < 3) return null;
  return {
    hull,
    maxXZ,
    band: Math.max(0.11, 0.12 * maxXZ)
  };
}

/**
 * Merge labels on coplanar-ish façade shards: majority vote by surface area per bucket.
 */
function applyCoplanarFaceKeyConsensus(group, faceMap, bCenter){
  if(!group || !faceMap || faceMap.size < 2) return;
  // Phase 5: pull Phase 3 shellScore so shell-confident meshes get extra weight.
  const diagMap = (typeof uploadClassDiag !== 'undefined' && uploadClassDiag) ? uploadClassDiag : null;
  const shellWeightFor = uuid => {
    if(!diagMap) return 1.0;
    const d = diagMap.get(uuid);
    const s = d && Number.isFinite(d.shellScore) ? Math.max(0, Math.min(1, d.shellScore)) : 0;
    // shell-confident meshes count ~3.3x more per unit area; interior partitions
    // still contribute but cannot dominate.
    return 0.3 + 0.7 * s;
  };
  const buckets = new Map();
  group.traverse(child => {
    if(!child.isMesh) return;
    const fk = faceMap.get(child.uuid);
    if(!fk) return;
    const n = computeMeshNormalForWind(child, bCenter);
    if(!n || n.lengthSq() < 1e-12) return;
    const meshBox = new THREE.Box3().setFromObject(child);
    if(meshBox.isEmpty()) return;
    const mc = meshBox.getCenter(new THREE.Vector3());
    const d = n.x * mc.x + n.y * mc.y + n.z * mc.z;
    const key = [
      Math.round(n.x * 64) / 64,
      Math.round(n.y * 64) / 64,
      Math.round(n.z * 64) / 64,
      Math.round(d * 4) / 4
    ].join('|');
    const area = computeMeshSurfaceAreaWorld(child);
    if(!Number.isFinite(area) || area <= 0) return;
    const shellW = shellWeightFor(child.uuid);
    const weight = area * shellW;
    if(!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push({ uuid: child.uuid, area, weight, shellW, fk });
  });
  buckets.forEach(arr => {
    if(arr.length < 2) return;
    // If every mesh in the bucket is shell-weak (likely mis-bucketed interior
    // partitions), skip consensus — don't amplify the error across the bucket.
    const maxShell = arr.reduce((m, e) => Math.max(m, e.shellW), 0);
    if(maxShell < 0.4) return;
    const tally = {};
    let total = 0;
    arr.forEach(({ weight, fk }) => {
      tally[fk] = (tally[fk] || 0) + weight;
      total += weight;
    });
    if(total <= 0) return;
    let winner = arr[0].fk, max = 0;
    for(const k in tally){
      if(tally[k] > max){ max = tally[k]; winner = k; }
    }
    // Tightened threshold: 55% of weighted votes to reduce flip-flopping
    // between near-ties. Shell weighting means a single confident outlier with
    // a large face can still override a weakly-classified majority.
    if(max / total >= 0.55){
      arr.forEach(({ uuid }) => { faceMap.set(uuid, winner); });
    }
  });
}

function uploadModelOverrideStorageKey(){
  return uploadedModelName || 'unnamed';
}

function loadUploadFaceOverridesObject(){
  try{
    const raw = sessionStorage.getItem(UPLOAD_FACE_OVERRIDE_STORAGE);
    if(!raw) return {};
    const all = JSON.parse(raw);
    const k = uploadModelOverrideStorageKey();
    return all[k] && typeof all[k] === 'object' ? all[k] : {};
  }catch(e){
    return {};
  }
}

function saveUploadFaceOverride(uuid, faceKey){
  if(!isValidUploadFaceKey(faceKey)) return;
  try{
    const raw = sessionStorage.getItem(UPLOAD_FACE_OVERRIDE_STORAGE);
    const all = raw ? JSON.parse(raw) : {};
    const k = uploadModelOverrideStorageKey();
    if(!all[k]) all[k] = {};
    all[k][uuid] = faceKey;
    sessionStorage.setItem(UPLOAD_FACE_OVERRIDE_STORAGE, JSON.stringify(all));
  }catch(e){}
  // Phase 4: capture the correction as a training sample for the next retrain.
  captureUploadCorrection(uuid, faceKey);
}

/**
 * Phase 4 — active-learning capture.
 *
 * A face-key override is a strong label:
 *   windward | leeward | sidewall1 | sidewall2  → wall  (semantic class 1)
 *   roof_ww  | roof_lw | roof_cw | roof_hip_l | roof_hip_r → roof  (class 2)
 *
 * When the user corrects a face-key we cache the mesh's 22-d feature vector
 * plus the derived semantic label into session storage. Call
 * window.exportUploadTrainingCorrections() from the console to download the
 * accumulated corrections as JSONL ready to append to
 * mesh-training-samples.jsonl.
 */
const UPLOAD_TRAIN_CORRECTIONS_STORAGE = 'uploadTrainCorrections.v1';

function faceKeyToSemanticLabel(fk){
  if(!fk) return null;
  if(fk === 'windward' || fk === 'leeward' || fk === 'sidewall1' || fk === 'sidewall2') return 1;
  if(fk === 'roof_ww' || fk === 'roof_lw' || fk === 'roof_cw' || fk === 'roof_hip_l' || fk === 'roof_hip_r') return 2;
  return null;
}

function captureUploadCorrection(uuid, faceKey){
  try{
    const label = faceKeyToSemanticLabel(faceKey);
    if(label === null) return;
    const diag = (typeof uploadClassDiag !== 'undefined' && uploadClassDiag)
      ? uploadClassDiag.get(uuid) : null;
    if(!diag || !diag.features || diag.features.length !== 22) return;
    const raw = sessionStorage.getItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE);
    const all = raw ? JSON.parse(raw) : {};
    const k = uploadModelOverrideStorageKey();
    if(!all[k]) all[k] = {};
    // Phase 12: capture the model's prior prediction at correction time so
    // retraining (and post-hoc analytics) can distinguish boundary-ambiguity
    // corrections from confidently-wrong corrections. The latter are gold —
    // they're the cases where the model needs the most adjustment. All fields
    // are additive; legacy entries without these fields still parse cleanly.
    const numOrNull = v => (Number.isFinite(v) ? Number(v) : null);
    all[k][uuid] = {
      features: Array.from(diag.features).map(v => Number.isFinite(v) ? Number(v) : 0),
      label,
      faceKey,
      ts: Date.now(),
      // Prior face-key prediction + confidence (Phase 8c/9 signal).
      priorFaceKey: diag.faceKey || null,
      priorFaceConf: numOrNull(diag.faceKeyConfidence),
      // Prior MLP semantic class + class confidence.
      priorMlpClass: diag.mlpClass || null,
      priorMlpP: numOrNull(diag.mlpP),
      // Prior analytic-fallback prediction (used by Phase 9 mismatch trigger).
      priorAnalyticKey: diag.analyticKey || null,
      priorAnalyticConf: numOrNull(diag.analyticConf),
      // IFC short-circuit provenance (Phase 1a) — non-zero when an IFC type
      // hint resolved the class without MLP inference.
      priorViaIfc: !!diag.viaIfc,
      priorIfcConf: numOrNull(diag.ifcConf),
      // Phase 17: capture the model's prior zone assignment (Phase 13/14/16) so
      // retraining can distinguish face-key corrections from zone-boundary
      // corrections on zoned faces (sidewall1/2, roof_ww/cw/hip_l/hip_r). null
      // for un-zoned faces (windward, leeward, roof_lw) and for legacy diags
      // that pre-date Phase 13.
      priorZoneIndex: numOrNull(diag.zoneIndex)
    };
    sessionStorage.setItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE, JSON.stringify(all));
    // Phase 7a: refresh the face-inspect retrain row (counter + button state).
    if(typeof updateRetrainUiRow === 'function') updateRetrainUiRow();
  }catch(e){ /* best-effort capture */ }
}

/**
 * Export accumulated face-key corrections as downloadable JSONL. Each non-empty
 * line parses as {"features":[f0..f21],"label":1|2,"faceKey":"...","uuid":"...","model":"..."}
 * — features and label are the fields train_mesh_classifier.py expects; the
 * extra fields are metadata for provenance.
 */
function exportUploadTrainingCorrections(){
  try{
    const raw = sessionStorage.getItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE);
    if(!raw){ console.log('No upload training corrections captured yet.'); return null; }
    const all = JSON.parse(raw);
    const lines = [];
    Object.keys(all).forEach(model => {
      const entries = all[model] || {};
      Object.keys(entries).forEach(uuid => {
        const e = entries[uuid];
        if(!e || !Array.isArray(e.features) || e.features.length !== 22) return;
        // Phase 12: emit prior model state alongside the corrected label.
        // train_mesh_classifier.py only requires `features` + `label`; the rest
        // is provenance + active-learning metadata. Older legacy entries that
        // pre-date Phase 12 won't have prior* fields — those default to null
        // here rather than failing the export.
        const row = {
          features: e.features,
          label: e.label,
          faceKey: e.faceKey || null,
          uuid,
          model,
          priorFaceKey: e.priorFaceKey != null ? e.priorFaceKey : null,
          priorFaceConf: Number.isFinite(e.priorFaceConf) ? e.priorFaceConf : null,
          priorMlpClass: e.priorMlpClass != null ? e.priorMlpClass : null,
          priorMlpP: Number.isFinite(e.priorMlpP) ? e.priorMlpP : null,
          priorAnalyticKey: e.priorAnalyticKey != null ? e.priorAnalyticKey : null,
          priorAnalyticConf: Number.isFinite(e.priorAnalyticConf) ? e.priorAnalyticConf : null,
          priorViaIfc: e.priorViaIfc === true,
          priorIfcConf: Number.isFinite(e.priorIfcConf) ? e.priorIfcConf : null,
          // Phase 17: prior zone assignment — null on legacy entries.
          priorZoneIndex: Number.isFinite(e.priorZoneIndex) ? e.priorZoneIndex : null
        };
        lines.push(JSON.stringify(row));
      });
    });
    if(!lines.length){ console.log('No valid correction rows to export.'); return null; }
    const body = lines.join('\n') + '\n';
    try{
      const blob = new Blob([body], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'mesh-training-corrections.jsonl';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 1500);
    }catch(e){ /* non-browser context — return text */ }
    console.log('Exported ' + lines.length + ' correction rows.');
    return body;
  }catch(e){ console.warn('exportUploadTrainingCorrections failed:', e); return null; }
}

function clearUploadTrainingCorrections(){
  try{ sessionStorage.removeItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE); }catch(e){}
}

/**
 * Phase 6 — pull captured face-key corrections out of sessionStorage, forward
 * them to MeshClassifier.retrainWithCorrections (which upsamples + trains a
 * fresh MLP), then reclassify the currently loaded uploaded model against the
 * new weights. Intended to be invoked from the DevTools console after the user
 * has corrected a handful of misclassified faces.
 *
 * @param {object} [opts] — forwarded to retrainWithCorrections (upsample, noiseSigma).
 * @returns {Promise<{ok:boolean, samples:number, reclassified:boolean}>}
 */
/**
 * Phase 19 — pure predicate deciding whether a captured correction should
 * survive into the retrain row set. Drops:
 *  - structurally invalid rows (missing/wrong-length features, bad label)
 *  - high-confidence reaffirmations (priorFaceKey === faceKey AND
 *    priorFaceConf >= 0.7) — those teach the MLP nothing new because the
 *    model was already correct and confident; including them dilutes the
 *    true-correction signal during retraining.
 *
 * Pass `opts.keepReaffirmations: true` to disable the reaffirmation filter
 * (useful for diagnostics / training-curve experiments).
 *
 * Threshold of 0.7 matches the HIGH band in summarizeUploadCorrections (Phase 18).
 */
function shouldKeepCorrectionForRetrain(e, opts){
  if(!e || !Array.isArray(e.features) || e.features.length !== 22) return false;
  if(e.label !== 0 && e.label !== 1 && e.label !== 2 && e.label !== 3) return false;
  const keepReaffirmations = !!(opts && opts.keepReaffirmations);
  if(!keepReaffirmations){
    const reaffirmed = e.priorFaceKey != null && e.priorFaceKey === e.faceKey;
    const highConf = Number.isFinite(e.priorFaceConf) && e.priorFaceConf >= 0.7;
    if(reaffirmed && highConf) return false;
  }
  return true;
}

async function retrainUploadWithCorrections(opts){
  try{
    if(typeof MeshClassifier === 'undefined' || !MeshClassifier || typeof MeshClassifier.retrainWithCorrections !== 'function'){
      console.warn('retrainUploadWithCorrections: MeshClassifier.retrainWithCorrections unavailable');
      return { ok:false, samples:0, reclassified:false };
    }
    const raw = sessionStorage.getItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE);
    if(!raw){ console.log('retrainUploadWithCorrections: no captured corrections in sessionStorage'); return { ok:false, samples:0, reclassified:false }; }
    const all = JSON.parse(raw);
    const rows = [];
    let droppedReaffirmations = 0;
    Object.keys(all || {}).forEach(model => {
      const entries = all[model] || {};
      Object.keys(entries).forEach(uuid => {
        const e = entries[uuid];
        if(!shouldKeepCorrectionForRetrain(e, opts)){
          // Phase 19: distinguish drops due to reaffirmation-filter vs structural
          // invalidity so the log line is meaningful.
          if(e && Array.isArray(e.features) && e.features.length === 22 &&
             (e.label === 0 || e.label === 1 || e.label === 2 || e.label === 3) &&
             e.priorFaceKey != null && e.priorFaceKey === e.faceKey &&
             Number.isFinite(e.priorFaceConf) && e.priorFaceConf >= 0.7){
            droppedReaffirmations++;
          }
          return;
        }
        rows.push({ features: e.features, label: e.label | 0 });
      });
    });
    if(!rows.length){ console.log('retrainUploadWithCorrections: no valid correction rows'); return { ok:false, samples:0, reclassified:false }; }
    // Phase 18: log the corpus shape so the developer sees what kind of
    // corrections they're retraining on (face-key changes vs reaffirmations,
    // zoned vs un-zoned, prior-confidence bands).
    try{
      const summary = summarizeUploadCorrections(all);
      console.log('retrainUploadWithCorrections: corpus summary', summary,
                  '(dropped ' + droppedReaffirmations + ' high-conf reaffirmation(s) from retrain set)');
    }catch(_){ /* analytics is best-effort */ }
    const res = MeshClassifier.retrainWithCorrections(rows, opts || {});
    if(!res || !res.ok){ console.warn('retrainUploadWithCorrections: retrain failed', res); return { ok:false, samples:rows.length, reclassified:false }; }
    let reclassified = false;
    try{
      if(typeof classifyUploadedFaces === 'function'){ await classifyUploadedFaces(); reclassified = true; }
    }catch(e){ console.warn('retrainUploadWithCorrections: classifyUploadedFaces threw', e); }
    console.log('retrainUploadWithCorrections: retrained on ' + rows.length + ' correction(s); live reclassified=' + reclassified);
    return { ok:true, samples:rows.length, reclassified, droppedReaffirmations };
  }catch(e){
    console.warn('retrainUploadWithCorrections failed:', e);
    return { ok:false, samples:0, reclassified:false };
  }
}

if(typeof window !== 'undefined'){
  window.exportUploadTrainingCorrections = exportUploadTrainingCorrections;
  window.clearUploadTrainingCorrections = clearUploadTrainingCorrections;
  window.retrainUploadWithCorrections = retrainUploadWithCorrections;
}

/**
 * Phase 18 — walk captured corrections and produce a breakdown of the active-
 * learning corpus shape: face-key changes vs reaffirmations, prior-confidence
 * banding, per-face-key counts, and priorZoneIndex distribution (the Phase 17
 * signal). Pure function over the parsed sessionStorage object — no DOM or
 * window access — so the same logic is reusable in tests / Node consoles.
 *
 * Confidence bands match the Phase 9 gate threshold (IFC_AI_REFINE_FACE_CONF_MIN
 * = 0.15) and the Phase 18 high-conf threshold (0.7) used to flag rows where
 * the user merely reaffirmed an already-confident prediction.
 *
 * @param {object|null} all — { [model]: { [uuid]: capturedRecord } }
 * @returns {{
 *   total: number,
 *   byChangeKind: { faceKeyChanged: number, reaffirmation: number, unknownPrior: number },
 *   byPriorConfBand: { high: number, mid: number, low: number, missing: number },
 *   byPriorFaceKey: object,
 *   byPriorZoneIndex: object,
 *   zonedCorrections: number,
 *   unzonedCorrections: number
 * }}
 */
function summarizeUploadCorrections(all){
  const out = {
    total: 0,
    byChangeKind: { faceKeyChanged: 0, reaffirmation: 0, unknownPrior: 0 },
    byPriorConfBand: { high: 0, mid: 0, low: 0, missing: 0 },
    byPriorFaceKey: Object.create(null),
    byPriorZoneIndex: Object.create(null),
    zonedCorrections: 0,
    unzonedCorrections: 0
  };
  if(!all || typeof all !== 'object') return out;
  const HIGH = 0.7;
  const LOW  = 0.15; // matches IFC_AI_REFINE_FACE_CONF_MIN
  Object.keys(all).forEach(model => {
    const entries = all[model];
    if(!entries || typeof entries !== 'object') return;
    Object.keys(entries).forEach(uuid => {
      const e = entries[uuid];
      if(!e || !Array.isArray(e.features) || e.features.length !== 22) return;
      out.total++;
      // Change-kind: did the user override the face-key?
      if(e.priorFaceKey == null){
        out.byChangeKind.unknownPrior++;
      } else if(e.priorFaceKey === e.faceKey){
        out.byChangeKind.reaffirmation++;
      } else {
        out.byChangeKind.faceKeyChanged++;
      }
      // Prior-confidence band.
      if(!Number.isFinite(e.priorFaceConf)){
        out.byPriorConfBand.missing++;
      } else if(e.priorFaceConf >= HIGH){
        out.byPriorConfBand.high++;
      } else if(e.priorFaceConf >= LOW){
        out.byPriorConfBand.mid++;
      } else {
        out.byPriorConfBand.low++;
      }
      // Per-face-key + per-zone counters. Use string keys so JSON survives.
      const fk = e.priorFaceKey || 'unknown';
      out.byPriorFaceKey[fk] = (out.byPriorFaceKey[fk] || 0) + 1;
      const zi = Number.isFinite(e.priorZoneIndex) ? String(e.priorZoneIndex) : 'null';
      out.byPriorZoneIndex[zi] = (out.byPriorZoneIndex[zi] || 0) + 1;
      if(Number.isFinite(e.priorZoneIndex)) out.zonedCorrections++;
      else out.unzonedCorrections++;
    });
  });
  return out;
}

if(typeof window !== 'undefined'){
  window.summarizeUploadCorrections = function(){
    try{
      const raw = sessionStorage.getItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE);
      const all = raw ? JSON.parse(raw) : null;
      const summary = summarizeUploadCorrections(all);
      console.log('Upload corrections corpus summary:', summary);
      return summary;
    }catch(e){ console.warn('summarizeUploadCorrections failed:', e); return null; }
  };
}

/**
 * Phase 7a — count captured correction rows (across all uploaded models) so the
 * face-inspect UI can show a live summary. Returns 0 on error / empty storage.
 */
function countUploadTrainingCorrections(){
  try{
    const raw = sessionStorage.getItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE);
    if(!raw) return 0;
    const all = JSON.parse(raw);
    let n = 0;
    Object.keys(all || {}).forEach(model => {
      const entries = all[model] || {};
      Object.keys(entries).forEach(uuid => {
        const e = entries[uuid];
        if(e && Array.isArray(e.features) && e.features.length === 22) n += 1;
      });
    });
    return n;
  }catch(e){ return 0; }
}

/**
 * Phase 7a — refresh the retrain-row UI: show it when the face-inspect panel
 * is visible on an uploaded model, update the summary label, and enable/disable
 * the Retrain button based on whether any corrections are captured.
 */
function updateRetrainUiRow(){
  try{
    const row = document.getElementById('fi-retrain-row');
    if(!row) return;
    // Only show when the override row is visible (i.e. uploaded-model face selected).
    const overrideRow = document.getElementById('fi-override-row');
    const overrideVisible = overrideRow && overrideRow.style.display !== 'none';
    if(!overrideVisible){ row.style.display = 'none'; return; }
    const n = countUploadTrainingCorrections();
    const hasPersisted = !!(typeof MeshClassifier !== 'undefined' && MeshClassifier && typeof MeshClassifier.hasPersistedWeights === 'function' && MeshClassifier.hasPersistedWeights());
    const summary = document.getElementById('fi-retrain-summary');
    if(summary){
      summary.textContent = n + ' correction' + (n === 1 ? '' : 's') + (hasPersisted ? ' · persisted' : '');
    }
    const btn = document.getElementById('fi-retrain-apply');
    if(btn){ btn.disabled = (n === 0); btn.style.opacity = (n === 0 ? 0.55 : 1); }
    const reset = document.getElementById('fi-retrain-reset');
    if(reset){ reset.disabled = !hasPersisted; reset.style.opacity = (hasPersisted ? 1 : 0.55); }
    row.style.display = 'flex';
  }catch(e){ /* best-effort */ }
}

/**
 * Phase 7a — wired to the "🧠 Retrain" button in the face-inspect panel. Kicks
 * off retrainUploadWithCorrections and surfaces the result inline in the
 * summary label (no modal).
 */
async function retrainUploadFromUi(){
  const btn = document.getElementById('fi-retrain-apply');
  const summary = document.getElementById('fi-retrain-summary');
  try{
    if(btn){ btn.disabled = true; btn.textContent = '⏳ Retraining…'; }
    const res = await retrainUploadWithCorrections();
    if(summary){
      if(res && res.ok){
        summary.textContent = 'Retrained on ' + res.samples + ' · persisted';
      }else{
        summary.textContent = 'Retrain failed';
      }
    }
  }catch(e){
    if(summary) summary.textContent = 'Retrain error';
    console.warn('retrainUploadFromUi failed:', e);
  }finally{
    if(btn){ btn.textContent = '🧠 Retrain'; btn.disabled = false; }
    // Let the label linger briefly, then refresh counters.
    setTimeout(updateRetrainUiRow, 1800);
  }
}

/**
 * Phase 7a — wired to the "Reset" button. Clears the persisted retrained
 * weights from localStorage. Note: the current in-memory MLP stays until the
 * next page reload (or we could call initMlp() to immediately revert).
 */
function resetMeshClassifierWeightsFromUi(){
  try{
    if(typeof MeshClassifier === 'undefined' || !MeshClassifier || typeof MeshClassifier.clearPersistedWeights !== 'function') return;
    const ok = MeshClassifier.clearPersistedWeights();
    const summary = document.getElementById('fi-retrain-summary');
    if(summary) summary.textContent = ok ? 'Persisted weights cleared · reload to apply' : 'Clear failed';
    setTimeout(updateRetrainUiRow, 2200);
  }catch(e){ console.warn('resetMeshClassifierWeightsFromUi failed:', e); }
}

if(typeof window !== 'undefined'){
  window.retrainUploadFromUi = retrainUploadFromUi;
  window.resetMeshClassifierWeightsFromUi = resetMeshClassifierWeightsFromUi;
  window.updateRetrainUiRow = updateRetrainUiRow;
  window.countUploadTrainingCorrections = countUploadTrainingCorrections;
}

function clearUploadFaceOverride(uuid){
  try{
    const raw = sessionStorage.getItem(UPLOAD_FACE_OVERRIDE_STORAGE);
    if(!raw) return;
    const all = JSON.parse(raw);
    const k = uploadModelOverrideStorageKey();
    if(all[k] && all[k][uuid]) delete all[k][uuid];
    sessionStorage.setItem(UPLOAD_FACE_OVERRIDE_STORAGE, JSON.stringify(all));
  }catch(e){}
  // Phase 4: also retract the training correction row so it isn't exported.
  try{
    const raw2 = sessionStorage.getItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE);
    if(!raw2) return;
    const allC = JSON.parse(raw2);
    const k2 = uploadModelOverrideStorageKey();
    if(allC[k2] && allC[k2][uuid]){
      delete allC[k2][uuid];
      sessionStorage.setItem(UPLOAD_TRAIN_CORRECTIONS_STORAGE, JSON.stringify(allC));
    }
  }catch(e){}
  // Phase 7a: refresh the face-inspect retrain row (counter + button state).
  if(typeof updateRetrainUiRow === 'function') updateRetrainUiRow();
}

function applyUploadFaceOverrideFromUi(){
  if(!lastUploadHoverUuid || !uploadedModelGroup) return;
  const sel = document.getElementById('fi-override-select');
  if(!sel) return;
  const v = sel.value;
  if(!isValidUploadFaceKey(v)) return;
  saveUploadFaceOverride(lastUploadHoverUuid, v);
  uploadFaceMap.set(lastUploadHoverUuid, v);
  const d = uploadClassDiag.get(lastUploadHoverUuid);
  if(d) d.faceKey = v;
  recolorUploadedModel();
  patchUploadedModelResultAreas(aggregateUploadedMeshAreasByFaceKey());
  toast('Zone override saved (this session).');
}

function clearUploadFaceOverrideFromUi(){
  if(!lastUploadHoverUuid) return;
  clearUploadFaceOverride(lastUploadHoverUuid);
  void applyWindToUploadedModel();
  toast('Override cleared — mesh reclassified.');
}

function applyUploadFaceOverridesFromStorage(faceMap){
  const o = loadUploadFaceOverridesObject();
  for(const uuid in o){
    if(!isValidUploadFaceKey(o[uuid])) continue;
    if(faceMap.has(uuid)) faceMap.set(uuid, o[uuid]);
  }
}

/**
 * Phase 13 — assign each IFC mesh a zone index per AS/NZS 1170.2 distance-from-
 * windward-edge bands. Mirrors the parametric zone scheme in
 * `sidewallCpZones()` (4 bands at 0..1h, 1h..2h, 2h..3h, >3h) and
 * `roofUpwindCpZones()` (5 bands at 0..0.5h, 0.5h..1h, 1h..2h, 2h..3h, >3h).
 *
 * Returns the index into `zones[]` whose [cumulative_start, cumulative_end]
 * window contains the centroid's projected distance from the windward edge.
 * Returns null when zoning doesn't apply (windward/leeward/roof_lw/roof_cw)
 * or when the zones array is missing.
 *
 * @param {string} faceKey         — assigned face-key
 * @param {{x,y,z}} mc             — mesh centroid (world)
 * @param {{x,y,z}} bCenter        — building bbox centre
 * @param {{x,y,z}} bSize          — building bbox size (extents along world axes)
 * @param {{x,y,z}} windDir        — unit vector pointing from windward to leeward
 * @param {Array<{width:number}>} zones — face's parametric zones array
 * @returns {number|null}
 */
function assignWindEdgeZone(faceKey, mc, bCenter, bSize, windDir, zones){
  if(!zones || !zones.length) return null;
  // Edge-distance zoning applies to surfaces whose Cp,e varies with distance
  // along the wind axis. Per AS/NZS 1170.2 the reference plane differs:
  //   windward edge (Phase 13/14): side walls (Table 5.2(C)), roof_ww and
  //     roof_cw (Table 5.3(A)), and hip-end UPWIND face roof_hip_r (whose
  //     bands run from u=0 at the windward eave to u=horiz at the ridge).
  //   ridge plane (Phase 16): hip-end DOWNWIND face roof_hip_l, whose bands
  //     run from u=uRidge to u=uLee per hipEndDownwindCpZonesTableA. The
  //     returned zone widths are the in-band u-extents on the leeward half,
  //     so we walk them from the ridge plane (signed projW = 0 at bbox
  //     centre, which coincides with the ridge for a symmetric hip roof).
  // windward/leeward walls and roof_lw (gable) are uniform → not in ZONED.
  const ZONED = {
    sidewall1: 'windward_edge', sidewall2: 'windward_edge',
    roof_ww: 'windward_edge', roof_cw: 'windward_edge',
    roof_hip_r: 'windward_edge',
    roof_hip_l: 'ridge'
  };
  const ref = ZONED[faceKey];
  if(!ref) return null;
  // Project (mc - bCenter) onto windDir → signed distance along wind from
  // the bbox centre. windDir points windward→leeward by convention (a wall
  // with normal·windDir < 0 is windward; see mesh-classifier.js Phase 8a).
  const dx = mc.x - bCenter.x;
  const dy = mc.y - bCenter.y;
  const dz = mc.z - bCenter.z;
  const projW = dx * windDir.x + dy * windDir.y + dz * windDir.z;
  if(!Number.isFinite(projW)) return null;
  let dist;
  if(ref === 'ridge'){
    // Ridge-relative: dist measured from bbox centre toward the leeward eave.
    // A hip_l mesh whose centroid lands on the windward half of the bbox is
    // a geometry mismatch (face-key was misassigned upstream); skip zoning
    // and let the caller fall back to the face-uniform Cp,e.
    if(projW < 0) return null;
    dist = projW;
  } else {
    // Half-extent of axis-aligned bbox along windDir: the bbox-corner
    // projection collapses to |sx*wx| + |sy*wy| + |sz*wz| (sum over axes).
    const halfExt = 0.5 * (Math.abs(bSize.x * windDir.x)
                          + Math.abs(bSize.y * windDir.y)
                          + Math.abs(bSize.z * windDir.z));
    // Distance from the windward edge: 0 at the windward face of the bbox,
    // 2*halfExt = effD at the leeward face.
    dist = projW + halfExt;
    if(!Number.isFinite(dist)) return null;
    if(dist <= 0) return 0; // upwind overhang — zone 0
  }
  // Walk cumulative widths: zone i covers [cumStart, cumStart + zones[i].width).
  let cumStart = 0;
  for(let i = 0; i < zones.length; i++){
    const w = zones[i] && Number.isFinite(zones[i].width) ? zones[i].width : 0;
    const cumEnd = cumStart + w;
    if(dist < cumEnd) return i;
    cumStart = cumEnd;
  }
  // Beyond all zones (mesh past the leeward extent — overhang or numerical
  // slop): clamp to the last zone (the >3h or ridge-relative tail band).
  return zones.length - 1;
}

function meshNeedsIfcAiRefinement(bestP, faceKey, analyticKey, faceKeyConfidence){
  if(!Number.isFinite(bestP)) return true;
  if(bestP <= IFC_AI_REFINE_CONF_MAX) return true;
  if(!faceKey) return true;
  if(analyticKey && faceKey !== analyticKey) return true;
  // Phase 9: low face-key confidence (e.g. wall normal ~45° to wind, or roof
  // pitch within hair of flatPitchDeg threshold) is itself a refinement trigger
  // even when MLP class confidence is high — the *face* assignment is the
  // ambiguous bit, not the wall/roof decision.
  if(Number.isFinite(faceKeyConfidence) && faceKeyConfidence < IFC_AI_REFINE_FACE_CONF_MIN) return true;
  return false;
}

/**
 * Skip meshes that are clearly inside the volume (floors, ceilings, partitions) so the MLP
 * cannot mislabel them as exterior wall/roof. Plan "shell" alone misses inset/recessed façades
 * (courtyards, deep porches); those are kept when mean outward normal places the mesh outside
 * the plan centroid (same test used for analytic fallback).
 * @param {object|null} hullCtx — optional { hull, band } from convex hull of mesh centroids (XZ)
 */
function geoSkipNonEnvelopeMesh(mc, bbox, bCenter, nAcc, meshBox, hullCtx){
  const bmin = bbox.min, bmax = bbox.max;
  const extX = bmax.x - bmin.x, extY = bmax.y - bmin.y, extZ = bmax.z - bmin.z;
  const ux = extX > 1e-6 ? (mc.x - bmin.x) / extX : 0.5;
  const uy = extY > 1e-6 ? (mc.y - bmin.y) / extY : 0.5;
  const uz = extZ > 1e-6 ? (mc.z - bmin.z) / extZ : 0.5;
  const ay = Math.abs(nAcc.y);
  const horiz = Math.hypot(nAcc.x, nAcc.z);
  // Horizontal slabs (ground floor, intermediate floors, ceilings): not roof membrane.
  // Old rule required uy>0.05, so ground plates at uy≈0 were misclassified as walls.
  // Use mesh top height: only the upper band can be a flat roof deck / horizontal roof.
  if(ay > 0.72 && horiz < 0.68){
    const meshTopY = meshBox.max.y;
    const roofBandFloor = bmax.y - Math.max(extY * 0.09, 0.12);
    if(meshTopY < roofBandFloor) return true;
  }
  // IFC floor / footing plates: thin vs footprint, on/near ground — not curtain wall (use min span so wide façades aren't skipped)
  const msz = meshBox.getSize(new THREE.Vector3());
  const nearGround = (meshBox.min.y - bmin.y) < extY * 0.07;
  const minSpan = Math.max(Math.min(msz.x, msz.z), 1e-6);
  const slabLike = msz.y < extY * 0.18 && msz.y < minSpan * 0.22;
  if(nearGround && slabLike) return true;
  // Vertical walls/partitions: skip only when clearly interior (not on bbox shell AND not outside in plan vs centroid)
  const SHELL = 0.13;
  const inPlanShell = ux < SHELL || ux > 1 - SHELL || uz < SHELL || uz > 1 - SHELL;
  if(ay < 0.52 && uy < 0.74 && !inPlanShell){
    if(hullCtx && hullCtx.hull && hullCtx.hull.length >= 3){
      const dEdge = distPointToPolygonBoundaryXZ(mc.x, mc.z, hullCtx.hull);
      if(dEdge < hullCtx.band) return false;
    }
    if(typeof MeshClassifier === 'undefined' || !MeshClassifier.outwardAlignedNormal) return true;
    const n = MeshClassifier.outwardAlignedNormal(nAcc, mc, bCenter);
    const nHx = n.x, nHz = n.z;
    const hLen = Math.hypot(nHx, nHz);
    if(hLen < 1e-5) return true;
    const nH = new THREE.Vector3(nHx / hLen, 0, nHz / hLen);
    const vH = new THREE.Vector3(mc.x - bCenter.x, 0, mc.z - bCenter.z);
    const outDot = vH.dot(nH);
    const halfH = Math.max(extX, extZ) * 0.5;
    if(outDot < halfH * 0.08) return true;
  }
  return false;
}

/**
 * Triangle area–weighted mean normal in world space (stable for glazed façades: small mullions
 * and frames are not drowned by mixed vertex normals from openings).
 */
function computeMeshAreaWeightedNormal(mesh){
  const g = mesh.geometry;
  const pos = g.attributes.position;
  if(!pos) return null;
  mesh.updateMatrixWorld(true);
  const m = mesh.matrixWorld;
  const idx = g.index;
  const nSum = new THREE.Vector3();
  const vA = new THREE.Vector3(), vB = new THREE.Vector3(), vC = new THREE.Vector3();
  const ab = new THREE.Vector3(), ac = new THREE.Vector3(), cr = new THREE.Vector3();

  function addTri(i0, i1, i2){
    vA.fromBufferAttribute(pos, i0).applyMatrix4(m);
    vB.fromBufferAttribute(pos, i1).applyMatrix4(m);
    vC.fromBufferAttribute(pos, i2).applyMatrix4(m);
    ab.subVectors(vB, vA);
    ac.subVectors(vC, vA);
    cr.crossVectors(ab, ac);
    const area = cr.length() * 0.5;
    if(area < 1e-18) return;
    cr.normalize();
    nSum.addScaledVector(cr, area);
  }

  if(idx){
    const ia = idx.array;
    for(let i = 0; i < ia.length; i += 3) addTri(ia[i], ia[i + 1], ia[i + 2]);
  } else {
    const nc = pos.count;
    for(let i = 0; i < nc; i += 3) addTri(i, i + 1, i + 2);
  }
  if(nSum.lengthSq() < 1e-20) return null;
  return nSum.normalize();
}

/**
 * Area-weighted normals for triangles that face *outward* from the building centroid (dot(n, triCenter - bCenter) > 0).
 * Fixes louvers / slatted cladding where opposite-facing triangles cancel in a plain area sum.
 */
function computeOutwardAreaWeightedNormal(mesh, bCenter){
  const g = mesh.geometry;
  const pos = g.attributes.position;
  if(!pos) return null;
  mesh.updateMatrixWorld(true);
  const m = mesh.matrixWorld;
  const idx = g.index;
  const nSum = new THREE.Vector3();
  const vA = new THREE.Vector3(), vB = new THREE.Vector3(), vC = new THREE.Vector3();
  const ab = new THREE.Vector3(), ac = new THREE.Vector3(), cr = new THREE.Vector3();
  const tc = new THREE.Vector3();
  const toC = new THREE.Vector3();

  function addTri(i0, i1, i2){
    vA.fromBufferAttribute(pos, i0).applyMatrix4(m);
    vB.fromBufferAttribute(pos, i1).applyMatrix4(m);
    vC.fromBufferAttribute(pos, i2).applyMatrix4(m);
    ab.subVectors(vB, vA);
    ac.subVectors(vC, vA);
    cr.crossVectors(ab, ac);
    const area = cr.length() * 0.5;
    if(area < 1e-18) return;
    cr.normalize();
    tc.addVectors(vA, vB).add(vC).multiplyScalar(1 / 3);
    toC.subVectors(tc, bCenter);
    if(toC.dot(cr) <= 1e-6) return;
    nSum.addScaledVector(cr, area);
  }

  if(idx){
    const ia = idx.array;
    for(let i = 0; i < ia.length; i += 3) addTri(ia[i], ia[i + 1], ia[i + 2]);
  } else {
    const nc = pos.count;
    for(let i = 0; i < nc; i += 3) addTri(i, i + 1, i + 2);
  }
  if(nSum.lengthSq() < 1e-20) return null;
  return nSum.normalize();
}

function computeMeshNormalForWind(mesh, bCenter){
  const o = computeOutwardAreaWeightedNormal(mesh, bCenter);
  if(o) return o;
  const a = computeMeshAreaWeightedNormal(mesh);
  if(a) return a;
  return null;
}

/** Total triangle surface area (m²) in world space — for uploaded IFC/BIM meshes, not parametric sliders. */
function computeMeshSurfaceAreaWorld(mesh){
  const g = mesh.geometry;
  const pos = g && g.attributes.position;
  if(!pos) return 0;
  mesh.updateMatrixWorld(true);
  const m = mesh.matrixWorld;
  const idx = g.index;
  let sum = 0;
  const vA = new THREE.Vector3(), vB = new THREE.Vector3(), vC = new THREE.Vector3();
  const ab = new THREE.Vector3(), ac = new THREE.Vector3(), cr = new THREE.Vector3();
  function addTri(i0, i1, i2){
    vA.fromBufferAttribute(pos, i0).applyMatrix4(m);
    vB.fromBufferAttribute(pos, i1).applyMatrix4(m);
    vC.fromBufferAttribute(pos, i2).applyMatrix4(m);
    ab.subVectors(vB, vA);
    ac.subVectors(vC, vA);
    cr.crossVectors(ab, ac);
    sum += cr.length() * 0.5;
  }
  if(idx){
    const ia = idx.array;
    for(let i = 0; i < idx.count; i += 3) addTri(ia[i], ia[i + 1], ia[i + 2]);
  } else {
    for(let i = 0; i < pos.count; i += 3) addTri(i, i + 1, i + 2);
  }
  return sum;
}

/** Sum classified mesh areas per wind face key (uploaded model geometry). */
function aggregateUploadedMeshAreasByFaceKey(){
  const totals = {};
  if(!uploadedModelGroup || !uploadFaceMap.size) return totals;
  uploadedModelGroup.traverse(child => {
    if(!child.isMesh) return;
    const fk = uploadFaceMap.get(child.uuid);
    if(!fk) return;
    const a = computeMeshSurfaceAreaWorld(child);
    if(!Number.isFinite(a) || a <= 0) return;
    totals[fk] = (totals[fk] || 0) + a;
  });
  return totals;
}

/**
 * Phase 21 — bin each TRIANGLE of an uploaded mesh into a zone band by its
 * world-space centroid. Replaces Phase 20's mesh-centroid bucketing for the
 * zoned face-keys: a single mesh that physically straddles 0-1h and 1h-2h
 * now contributes to both bands rather than dumping all area into one. Pure
 * helper — geometry-only, no side effects on faceMap or diagnostics.
 *
 * Per-triangle assignment reuses assignWindEdgeZone (Phase 13/14/16) on the
 * triangle centroid. A triangle whose centroid yields zoneIndex=null (e.g.
 * upwind of windward eave for a hip_l face) buckets under 'unzoned'.
 *
 * @param {THREE.Mesh} mesh
 * @param {string} faceKey
 * @param {THREE.Vector3} bCenter — building bbox centre (world)
 * @param {THREE.Vector3} bSize   — building bbox extents (world)
 * @param {THREE.Vector3} windDir — unit windward→leeward
 * @param {Array<{width:number}>} zones — face's parametric zones
 * @returns {object} { [zIdx|'unzoned']: areaSum }
 */
function binTriangleAreasByZoneForMesh(mesh, faceKey, bCenter, bSize, windDir, zones){
  const out = {};
  const g = mesh && mesh.geometry;
  const pos = g && g.attributes && g.attributes.position;
  if(!pos || !zones || !zones.length) return out;
  mesh.updateMatrixWorld(true);
  const m = mesh.matrixWorld;
  const idx = g.index;
  const vA = new THREE.Vector3(), vB = new THREE.Vector3(), vC = new THREE.Vector3();
  const ab = new THREE.Vector3(), ac = new THREE.Vector3(), cr = new THREE.Vector3();
  const tc = new THREE.Vector3();
  function addTri(i0, i1, i2){
    vA.fromBufferAttribute(pos, i0).applyMatrix4(m);
    vB.fromBufferAttribute(pos, i1).applyMatrix4(m);
    vC.fromBufferAttribute(pos, i2).applyMatrix4(m);
    ab.subVectors(vB, vA);
    ac.subVectors(vC, vA);
    cr.crossVectors(ab, ac);
    const a = cr.length() * 0.5;
    if(!Number.isFinite(a) || a <= 0) return;
    tc.set(
      (vA.x + vB.x + vC.x) / 3,
      (vA.y + vB.y + vC.y) / 3,
      (vA.z + vB.z + vC.z) / 3
    );
    const zi = assignWindEdgeZone(faceKey, tc, bCenter, bSize, windDir, zones);
    const k = (zi == null) ? 'unzoned' : String(zi);
    out[k] = (out[k] || 0) + a;
  }
  if(idx){
    const ia = idx.array;
    for(let i = 0; i < idx.count; i += 3) addTri(ia[i], ia[i + 1], ia[i + 2]);
  } else {
    for(let i = 0; i < pos.count; i += 3) addTri(i, i + 1, i + 2);
  }
  return out;
}

/**
 * Recompute building bbox + wind direction from the current scene state. Used
 * by Phase 21's per-triangle binner so the area aggregator doesn't have to
 * thread these through from classifyUploadedFaces. Returns null if the
 * uploaded model isn't loaded.
 */
function getUploadModelWindContext(){
  if(!uploadedModelGroup) return null;
  uploadedModelGroup.updateMatrixWorld(true);
  const bbox = new THREE.Box3().setFromObject(uploadedModelGroup);
  if(!Number.isFinite(bbox.min.x) || !Number.isFinite(bbox.max.x)) return null;
  const bCenter = bbox.getCenter(new THREE.Vector3());
  const bSize = bbox.getSize(new THREE.Vector3());
  const relW = ((S.windAngle - S.mapBuildingAngle) % 360 + 360) % 360;
  const ang = (relW + (S.R && S.R.angleOff ? S.R.angleOff : 0)) * Math.PI / 180;
  const windDirLocal = new THREE.Vector3(-Math.sin(ang), 0, -Math.cos(ang));
  const windDir = windDirLocal.clone().applyAxisAngle(new THREE.Vector3(0, 1, 0), -S.mapBuildingAngle * Math.PI / 180);
  windDir.normalize();
  return { bCenter, bSize, windDir };
}

/**
 * Phase 20/21 — sum classified mesh areas per (faceKey, zoneIndex) for the
 * uploaded model. For zoned face-keys (sidewall1/2, roof_ww, roof_cw, hip
 * ends) Phase 21 walks individual triangles via binTriangleAreasByZoneForMesh
 * so a mesh that straddles two bands contributes to both. For un-zoned face-
 * keys (windward, leeward, roof_lw) we still take the whole mesh area under
 * key 'all'. Phase 20's mesh-centroid path is the documented fallback when
 * the wind context isn't available (e.g. tests outside the THREE scene).
 *
 * Returns: { [faceKey]: { [zIdx|'unzoned'|'all']: areaSum } }
 */
function aggregateUploadedMeshAreasByZone(){
  const out = {};
  if(!uploadedModelGroup || !uploadFaceMap.size) return out;
  const ZONED_FK = /^(sidewall1|sidewall2|roof_ww|roof_cw|roof_hip_l|roof_hip_r)$/;
  // Phase 21: reuse a single wind context for all meshes (bbox + windDir
  // don't change within a single classification pass). null context falls
  // back to Phase 20's per-mesh centroid bucketing.
  let ctx = null;
  try{ ctx = getUploadModelWindContext(); }catch(_){ ctx = null; }
  uploadedModelGroup.traverse(child => {
    if(!child.isMesh) return;
    const fk = uploadFaceMap.get(child.uuid);
    if(!fk) return;
    if(!out[fk]) out[fk] = {};
    if(!ZONED_FK.test(fk)){
      const a = computeMeshSurfaceAreaWorld(child);
      if(!Number.isFinite(a) || a <= 0) return;
      out[fk].all = (out[fk].all || 0) + a;
      return;
    }
    // Zoned face-key.
    const zonesForFk = (S.R && S.R.faces && S.R.faces[fk] && Array.isArray(S.R.faces[fk].zones))
      ? S.R.faces[fk].zones : null;
    if(ctx && zonesForFk){
      // Phase 21: per-triangle bucketing.
      const bins = binTriangleAreasByZoneForMesh(child, fk, ctx.bCenter, ctx.bSize, ctx.windDir, zonesForFk);
      Object.keys(bins).forEach(k => {
        out[fk][k] = (out[fk][k] || 0) + bins[k];
      });
      return;
    }
    // Phase 20 fallback: whole-mesh area at the per-mesh centroid band.
    const a = computeMeshSurfaceAreaWorld(child);
    if(!Number.isFinite(a) || a <= 0) return;
    const diag = uploadClassDiag.get(child.uuid);
    const zi = diag && Number.isFinite(diag.zoneIndex) ? diag.zoneIndex : null;
    const k = zi == null ? 'unzoned' : String(zi);
    out[fk][k] = (out[fk][k] || 0) + a;
  });
  return out;
}

/**
 * Phase 20 — pure helper that rewrites `face.zones` so per-band areas and
 * forces reflect the uploaded mesh distribution rather than the parametric
 * slider geometry. Each band keeps its parametric `Cpe`, `p`, and `dist`
 * label (those depend on h/d, pitch, and wind direction, not on building
 * dimensions); only `area` and `force` are recomputed from `byZoneEntry`.
 *
 * Bands with no uploaded mesh keep area=0 / force=0 so engineers still see
 * the full table-5.2(C) / 5.3(A) banding (a "0 m²" row makes it explicit
 * that no IFC face landed in that band). Unzoned mesh area is returned
 * separately so the face total can still include meshes whose centroid
 * fell outside any band.
 *
 * Returns: { totalArea, totalForce, unzonedArea } (all numbers, never NaN).
 */
function applyZonedAreasToFace(face, byZoneEntry){
  const out = { totalArea: 0, totalForce: 0, unzonedArea: 0 };
  if(!face || !Array.isArray(face.zones) || !byZoneEntry) return out;
  for(let i = 0; i < face.zones.length; i++){
    const z = face.zones[i];
    if(!z) continue;
    const a = byZoneEntry[String(i)];
    const area = Number.isFinite(a) && a > 0 ? a : 0;
    z.area = area;
    const p = Number.isFinite(z.p) ? z.p : 0;
    z.force = Math.abs(p * area);
    out.totalArea += area;
    out.totalForce += z.force;
  }
  const u = byZoneEntry.unzoned;
  if(Number.isFinite(u) && u > 0){
    out.unzonedArea = u;
    out.totalArea += u;
    // Unzoned meshes get the face-average pressure (no band assignment).
    const pAvg = Number.isFinite(face.p) ? face.p : 0;
    out.totalForce += Math.abs(pAvg * u);
  }
  return out;
}

/**
 * Replace parametric (slider-based) face areas in S.R with sums from uploaded
 * mesh geometry. For face-keys with edge-distance bands (sidewall1/2,
 * roof_ww, roof_cw, hip ends), keep `face.zones[]` populated with parametric
 * Cp,e/p but rewrite per-band area/force from the actual mesh distribution
 * (Phase 20). For un-zoned face-keys (windward, leeward, roof_lw), drop
 * `face.zones` and use the simple total-area path.
 */
function patchUploadedModelResultAreas(totals){
  if(!S.R || !S.R.faces || !totals) return;
  let byZone = null;
  try{ byZone = aggregateUploadedMeshAreasByZone(); }catch(_){ byZone = null; }
  for(const k of Object.keys(S.R.faces)){
    const t = totals[k];
    if(t == null || !Number.isFinite(t) || t < 1e-12) continue;
    const face = S.R.faces[k];
    const bz = byZone && byZone[k];
    if(Array.isArray(face.zones) && face.zones.length && bz && Object.keys(bz).length){
      // Phase 20: per-band areas/forces from upload meshes; face total = sum.
      const z = applyZonedAreasToFace(face, bz);
      // If no mesh actually landed in any band (everything 'unzoned'), fall
      // through to the legacy single-pressure path so we don't strand the
      // face with zones[] full of zero areas.
      if(z.totalArea > 0){
        face.area = z.totalArea;
        face.force = z.totalForce;
        continue;
      }
    }
    face.area = t;
    face.force = Math.abs(face.p * t);
    if(face.zones) delete face.zones;
  }
}

/**
 * When the MLP is unsure on tiny IFC shards (openings), still classify envelope from normals + wind
 * if the mesh reads as exterior (outward from plan centroid).
 */
function analyticEnvelopeFallback(mc, bbox, bCenter, meshBox, nAcc, ctx){
  const SEM = MeshClassifier.SEM;
  const bmin = bbox.min, bmax = bbox.max;
  const extY = bmax.y - bmin.y;
  const uy = extY > 1e-6 ? (mc.y - bmin.y) / extY : 0.5;
  const n = MeshClassifier.outwardAlignedNormal(nAcc, mc, bCenter);
  const ay = Math.abs(n.y);
  const horiz = Math.hypot(n.x, n.z);
  const halfH = Math.max(bmax.x - bmin.x, bmax.z - bmin.z) * 0.5;
  const vH = new THREE.Vector3(mc.x - bCenter.x, 0, mc.z - bCenter.z);

  // Roof / top enclosure (sloped or flat)
  if(uy > 0.76 && ay > 0.26){
    const fk = MeshClassifier.semanticToFaceKey(SEM.roof, nAcc, mc, ctx);
    if(fk) return fk;
  }
  // Thin horizontal plates near ground (footing / floor slab) — not curtain wall
  const msz = meshBox.getSize(new THREE.Vector3());
  const nearGround = (meshBox.min.y - bmin.y) < extY * 0.06;
  const minSpan = Math.max(Math.min(msz.x, msz.z), 1e-6);
  const slabLike = msz.y < extY * 0.18 && msz.y < minSpan * 0.22;
  if(nearGround && slabLike && ay < 0.78) return null;
  // Vertical envelope — reject interior partitions (similar normal but center not outside shell)
  if(ay < 0.62 && horiz > 0.32){
    const nH = new THREE.Vector3(n.x, 0, n.z);
    if(nH.lengthSq() < 1e-8) return null;
    nH.normalize();
    if(vH.dot(nH) < halfH * 0.08) return null;
    return MeshClassifier.semanticToFaceKey(SEM.wall, nAcc, mc, ctx);
  }
  return null;
}

function isValidUploadFaceKey(k){
  return typeof k === 'string' && /^(windward|leeward|sidewall1|sidewall2|roof_ww|roof_lw|roof_cw|roof_hip_l|roof_hip_r)$/.test(k);
}

function ifcAiCacheKey(){
  const relW = ((S.windAngle - S.mapBuildingAngle) % 360 + 360) % 360;
  const sec = Math.round(relW / 45) % 8;
  return (uploadedModelName || 'model') + '#w' + sec;
}

function pruneIfcAiCache(){
  while(ifcAiResultCache.size > IFC_AI_CACHE_MAX){
    const first = ifcAiResultCache.keys().next().value;
    ifcAiResultCache.delete(first);
  }
}

function mergeIfcAiFaceResults(faces){
  if(!faces || !faces.length) return;
  for(let i = 0; i < faces.length; i++){
    const row = faces[i];
    if(!row || !row.uuid || !row.faceKey) continue;
    if(!isValidUploadFaceKey(row.faceKey)) continue;
    uploadFaceMap.set(row.uuid, row.faceKey);
  }
}

async function parseIfcAiErrorResponseBody(res){
  const text = await res.text();
  if(!text) return '';
  try{
    const j = JSON.parse(text);
    if(j && typeof j.error === 'string') return j.error;
    if(j && j.error != null) return String(j.error);
  }catch(e){
    /* non-JSON (e.g. HTML error page) */
  }
  return text.length > 400 ? text.slice(0, 397) + '…' : text;
}

async function fetchIfcAiRefinement(payload){
  const url = (typeof window !== 'undefined' && window.WIND_ANALYSIS_IFC_AI_URL) || '';
  if(!url) return null;
  const headers = { 'Content-Type': 'application/json' };
  const authHdrs = await cwDetectionAuthHeaders({ forceRefresh: true });
  Object.assign(headers, authHdrs);
  const res = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(payload)
  });
  if(!res.ok){
    const detail = await parseIfcAiErrorResponseBody(res);
    const msg = detail ? 'IFC AI HTTP ' + res.status + ': ' + detail : 'IFC AI HTTP ' + res.status;
    throw new Error(msg);
  }
  try{
    return await res.json();
  }catch(e){
    throw new Error('IFC AI: response was not valid JSON (' + (e && e.message ? e.message : 'parse error') + ')');
  }
}
/** Split large mesh lists so each POST stays within server IFC_AI_MESH_MAX (one Claude call + one quota slot per chunk). */
async function fetchIfcAiRefinementChunked(payload){
  const meshes = payload.meshes;
  if(!meshes || !meshes.length) return { faces: [] };
  let chunk = IFC_AI_MESH_CHUNK;
  if(typeof window !== 'undefined' && Number.isFinite(window.IFC_AI_MESH_CHUNK)){
    chunk = Math.max(1, Math.min(500, Math.floor(window.IFC_AI_MESH_CHUNK)));
  }
  const base = {
    version: payload.version,
    modelName: payload.modelName,
    wind: payload.wind
  };
  const allFaces = [];
  for(let i = 0; i < meshes.length; i += chunk){
    const part = meshes.slice(i, i + chunk);
    const data = await fetchIfcAiRefinement({ ...base, meshes: part });
    if(data && Array.isArray(data.faces)) allFaces.push(...data.faces);
  }
  return { faces: allFaces };
}

function initIfcAiAssistFromStorage(){
  try{
    const el = document.getElementById('ifc-ai-assist');
    if(!el){
      S.ifcAiAssist = false;
      return;
    }
    const v = sessionStorage.getItem(IFC_AI_SESSION_STORAGE);
    const on = v === '1';
    S.ifcAiAssist = on;
    el.checked = on;
  }catch(e){}
}

function updateIfcAiAssistUI(){
  const el = document.getElementById('ifc-ai-assist');
  const hint = document.getElementById('ifc-ai-assist-hint');
  const allowed = hasPlanFeature('cloudAi');
  if(!el){
    S.ifcAiAssist = false;
    return;
  }
  el.disabled = !allowed;
  if(allowed){
    try{
      const v = sessionStorage.getItem(IFC_AI_SESSION_STORAGE);
      const want = v === '1';
      el.checked = want;
      S.ifcAiAssist = want;
    }catch(e2){
      el.checked = false;
      S.ifcAiAssist = false;
    }
  } else {
    el.checked = false;
    S.ifcAiAssist = false;
  }
  if(hint){
    hint.textContent = allowed
      ? 'When enabled, meshes that are low-confidence or disagree with analytic labels are sent for optional AI review when a local/remote refinement endpoint is configured.'
: 'Cloud AI refinement is optional in this local build.';
  }
}

function onIfcAiAssistChange(){
  const el = document.getElementById('ifc-ai-assist');
  if(!hasPlanFeature('cloudAi')){
    if(el) el.checked = false;
    S.ifcAiAssist = false;
    try{ sessionStorage.setItem(IFC_AI_SESSION_STORAGE, '0'); }catch(e){}
    toast('Cloud AI refinement endpoint is optional in this local build.');
    return;
  }
  S.ifcAiAssist = !!(el && el.checked);
  try{
    sessionStorage.setItem(IFC_AI_SESSION_STORAGE, S.ifcAiAssist ? '1' : '0');
  }catch(e){}
  if(uploadedModelGroup && uploadedModelVisible && !parametricVisible){
    void applyWindToUploadedModel();
  }
}

// Uploaded meshes: MLP semantic class (interior/wall/roof/floor), then analytic normal+wind mapping — no bbox heuristics.
// Optional: user-opt-in cloud refinement for low-confidence / disagreeing meshes (WIND_ANALYSIS_IFC_AI_URL).
async function classifyUploadedFaces(){
  if(!uploadedModelGroup) return;
  uploadFaceMap.clear();
  uploadClassDiag.clear();

  if(typeof MeshClassifier === 'undefined' || !MeshClassifier.isReady() || !MeshClassifier.usesMlp()){
    console.warn('MeshClassifier unavailable — upload wind zones skipped.');
    return;
  }

  try{
  uploadedModelGroup.updateMatrixWorld(true);

  const bbox = new THREE.Box3().setFromObject(uploadedModelGroup);
  const bCenter = bbox.getCenter(new THREE.Vector3());
  const bSize = bbox.getSize(new THREE.Vector3());
  const xzSamples = collectMeshCentersXZFromGroup(uploadedModelGroup);
  const hullCtx = buildUploadFootprintHullContext(xzSamples, bbox);

  // Match buildWindIndicator / calc(): wind relative to building front, then same world rotation as grpWind
  const relW = ((S.windAngle - S.mapBuildingAngle) % 360 + 360) % 360;
  const ang = (relW + (S.R.angleOff || 0)) * Math.PI / 180;
  const windDirLocal = new THREE.Vector3(-Math.sin(ang), 0, -Math.cos(ang));
  const windDir = windDirLocal.clone().applyAxisAngle(new THREE.Vector3(0, 1, 0), -S.mapBuildingAngle * Math.PI / 180);
  windDir.normalize();
  const perpDir = new THREE.Vector3(-windDir.z, 0, windDir.x).normalize();

  const ML_CONF = 0.28;

  const normalMatrix = new THREE.Matrix3();
  const nAcc = new THREE.Vector3();
  const nWorld = new THREE.Vector3();
  const nH = new THREE.Vector3();

  const meshClassifyMeta = new Map();

  uploadedModelGroup.traverse(child => {
    if(!child.isMesh || !child.geometry) return;

    const meshBox = new THREE.Box3().setFromObject(child);
    if(meshBox.isEmpty()) return;
    const mc = meshBox.getCenter(new THREE.Vector3());

    const normAttr = child.geometry.attributes.normal;
    const posAttr = child.geometry.attributes.position;
    if(!normAttr || !posAttr || !normAttr.count) return;

    const nArea = computeMeshNormalForWind(child, bCenter);
    if(nArea){
      nAcc.copy(nArea);
    } else {
      normalMatrix.getNormalMatrix(child.matrixWorld);
      nAcc.set(0, 0, 0);
      const nv = normAttr.count;
      for(let i = 0; i < nv; i++){
        nWorld.fromBufferAttribute(normAttr, i);
        nWorld.applyMatrix3(normalMatrix).normalize();
        nAcc.add(nWorld);
      }
      nAcc.divideScalar(nv);
      const alen = nAcc.length();
      if(alen <= 1e-4) return;
      nAcc.divideScalar(alen);
    }

    if(geoSkipNonEnvelopeMesh(mc, bbox, bCenter, nAcc, meshBox, hullCtx)) return;

    // Phase 1a: use classifyMesh wrapper — runs IFC-first short-circuit before the MLP.
    // Falls back cleanly to extractFeatures+inferProbs for meshes without definitive IFC tags
    // (IFCMEMBER, IFCBEAM, IFCCOLUMN, IFCPLATE, IFCCOVERING, IFCBUILDINGELEMENTPROXY, and
    // mid-height IFCSLAB). Result shape matches the legacy two-step call.
    let classRes = MeshClassifier.classifyMesh
      ? MeshClassifier.classifyMesh(child, bbox, bCenter, bSize, mc, meshBox, nAcc)
      : (function(){
          const f = MeshClassifier.extractFeatures(child, bbox, bCenter, bSize, mc, meshBox, nAcc);
          return { probs: MeshClassifier.inferProbs(f), viaIfc: false, ifcClass: null, ifcConf: 0, features: f };
        })();

    // Phase 3: hull-aware post-processing — conservative shell/interior overrides for
    // soft MLP predictions. IFC short-circuit results are never overridden.
    let hullEdgeDist = Infinity;
    if(hullCtx && hullCtx.hull && hullCtx.hull.length >= 3){
      hullEdgeDist = distPointToPolygonBoundaryXZ(mc.x, mc.z, hullCtx.hull);
    }
    if(MeshClassifier.applyHullRefinement){
      classRes = MeshClassifier.applyHullRefinement(classRes, bbox, bCenter, mc, meshBox, nAcc, hullEdgeDist);
    }
    const probs = classRes.probs;
    let best = 0;
    let bestP = probs[0];
    for(let pi = 1; pi < probs.length; pi++){
      if(probs[pi] > bestP){ bestP = probs[pi]; best = pi; }
    }

    // Phase 9: capture per-call face-key confidence via ctx.confidenceOut.
    // semanticToFaceKey writes ctx.confidenceOut.lastConfidence ∈ [0,1] —
    // 0 = on a decision boundary, 1 = unambiguous (see Phase 8c in mesh-classifier.js).
    const confidenceOut = { lastConfidence: 0 };
    const ctx = { windDir, perpDir, bCenter, nH, confidenceOut };
    const SEM = MeshClassifier.SEM;
    const analyticKey = analyticEnvelopeFallback(mc, bbox, bCenter, meshBox, nAcc, ctx);
    const analyticConf = analyticKey ? confidenceOut.lastConfidence : 0;
    let faceKey = null;
    let mlpFaceConf = 0;
    if(bestP >= ML_CONF && best === SEM.wall){
      confidenceOut.lastConfidence = 0;
      faceKey = MeshClassifier.semanticToFaceKey(SEM.wall, nAcc, mc, ctx);
      mlpFaceConf = faceKey ? confidenceOut.lastConfidence : 0;
    } else if(bestP >= ML_CONF && best === SEM.roof){
      confidenceOut.lastConfidence = 0;
      faceKey = MeshClassifier.semanticToFaceKey(SEM.roof, nAcc, mc, ctx);
      mlpFaceConf = faceKey ? confidenceOut.lastConfidence : 0;
    }
    let faceKeyConfidence = mlpFaceConf;
    if(!faceKey){ faceKey = analyticKey; faceKeyConfidence = analyticConf; }

    const hullNear = isFinite(hullEdgeDist) && hullCtx && hullCtx.band
      ? hullEdgeDist < hullCtx.band : false;
    const pp = classRes.postPass || null;
    // Phase 4: cache the 22-d feature vector so face-key overrides can emit JSONL
    // training rows without re-running feature extraction.
    const featArr = classRes.features
      ? (Array.isArray(classRes.features) ? classRes.features.slice() : Array.from(classRes.features))
      : null;
    // Phase 13: assign distance-from-windward-edge zone for sidewall/roof_ww/
    // hip-end meshes. The parametric pipeline has already populated
    // S.R.faces[fk].zones[] with Cp,e per band (Table 5.2(C) for sidewalls,
    // Table 5.3(A) for roof_ww). We just walk cumulative widths to find which
    // band the mesh centroid lives in. No-op for windward/leeward/roof_lw/
    // roof_cw — those are uniform or use a different zone scheme.
    const zonesForFk = (faceKey && S.R && S.R.faces && S.R.faces[faceKey] && Array.isArray(S.R.faces[faceKey].zones))
      ? S.R.faces[faceKey].zones : null;
    const zoneIndex = (faceKey && zonesForFk)
      ? assignWindEdgeZone(faceKey, mc, bCenter, bSize, windDir, zonesForFk)
      : null;
    uploadClassDiag.set(child.uuid, {
      mlpClass: IFC_AI_CLASS_LABELS[best] || String(best),
      mlpP: bestP,
      analyticKey: analyticKey || null,
      analyticConf,
      faceKey: faceKey || null,
      // Phase 9: confidence ∈ [0,1] of the chosen face-key. Low values flag
      // walls/roofs sitting near a decision boundary (e.g. ~45° to wind for
      // walls, near flatPitchDeg threshold for roofs). UI/diagnostics can
      // surface low-conf meshes for human override.
      faceKeyConfidence,
      // Phase 13: zone index into S.R.faces[faceKey].zones[]. null when the
      // face-key has no edge-distance zoning (windward, leeward, roof_lw, roof_cw).
      zoneIndex,
      viaIfc: !!classRes.viaIfc,
      ifcConf: classRes.ifcConf || 0,
      hullNear,
      postOverride: pp ? pp.override : null,
      shellScore: pp ? pp.shellScore : 0,
      features: featArr
    });

    if(faceKey) uploadFaceMap.set(child.uuid, faceKey);
    meshClassifyMeta.set(child.uuid, {
      bestP,
      best,
      analyticKey: analyticKey || null,
      // Phase 9: stash face-key confidence so the cloud-AI refinement gate
      // can flag boundary-ambiguous walls/roofs without re-running classification.
      faceKeyConfidence
    });
  });

  applyCoplanarFaceKeyConsensus(uploadedModelGroup, uploadFaceMap, bCenter);

  // Phase 10: aggregate face-key confidence so we can flag ambiguous classifications
  // at the summary level. A wall whose normal sits ~45° to the wind, or a roof whose
  // pitch sits at the flat/sloped boundary, gets a confidence near 0 — surfacing the
  // count gives users one number to glance at instead of hovering every mesh.
  let lowFaceConfCount = 0;
  let classifiedFaceCount = 0;
  uploadClassDiag.forEach(d => {
    if(!d || !d.faceKey) return;
    classifiedFaceCount++;
    const c = d.faceKeyConfidence;
    if(Number.isFinite(c) && c < IFC_AI_REFINE_FACE_CONF_MIN) lowFaceConfCount++;
  });
  const lowFaceTag = classifiedFaceCount > 0 && lowFaceConfCount > 0
    ? ' · ' + lowFaceConfCount + ' low-confidence (~boundary)'
    : '';
  console.log('Upload classification:', uploadFaceMap.size, 'meshes (MLP + analytic wind) of',
    countMeshes(uploadedModelGroup), 'total' + lowFaceTag);
  // Only toast on meaningful counts: 1 ambiguous mesh in 200 isn't worth
  // interrupting the user. Fire when either (a) >=3 ambiguous AND >=5% of faces,
  // or (b) >=10% regardless of absolute count (catches small models).
  if(classifiedFaceCount > 0 && lowFaceConfCount > 0){
    const frac = lowFaceConfCount / classifiedFaceCount;
    const noteworthy = (lowFaceConfCount >= 3 && frac >= 0.05) || frac >= 0.10;
    if(noteworthy){
      const pct = Math.round(frac * 100);
      toast(lowFaceConfCount + ' of ' + classifiedFaceCount + ' faces flagged as ambiguous (' + pct + '%) — hover for details.');
    }
  }

  const cloudMeshRows = [];
  if(S.ifcAiAssist){
    uploadedModelGroup.traverse(child => {
      if(!child.isMesh || !child.geometry) return;
      const meta = meshClassifyMeta.get(child.uuid);
      if(!meta) return;
      const fkNow = uploadFaceMap.get(child.uuid) || null;
      if(!meshNeedsIfcAiRefinement(meta.bestP, fkNow, meta.analyticKey, meta.faceKeyConfidence)) return;
      const meshBox = new THREE.Box3().setFromObject(child);
      if(meshBox.isEmpty()) return;
      const mc = meshBox.getCenter(new THREE.Vector3());
      const nArea = computeMeshNormalForWind(child, bCenter);
      const nOut = nArea ? nArea.clone() : new THREE.Vector3(0, 1, 0);
      const ud = child.userData || {};
      cloudMeshRows.push({
        uuid: child.uuid,
        expressID: ud.ifcExpressID != null ? ud.ifcExpressID : null,
        ifcLineType: ud.ifcLineType != null ? ud.ifcLineType : null,
        centroid: [mc.x, mc.y, mc.z],
        normal: [nOut.x, nOut.y, nOut.z],
        area: computeMeshSurfaceAreaWorld(child),
        mlpBestP: meta.bestP,
        mlpClass: IFC_AI_CLASS_LABELS[meta.best] || String(meta.best),
        localFaceKey: fkNow,
        // Phase 9: 'low_face_conf' surfaces meshes whose wall/roof class is
        // confident but whose face-key sits near a decision boundary (windward
        // vs sidewall at ~45°, flat vs sloped at the pitch threshold).
        refineReason: !Number.isFinite(meta.bestP)
          ? 'no_conf'
          : (meta.bestP <= IFC_AI_REFINE_CONF_MAX
            ? 'low_conf'
            : (fkNow && meta.analyticKey && fkNow !== meta.analyticKey
              ? 'mlp_analytic_mismatch'
              : (Number.isFinite(meta.faceKeyConfidence) && meta.faceKeyConfidence < IFC_AI_REFINE_FACE_CONF_MIN
                ? 'low_face_conf'
                : 'other')))
      });
    });
  }

  if(!S.ifcAiAssist || !cloudMeshRows.length){
    if(S.ifcAiAssist && meshClassifyMeta.size > 0 && cloudMeshRows.length === 0){
      console.log('Cloud AI: no meshes need refinement (selective mode).');
    }
    return;
  }

  if(!hasPlanFeature('cloudAi')){
    return;
  }
  if(!getActiveFirebaseUser()){
    try{
      if(sessionStorage.getItem('sw_ifc_ai_auth_warned') !== '1'){
        sessionStorage.setItem('sw_ifc_ai_auth_warned', '1');
        toast('Cloud AI refinement is not enabled in this local build.');
      }
    }catch(e3){
      toast('Cloud AI refinement is not enabled in this local build.');
    }
    return;
  }

  const apiUrl = (typeof window !== 'undefined' && window.WIND_ANALYSIS_IFC_AI_URL) || '';
  if(!apiUrl){
    try{
      if(sessionStorage.getItem('sw_ifc_ai_url_warned') !== '1'){
        sessionStorage.setItem('sw_ifc_ai_url_warned', '1');
        toast('Cloud AI: set window.WIND_ANALYSIS_IFC_AI_URL to your classifier endpoint.');
      }
    }catch(e2){
      toast('Cloud AI: set window.WIND_ANALYSIS_IFC_AI_URL to your classifier endpoint.');
    }
    return;
  }

  const ckey = ifcAiCacheKey();
  if(ifcAiResultCache.has(ckey)){
    mergeIfcAiFaceResults(ifcAiResultCache.get(ckey));
    return;
  }

  try{
    let effChunk = IFC_AI_MESH_CHUNK;
    if(typeof window !== 'undefined' && Number.isFinite(window.IFC_AI_MESH_CHUNK)){
      effChunk = Math.max(1, Math.min(500, Math.floor(window.IFC_AI_MESH_CHUNK)));
    }
    const nChunk = Math.ceil(cloudMeshRows.length / effChunk);
    toast(nChunk > 1
      ? 'Refining meshes (cloud, ' + nChunk + ' batches)…'
      : 'Refining meshes with Cloud AI…');
    const payload = {
      version: 1,
      modelName: uploadedModelName || '',
      wind: { relDeg: relW, mapBuildingAngle: S.mapBuildingAngle || 0 },
      building: {
        extentM: { x: bSize.x, y: bSize.y, z: bSize.z },
        footprintHullXZ: hullCtx && hullCtx.hull && hullCtx.hull.length >= 3
          ? hullCtx.hull.map(p => [p[0], p[1]])
          : []
      },
      meshes: cloudMeshRows
    };
    const data = await fetchIfcAiRefinementChunked(payload);
    if(data && Array.isArray(data.faces)){
      if(data.faces.length > 0){
        mergeIfcAiFaceResults(data.faces);
        ifcAiResultCache.set(ckey, data.faces);
        pruneIfcAiCache();
      }else if(cloudMeshRows.length > 0){
        toast('Cloud AI returned no face updates (model filtered all labels). Local classification unchanged — try again or check Netlify logs.');
      }
    }
  }catch(err){
    console.warn('IFC AI refinement failed:', err);
    let msg = err && err.message ? String(err.message) : String(err);
    if(msg.length > 300) msg = msg.slice(0, 297) + '…';
    toast('Cloud AI refinement failed — ' + msg + ' (using local classification only.)');
  }
  }finally{
    applyUploadFaceOverridesFromStorage(uploadFaceMap);
    uploadFaceMap.forEach((fk, uuid) => {
      const d = uploadClassDiag.get(uuid);
      if(d) d.faceKey = fk;
    });
  }
}

function countMeshes(group){
  let n = 0;
  group.traverse(c => { if(c.isMesh) n++; });
  return n;
}

// Recolor uploaded model meshes based on their face classification
/**
 * Phase 11: mix `baseColor` toward amber (0xffaa00) at `t` ∈ [0,1].
 * Used to flag meshes whose face-key confidence sits below the boundary
 * threshold — the caution-tape tint keeps the underlying face-key colour
 * identity (still reads as a windward wall, just visibly uncertain).
 */
function tintLowConfColor(baseColor, t){
  const tt = t < 0 ? 0 : (t > 1 ? 1 : t);
  const r = (baseColor >> 16) & 0xff;
  const g = (baseColor >> 8)  & 0xff;
  const b =  baseColor        & 0xff;
  const ar = 0xff, ag = 0xaa, ab = 0x00;
  const nr = Math.round(r * (1 - tt) + ar * tt);
  const ng = Math.round(g * (1 - tt) + ag * tt);
  const nb = Math.round(b * (1 - tt) + ab * tt);
  return (nr << 16) | (ng << 8) | nb;
}

function recolorUploadedModel(){
  if(!uploadedModelGroup || !S.R.faces) return;

  const F = S.R.faces;
  const showHeat = S.showHeatmap;
  const showPMap = S.showPressureMap;
  const isTr = S.viewMode === 'transparent';
  const isWf = S.viewMode === 'wireframe';

  // Default colors per face type (matches parametric model)
  const baseColors = {
    windward: 0x4488cc, leeward: 0x44aa66,
    sidewall1: 0xcc8844, sidewall2: 0xcc8844,
    roof_ww: 0xcc6622, roof_lw: 0x996633
  };

  uploadedModelGroup.traverse(child => {
    if(!child.isMesh) return;
    child.frustumCulled = false;

    const fk = uploadFaceMap.get(child.uuid);

    if(!fk){
      faceMap.delete(child.uuid);
      if(showHeat || showPMap){
        if(!uploadOrigMaterials.has(child.uuid)){
          uploadOrigMaterials.set(child.uuid, child.material);
        }
        child.material = new THREE.MeshBasicMaterial({
          color: 0x6a7582,
          transparent: isTr,
          opacity: isTr ? 0.4 : 0.92,
          wireframe: isWf,
          side: THREE.DoubleSide,
          depthTest: true,
          depthWrite: !isTr
        });
        return;
      }
      const orig = uploadOrigMaterials.get(child.uuid);
      if(orig){
        child.material = orig;
        uploadOrigMaterials.delete(child.uuid);
      }
      return;
    }

    // Preserve original material for later restoration
    if(!uploadOrigMaterials.has(child.uuid)){
      uploadOrigMaterials.set(child.uuid, child.material);
    }

    const faceData = F[fk];
    if(!faceData) return;

    // Phase 13: zone-aware colour. If the mesh has a zoneIndex pointing into
    // faceData.zones[], use that band's pressure rather than the area-weighted
    // average. Gives IFC sidewall1 / roof_ww meshes the same banded heatmap
    // the parametric model already shows.
    const diagForFk = uploadClassDiag.get(child.uuid);
    const zIdx = diagForFk && Number.isFinite(diagForFk.zoneIndex) ? diagForFk.zoneIndex : null;
    const zoneEntry = (zIdx !== null && Array.isArray(faceData.zones) && faceData.zones[zIdx])
      ? faceData.zones[zIdx] : null;
    const zonalP = zoneEntry && Number.isFinite(zoneEntry.p) ? zoneEntry.p : null;

    let col;
    if(showHeat || showPMap){
      // Heatmap/pressure modes encode pressure value in colour — leave them
      // alone. The toast count + inspector tooltip still surface low-conf
      // meshes in those views.
      const pForHeat = zonalP != null ? zonalP : faceData.p;
      col = Number.isFinite(pForHeat) ? heatCol(pForHeat) : (baseColors[fk] || 0x888888);
    } else {
      col = baseColors[fk] || 0x888888;
      // Phase 11: amber tint flags low face-key confidence in classification
      // view. uploadClassDiag is the source of truth — it carries the
      // confidence we computed during classification.
      const diag = uploadClassDiag.get(child.uuid);
      const fkc = diag ? diag.faceKeyConfidence : null;
      if(Number.isFinite(fkc) && fkc < IFC_AI_REFINE_FACE_CONF_MIN){
        col = tintLowConfColor(col, 0.3);
      }
    }

    // MeshBasicMaterial: heat/pressure colours must read on dark backgrounds without relying on lights
    if(showHeat || showPMap){
      child.material = new THREE.MeshBasicMaterial({
        color: col,
        transparent: isTr,
        opacity: isTr ? 0.4 : 0.9,
        wireframe: isWf,
        side: THREE.DoubleSide,
        depthTest: true,
        depthWrite: !isTr
      });
    } else {
      child.material = new THREE.MeshStandardMaterial({
        color: col,
        roughness: 0.6, metalness: 0.1,
        transparent: isTr,
        opacity: isTr ? 0.4 : 0.92,
        wireframe: isWf,
        side: THREE.DoubleSide
      });
    }

    // Register in faceMap — per-mesh area from geometry; omit zones (parametric strip bands only)
    const meshArea = computeMeshSurfaceAreaWorld(child);
    const fdHover = { ...faceData };
    delete fdHover.zones;
    // Phase 13: when this mesh has a zone, override the hover-shown Cp,e and
    // pressure with the zone's values so the inspector reports the band the
    // mesh actually sits in (not the area-weighted face average).
    const zonePerMeshP = zonalP != null ? zonalP : faceData.p;
    const zoneCpe = zoneEntry && Number.isFinite(zoneEntry.Cpe) ? zoneEntry.Cpe : faceData.Cp_e;
    const zoneLabel = zoneEntry && zoneEntry.dist ? zoneEntry.dist : null;
    // Phase 20: surface the band label in the inspector title so the IFC
    // upload model reads "Side Wall L (1h to 2h)" — matching the parametric
    // model's strip naming. The face-average name remains the fallback when
    // the mesh isn't zoned (windward/leeward/roof_lw, or off-band centroid).
    const baseName = (fdHover.name != null) ? String(fdHover.name) : '';
    const labeledName = zoneLabel ? (baseName + ' (' + zoneLabel + ')') : baseName;
    faceMap.set(child.uuid, {
      ...fdHover,
      key: fk,
      name: labeledName,
      // Per-mesh zone metadata (Phase 13). Null when the face-key has no
      // edge-distance zoning. The inspector tooltip prefers these over the
      // area-weighted face values when present.
      zoneIndex: zIdx,
      zoneLabel,
      zoneCpe,
      zonalP,
      // Override p / Cp_e with band-specific values so force computations and
      // tooltip show the per-mesh band, not the face average.
      p: zonePerMeshP,
      Cp_e: zoneCpe,
      area: meshArea,
      force: Math.abs((zonePerMeshP != null ? zonePerMeshP : faceData.p) * meshArea)
    });
  });
}

// Restore uploaded model to its original materials
function restoreUploadedMaterials(){
  if(!uploadedModelGroup) return;
  uploadedModelGroup.traverse(child => {
    if(!child.isMesh) return;
    const orig = uploadOrigMaterials.get(child.uuid);
    if(orig) child.material = orig;
  });
  // Remove from faceMap so hover tooltips stop
  uploadFaceMap.forEach((fk, uuid) => faceMap.delete(uuid));
}

// Build dimension lines on the uploaded model bounding box
function buildUploadedDims(){
  if(!uploadBBox || !grpUploadOverlay) return;
  const w = uploadBBox.width, d = uploadBBox.depth, h = uploadBBox.height;
  const off = 2.5;
  // IFC view request: suppress text labels on model overlays; keep guide lines only.
  dimLineNoLabel(new THREE.Vector3(-w/2,-.4,d/2+off), new THREE.Vector3(w/2,-.4,d/2+off), 0xffaa44, grpUploadOverlay);
  dimLineNoLabel(new THREE.Vector3(w/2+off,-.4,d/2), new THREE.Vector3(w/2+off,-.4,-d/2), 0x44aaff, grpUploadOverlay);
  dimLineNoLabel(new THREE.Vector3(w/2+off,0,d/2+off), new THREE.Vector3(w/2+off,h,d/2+off), 0x44ff88, grpUploadOverlay);
}

// Build zone labels on the uploaded model
function buildUploadedLabels(){
  // IFC view request: hide overlay text labels (wall role labels + pressure text).
  return;
}

// Build pressure arrows on uploaded model bounding box
function buildUploadedPArrows(){
  if(!uploadBBox || !grpUploadOverlay || !S.R.faces) return;
  const w = uploadBBox.width, d = uploadBBox.depth, h = uploadBBox.height;
  const F = S.R.faces;
  const fm = getWindFaceMap();
  const mx = Math.max(...Object.values(F).map(f=>Math.abs(f.p)), .1);

  function ar(o, dir, p, ml){
    const len = Math.max(Math.abs(p)/mx*ml, .3);
    const c = p>0 ? 0xff4444 : 0x4444ff;
    const dd = dir.clone().normalize(); if(p<0) dd.negate();
    grpUploadOverlay.add(new THREE.ArrowHelper(dd, o, len, c, len*.2, len*.12));
  }
  // Front/back pressure arrows — use wind role from face map
  for(let r=0;r<3;r++) for(let c=0;c<3;c++){
    const fx=(c+.5)/3-.5, fy=(r+.5)/3;
    ar(new THREE.Vector3(fx*w,fy*h,d/2), new THREE.Vector3(0,0,-1), F[fm.front].p, 3);
    ar(new THREE.Vector3(fx*w,fy*h,-d/2), new THREE.Vector3(0,0,1), F[fm.back].p, 3);
  }
  // Left/right pressure arrows
  for(let r=0;r<3;r++) for(let c=0;c<3;c++){
    const fz=(c+.5)/3-.5, fy=(r+.5)/3;
    ar(new THREE.Vector3(-w/2,fy*h,fz*d), new THREE.Vector3(1,0,0), F[fm.left].p, 3);
    ar(new THREE.Vector3(w/2,fy*h,fz*d), new THREE.Vector3(-1,0,0), F[fm.right].p, 3);
  }
  // Roof arrows — swap based on windward slope direction
  const frontIsWW = (fm.back !== 'windward');
  for(let r=0;r<2;r++) for(let c=0;c<3;c++){
    const fx=(c+.5)/3-.5;
    const frontP = frontIsWW ? F.roof_ww.p : F.roof_lw.p;
    const backP  = frontIsWW ? F.roof_lw.p : F.roof_ww.p;
    ar(new THREE.Vector3(fx*w,h+.5,(r===0?1:-1)*d/4), new THREE.Vector3(0,1,0), r===0?frontP:backP, 3);
  }
}

// Master function: apply full wind analysis to uploaded model
async function applyWindToUploadedModel(){
  if(!uploadedModelGroup || !uploadedModelVisible) return;
  if(!S.R.faces) calc();

  // Create/clear the overlay group
  if(grpUploadOverlay){
    clearGrp(grpUploadOverlay);
  } else {
    grpUploadOverlay = new THREE.Group();
    scene.add(grpUploadOverlay);
  }

  try{
    // 1. Classify faces by normal vs wind direction (optional async cloud refinement)
    await classifyUploadedFaces();

    // 1b. Face areas/forces from actual mesh geometry (not parametric width/depth sliders)
    patchUploadedModelResultAreas(aggregateUploadedMeshAreasByFaceKey());

    // 2. Recolor meshes by their face classification
    recolorUploadedModel();
  } catch(err){
    console.warn('Uploaded wind visualization failed (model may still be visible):', err);
    try{
      recolorUploadedModel();
    } catch(e2){
      console.warn('recolorUploadedModel fallback:', e2);
    }
  }

  // 3. Dimension lines
  if(S.showDimensions) buildUploadedDims();

  // 4. Zone labels
  if(S.showLabels) buildUploadedLabels();

  // 5. Pressure arrows
  if(S.showPressureArrows) buildUploadedPArrows();

  grpUploadOverlay.visible = true;
  if(typeof updateDocPreview === 'function') updateDocPreview();
}

/** Unlit materials so uploaded IFC/GLB is visible before async wind/heatmap runs (avoids dark PBR + lighting issues). */
function applyUploadedModelUnlitFallback(){
  if(!uploadedModelGroup) return;
  uploadedModelGroup.traverse(function(child){
    if(!child.isMesh) return;
    child.visible = true;
    child.frustumCulled = false;
    child.renderOrder = 0;
    if(child.material && typeof child.material.dispose === 'function'){
      try{ child.material.dispose(); }catch(e){ /* ignore */ }
    }
    child.material = new THREE.MeshBasicMaterial({
      color: 0x7a8fa8,
      side: THREE.DoubleSide,
      depthTest: true,
      depthWrite: true,
      fog: false
    });
  });
}

// ── Add model to scene ──
// opts.unlitFallback: IFC-only — replace materials with unlit grey so geometry shows before heatmap (textured GLB keeps original mats).
function addUploadedModel(object3d, name, opts){
  opts = opts || {};
  // Remove previous uploaded model
  if(uploadedModelGroup){
    scene.remove(uploadedModelGroup);
    disposeGroup(uploadedModelGroup);
  }
  ifcAiResultCache.clear();

  uploadedModelGroup = new THREE.Group();
  uploadedModelGroup.add(object3d);
  object3d.updateMatrixWorld(true);

  // Auto-scale and position the model
  const box = new THREE.Box3().setFromObject(object3d);
  if(box.isEmpty()){
    console.warn('addUploadedModel: empty bounding box');
    toast('⚠ Model loaded but has no visible geometry — check the file or try GLB/OBJ.');
  }
  const size = box.getSize(new THREE.Vector3());
  const center = box.getCenter(new THREE.Vector3());
  const maxDim = Math.max(size.x, size.y, size.z);
  if(!Number.isFinite(maxDim) || maxDim <= 0){
    console.error('addUploadedModel: invalid extent', { size, maxDim });
    toast('⚠ Model has invalid size — try re-exporting IFC or use GLB/OBJ.');
  }

  // Target size: match the parametric building's approximate footprint (sliders)
  const targetMax = Math.max(S.width || 20, S.depth || 15, S.height || 6) * 1.2;

  // Scale the model so its largest dimension matches the target
  let scale = 1;
  if(maxDim > 0.001 && Number.isFinite(maxDim)){
    scale = Math.min(1e12, Math.max(1e-12, targetMax / maxDim));
  }

  // Log for debugging
  console.log('Upload scale:', { maxDim, targetMax, scale, sizeX: size.x, sizeY: size.y, sizeZ: size.z });

  if(Math.abs(scale - 1) > 0.01){
    object3d.scale.multiplyScalar(scale);
    box.setFromObject(object3d);
    box.getSize(size);
    box.getCenter(center);
  }

  // Center the model on the ground plane at origin
  object3d.position.sub(center);
  // Lift slightly so thin slabs do not z-fight with the ground plane
  object3d.position.y += size.y / 2 + 0.02;

  // Store bounding box for wind analysis overlays
  uploadBBox = { width: size.x, depth: size.z, height: size.y };

  uploadModelCamExtent = Math.max(size.x, size.y, size.z, 1e-9);

  scene.add(uploadedModelGroup);
  uploadedModelGroup.visible = true;
  if(opts.unlitFallback){
    applyUploadedModelUnlitFallback();
  } else {
    uploadedModelGroup.traverse(function(o){
      if(o.isMesh){
        o.visible = true;
        o.frustumCulled = false;
      }
    });
  }
  uploadedModelVisible = true;
  uploadedModelName = name;

  // Hide parametric model — uploaded model takes over
  parametricVisible = false;
  if(typeof grpBuild !== 'undefined' && grpBuild) grpBuild.visible = false;
  if(typeof grpDim !== 'undefined' && grpDim) grpDim.visible = false;
  if(typeof grpLabel !== 'undefined' && grpLabel) grpLabel.visible = false;
  if(typeof grpArrows !== 'undefined' && grpArrows) grpArrows.visible = false;
  if(typeof grpInternal !== 'undefined' && grpInternal) grpInternal.visible = false;

  // Show toolbar controls
  document.getElementById('btn-toggle-model').style.display = '';
  document.getElementById('btn-remove-model').style.display = '';
  const btn = document.getElementById('btn-toggle-model');
  if(btn) btn.textContent = '🔀 Show Parametric';

  const dims = `${size.x.toFixed(1)} × ${size.z.toFixed(1)} × ${size.y.toFixed(1)} m`;
  showModelStatus('success', name, dims);
  toast('✅ Model loaded: ' + name);

  // Fit camera to model (do not floor extent — OrbitControls minDistance 4 would hide tiny models)
  const camExtent = Math.max(size.x, size.y, size.z, 1e-9);
  if(Number.isFinite(camExtent) && camExtent > 1e-8){
    fitOrbitControlsToModelExtent(camExtent);
    camera.near = Math.max(0.01, camExtent * 1e-4);
    camera.far = Math.max(2500, camExtent * 45);
    camera.updateProjectionMatrix();
    camera.position.set(camExtent * 1.5, camExtent * 1.0, camExtent * 1.8);
    controls.target.set(0, size.y / 2, 0);
    controls.update();
  } else {
    resetCamera();
  }
  setSceneBg();

  console.log('Model loaded:', name, 'Size:', dims, 'Scale applied:', scale);

  // Apply wind analysis (face classification, coloring, labels, dimensions)
  (async function(){
    try{
      if(typeof MeshClassifier !== 'undefined'){
        try { await MeshClassifier.ensureReady(); } catch(e) { console.warn('MeshClassifier.ensureReady', e); }
      }
      await applyWindToUploadedModel();
    } catch(err){
      console.warn('Post-upload wind pipeline:', err);
    }
  })();
}

function toggleUploadedModel(){
  if(!uploadedModelGroup) return;
  uploadedModelVisible = !uploadedModelVisible;
  uploadedModelGroup.visible = uploadedModelVisible;

  // Also toggle parametric model
  parametricVisible = !uploadedModelVisible;
  grpBuild.visible = parametricVisible;
  grpDim.visible = parametricVisible;
  grpLabel.visible = parametricVisible;
  grpArrows.visible = parametricVisible;
  grpInternal.visible = parametricVisible;

  // Toggle upload overlay
  if(grpUploadOverlay) grpUploadOverlay.visible = uploadedModelVisible;

  const btn = document.getElementById('btn-toggle-model');
  if(uploadedModelVisible){
    btn.textContent = '🔀 Show Parametric';
    showModelStatus('success', uploadedModelName, 'Uploaded model shown');
    // Re-apply wind analysis when switching to uploaded
    void applyWindToUploadedModel();
  } else {
    btn.textContent = '🔀 Show Uploaded';
    showModelStatus('success', uploadedModelName, 'Parametric model shown');
    // Restore original materials on uploaded model
    restoreUploadedMaterials();
    // Restore S.R face areas/forces from parametric dimensions (sliders)
    calc();
    rebuild();
    recalcPressures();
  }
  setSceneBg();
}

function removeUploadedModel(){
  if(!uploadedModelGroup) return;
  if(!confirm('Remove uploaded model "' + uploadedModelName + '"?')) return;

  scene.remove(uploadedModelGroup);
  disposeGroup(uploadedModelGroup);
  uploadedModelGroup = null;
  uploadedModelVisible = false;
  uploadedModelName = '';
  ifcAiResultCache.clear();

  // Clean up wind analysis state
  if(grpUploadOverlay){
    clearGrp(grpUploadOverlay);
    scene.remove(grpUploadOverlay);
    grpUploadOverlay = null;
  }
  uploadFaceMap.clear();
  uploadClassDiag.clear();
  uploadOrigMaterials.clear();
  uploadBBox = null;
  uploadModelCamExtent = 0;
  lastUploadHoverUuid = null;

  // Restore parametric model visibility
  parametricVisible = true;
  grpBuild.visible = true;
  grpDim.visible = true;
  grpLabel.visible = true;
  grpArrows.visible = true;
  grpInternal.visible = true;

  // Hide toolbar controls
  document.getElementById('btn-toggle-model').style.display = 'none';
  document.getElementById('btn-remove-model').style.display = 'none';

  // Clear status text
  const statusText = document.getElementById('model-status-text');
  if(statusText){ statusText.style.display = 'none'; statusText.textContent = ''; }
  const statusInfo = document.getElementById('model-status-info');
  if(statusInfo) statusInfo.textContent = '';

  // Reset camera
  resetCamera();
  setSceneBg();
  calc();
  rebuild();
  recalcPressures();
  toast('Model removed');
}

function disposeGroup(group){
  group.traverse(child => {
    if(child.geometry) child.geometry.dispose();
    if(child.material){
      if(Array.isArray(child.material)) child.material.forEach(m => m.dispose());
      else child.material.dispose();
    }
  });
}

function showModelStatus(type, text, info){
  const textEl = document.getElementById('model-status-text');
  const infoEl = document.getElementById('model-status-info');

  if(textEl){
    textEl.style.display = text ? 'inline' : 'none';
    if(type === 'loading') textEl.textContent = '⏳ ' + (text || '');
    else if(type === 'success') textEl.textContent = '✅ ' + (text || '');
    else if(type === 'error') textEl.textContent = '❌ ' + (text || '');
    else textEl.textContent = text || '';
  }
  if(infoEl) infoEl.textContent = info || '';
}

if(typeof MeshClassifier !== 'undefined'){
  MeshClassifier.ensureReady().then(function(){
    if(uploadedModelGroup && uploadedModelVisible && !parametricVisible) void applyWindToUploadedModel();
  }).catch(function(){});
}

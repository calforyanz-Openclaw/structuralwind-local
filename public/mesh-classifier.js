/**
 * Mesh semantic classifier: MLP only (trained once at load in-browser, no bbox heuristics).
 * Optional: fetch mesh-classifier-weights.json to override embedded training output.
 */
(function (global) {
  'use strict';

  var SEM = { interior: 0, wall: 1, roof: 2, floor: 3 };
  // v2 feature layout (22-d):
  //   f0..f2   centroid offset in normalised bbox frame, clamped ±1.5
  //   f3..f5   mesh size / max bbox extent, clamped [0,2]
  //   f6..f8   mean outward-aligned normal (x,y,z)
  //   f9       horizontal normal magnitude  = hypot(f6,f8)
  //   f10      height fraction hfrac = (mc.y - bmin.y) / (bmax.y - bmin.y)
  //   f11      log(volR*10+1)/3  — mesh volume ratio proxy (log-compressed)
  //   f12..f19 IFC 8-class one-hot:
  //              12=wall 13=curtain 14=opening 15=roof 16=slab 17=interior
  //              18=beam/member 19=proxy/covering/footing  (each in [0..1], scaled by conf)
  //   f20      vertical aspect = msz.y / max(msz.x, msz.z), clamped [0,4]
  //   f21      plan aspect     = min(msz.x, msz.z) / max(msz.x, msz.z), in [0,1]
  // Backward compat: if fetched weights are v1 (14-d), extractFeatures still emits 22 but
  // loadWeightsJson will pad/mask; if v2 weights are present they are used directly.
  var N_FEATURES = 22;
  var N_HIDDEN = 48;
  var N_CLASSES = 4;
  var IFC_ONEHOT_DIM = 8;

  var state = {
    ready: false,
    mlp: null
  };

  function relu(x) {
    return x > 0 ? x : 0;
  }

  function softmax(logits) {
    var m = Math.max.apply(null, logits);
    var ex = logits.map(function (z) { return Math.exp(z - m); });
    var s = ex.reduce(function (a, b) { return a + b; }, 0);
    return ex.map(function (e) { return e / s; });
  }

  function mlpForward(x, w) {
    var fin = w.inputDim;
    var hid = w.hiddenDim;
    var nclass = w.numClasses;
    var W1 = w.W1;
    var b1 = w.b1;
    var W2 = w.W2;
    var b2 = w.b2;
    var h = new Array(hid);
    var j, i, s;
    for (j = 0; j < hid; j++) {
      s = b1[j];
      for (i = 0; i < fin; i++) s += x[i] * W1[i * hid + j];
      h[j] = relu(s);
    }
    var logits = new Array(nclass);
    for (j = 0; j < nclass; j++) {
      s = b2[j];
      for (i = 0; i < hid; i++) s += h[i] * W2[i * nclass + j];
      logits[j] = s;
    }
    return softmax(logits);
  }

  /** Mulberry32 PRNG */
  function makeRng(seed) {
    return function () {
      var t = (seed += 0x6d2b79f5);
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  /**
   * Train small MLP (v2, 22-d features) on archetype-based synthetic data.
   * Archetypes mirror AS/NZS 1170.2 Figure 5.2 building shapes:
   *   low_flat, mid_flat, gable, hipped, monopitch, tall.
   * Runs once at page load (~1-2 s on modern CPUs for ~1800 samples × 10 epochs × 48 hidden).
   */
  function trainEmbeddedMlp(opts) {
    opts = opts || {};
    // Phase 6: optional curated corrections upsampled into the training set.
    // Shape: [{ features: number[22], label: 0|1|2|3 }, ...]
    var curatedExtras = Array.isArray(opts.curatedExtras) ? opts.curatedExtras : [];
    var upsample = Math.max(1, Math.min(200, opts.upsample | 0 || 40));
    var noiseSigma = typeof opts.noiseSigma === 'number' ? opts.noiseSigma : 0.02;
    var rng = makeRng(42);
    var gauss = function () {
      var u = 1 - rng();
      var v = 1 - rng();
      return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
    };
    var X = [];
    var y = [];

    // Archetype building bboxes (wide × tall × deep) for AS/NZS 1170.2 Figure 5.2 shapes.
    // Phase 7b: add non-rectangular footprints (L/T/U/courtyard) via a `notch`
    // rectangle that's subtracted from the bbox. Notch edges inside the bbox
    // become inner walls — used by the wall sampler to teach the MLP that
    // walls with centroids off the bbox perimeter are still walls, not interior.
    var ARCHETYPES = {
      low_flat:  { bmin: [-12, 0, -9],  bmax: [12, 4, 9],   roofKind: 'flat' },
      mid_flat:  { bmin: [-14, 0, -10], bmax: [14, 10, 10], roofKind: 'flat' },
      gable:     { bmin: [-10, 0, -6],  bmax: [10, 7, 6],   roofKind: 'gable' },
      hipped:    { bmin: [-9, 0, -7],   bmax: [9, 6.5, 7],  roofKind: 'hipped' },
      monopitch: { bmin: [-8, 0, -5],   bmax: [8, 5.5, 5],  roofKind: 'monopitch' },
      tall:      { bmin: [-7, 0, -7],   bmax: [7, 35, 7],   roofKind: 'flat' },
      l_shape:   { bmin: [-12, 0, -9],  bmax: [12, 5, 9],   roofKind: 'flat',     notch: { xMin: 2,   xMax: 12,  zMin: -9, zMax: -1 } },
      t_shape:   { bmin: [-13, 0, -8],  bmax: [13, 5, 8],   roofKind: 'flat',     notch: { xMin: 6,   xMax: 13,  zMin: -8, zMax: 3  } },
      u_shape:   { bmin: [-11, 0, -9],  bmax: [11, 6, 9],   roofKind: 'flat',     notch: { xMin: -4,  xMax: 4,   zMin: -9, zMax: 2  } },
      courtyard: { bmin: [-14, 0, -10], bmax: [14, 10, 10], roofKind: 'flat',     notch: { xMin: -5,  xMax: 5,   zMin: -4, zMax: 4  } }
    };
    var ARCH_KEYS = ['low_flat','mid_flat','gable','hipped','monopitch','tall','l_shape','t_shape','u_shape','courtyard'];

    // Phase 7b — test whether a mesh centroid falls inside the archetype's
    // notch (empty region). Roof / floor / interior samples there are
    // physically impossible and get re-rolled by emitSample.
    function pointInNotch(mc, notch) {
      if (!notch) return false;
      return mc[0] >= notch.xMin && mc[0] <= notch.xMax &&
             mc[2] >= notch.zMin && mc[2] <= notch.zMax;
    }

    function pickOneHot(lab, rng_) {
      // Map class → plausible IFC category one-hot. Randomise a bit to simulate mixed tagging.
      // Slots: 0=wall 1=curtain 2=opening 3=roof 4=slab 5=interior 6=beam-member 7=proxy
      var oh = [0,0,0,0,0,0,0,0];
      var r = rng_();
      if (lab === 0) { // interior
        if (r < 0.45) oh[5] = 1;
        else if (r < 0.75) oh[6] = 1;
        else oh[7] = 1;
      } else if (lab === 1) { // wall
        if (r < 0.55) oh[0] = 1;
        else if (r < 0.75) oh[1] = 1;
        else if (r < 0.88) oh[2] = 1;
        else oh[7] = 1;
      } else if (lab === 2) { // roof
        if (r < 0.65) oh[3] = 1;
        else if (r < 0.88) oh[4] = 1;
        else oh[7] = 1;
      } else { // floor
        if (r < 0.75) oh[4] = 1;
        else oh[7] = 1;
      }
      return oh;
    }

    function emitSample(lab, arch) {
      var A = ARCHETYPES[arch];
      var bmin = A.bmin, bmax = A.bmax, rk = A.roofKind;
      var extX = bmax[0] - bmin[0], extY = bmax[1] - bmin[1], extZ = bmax[2] - bmin[2];
      var notch = A.notch || null;
      var mc, msz, nv, side;
      if (lab === 1) {
        // Phase 7b: 40% chance to place the wall on the notch perimeter
        // (inner corner of L/T/U/courtyard). Its centroid lands INSIDE the
        // bbox — the whole point of this archetype family.
        var useNotch = !!notch && rng() < 0.4;
        if (useNotch) {
          var nside = (rng() * 4) | 0;
          var zLen = notch.zMax - notch.zMin, xLen = notch.xMax - notch.xMin;
          mc = [0, bmin[1] + rng() * (extY - 0.5) + 0.2, 0];
          if (nside === 0)      { mc[0] = notch.xMin; mc[2] = notch.zMin + rng() * zLen; nv = [1,  gauss() * 0.1, rng() * 0.3 - 0.15]; }
          else if (nside === 1) { mc[0] = notch.xMax; mc[2] = notch.zMin + rng() * zLen; nv = [-1, gauss() * 0.1, rng() * 0.3 - 0.15]; }
          else if (nside === 2) { mc[0] = notch.xMin + rng() * xLen; mc[2] = notch.zMin; nv = [rng() * 0.3 - 0.15, gauss() * 0.1, 1]; }
          else                  { mc[0] = notch.xMin + rng() * xLen; mc[2] = notch.zMax; nv = [rng() * 0.3 - 0.15, gauss() * 0.1, -1]; }
          var whN = rng() * Math.min(extY, 6) + 1;
          msz = [rng() * 3 + 0.8, whN, rng() * 1.1 + 0.2];
        } else {
          // Wall near one of the four vertical faces; height sampled over range.
          mc = [0, bmin[1] + rng() * (extY - 0.5) + 0.2, 0];
          side = (rng() * 4) | 0;
          if (side === 0) { mc[0] = bmax[0] - rng() * 1.2; mc[2] = bmin[2] + rng() * extZ; }
          else if (side === 1) { mc[0] = bmin[0] + rng() * 1.2; mc[2] = bmin[2] + rng() * extZ; }
          else if (side === 2) { mc[0] = bmin[0] + rng() * extX; mc[2] = bmax[2] - rng() * 1.2; }
          else { mc[0] = bmin[0] + rng() * extX; mc[2] = bmin[2] + rng() * 1.2; }
          // Tall wall strips are common on high-rise.
          var wh = arch === 'tall' ? (rng() * 10 + 2) : (rng() * Math.min(extY, 6) + 1);
          msz = [rng() * 3 + 0.8, wh, rng() * 1.1 + 0.2];
          nv = [side < 2 ? (side === 0 ? 1 : -1) : rng() * 0.35 - 0.175,
                gauss() * 0.1,
                side >= 2 ? (side === 2 ? 1 : -1) : rng() * 0.35 - 0.175];
        }
      } else if (lab === 2) {
        // Roof mesh with normal/position depending on roof kind.
        var ryCentre, ny, tilt;
        if (rk === 'flat') {
          ryCentre = bmax[1] - rng() * 0.9;
          ny = 1; tilt = 0.15;
          nv = [gauss() * tilt, ny, gauss() * tilt];
          msz = [rng() * extX * 0.5 + 2, rng() * 0.9 + 0.3, rng() * extZ * 0.5 + 2];
        } else if (rk === 'gable') {
          // Two opposing slopes along one axis.
          ryCentre = bmax[1] - rng() * 2;
          var slopeSide = rng() < 0.5 ? 1 : -1;
          nv = [rng() * 0.3 - 0.15, 0.75 + rng() * 0.15, slopeSide * (0.5 + rng() * 0.3)];
          msz = [rng() * extX * 0.7 + 3, rng() * 1.2 + 0.3, rng() * 5 + 2];
        } else if (rk === 'hipped') {
          // Mix of U/D slopes and R triangular ends.
          ryCentre = bmax[1] - rng() * 2;
          if (rng() < 0.5) {
            // U/D slope along Z
            nv = [rng() * 0.3 - 0.15, 0.72 + rng() * 0.15, (rng() < 0.5 ? 1 : -1) * (0.5 + rng() * 0.3)];
          } else {
            // R crosswind slope — normal dominated by X
            nv = [(rng() < 0.5 ? 1 : -1) * (0.45 + rng() * 0.3), 0.72 + rng() * 0.15, rng() * 0.3 - 0.15];
          }
          msz = [rng() * 4 + 2, rng() * 1.3 + 0.3, rng() * 4 + 2];
        } else if (rk === 'monopitch') {
          ryCentre = bmax[1] - rng() * 1.5;
          var dir = rng() < 0.5 ? 1 : -1;
          nv = [rng() * 0.2 - 0.1, 0.7 + rng() * 0.15, dir * (0.55 + rng() * 0.25)];
          msz = [rng() * extX * 0.8 + 3, rng() * 1 + 0.3, rng() * extZ * 0.8 + 3];
        } else {
          ryCentre = bmax[1] - rng() * 1;
          nv = [gauss() * 0.2, 1, gauss() * 0.2];
          msz = [rng() * 4 + 2, rng() * 1 + 0.3, rng() * 4 + 2];
        }
        mc = [bmin[0] + rng() * extX, ryCentre, bmin[2] + rng() * extZ];
      } else if (lab === 3) {
        // Floor at the bottom of the bbox.
        mc = [bmin[0] + rng() * extX, bmin[1] + rng() * 0.8, bmin[2] + rng() * extZ];
        msz = [rng() * extX * 0.6 + 3, rng() * 0.55 + 0.1, rng() * extZ * 0.6 + 3];
        nv = [rng() * 0.2 - 0.1, -1, rng() * 0.2 - 0.1];
      } else {
        // Interior mesh — typically small, random orientation, somewhere in the middle.
        mc = [bmin[0] + rng() * extX, bmin[1] + rng() * (extY - 1) + 0.5, bmin[2] + rng() * extZ];
        // bias interior slightly away from exterior envelope
        mc[0] = mc[0] * 0.8; mc[2] = mc[2] * 0.8;
        msz = [Math.abs(gauss() * 1) + 0.2, Math.abs(gauss() * 1) + 0.2, Math.abs(gauss() * 1) + 0.2];
        nv = [gauss(), gauss(), gauss()];
      }
      // Phase 7b — reject roof/floor/interior samples that land in the notch
      // (empty space outside the building). Wall samples intentionally sit on
      // the notch perimeter and are kept.
      if (notch && lab !== 1 && pointInNotch(mc, notch)) return;
      var ln = Math.sqrt(nv[0] * nv[0] + nv[1] * nv[1] + nv[2] * nv[2]) + 1e-9;
      nv = [nv[0] / ln, nv[1] / ln, nv[2] / ln];

      var extAx = [Math.max(extX, 1e-6), Math.max(extY, 1e-6), Math.max(extZ, 1e-6)];
      var halfAx = [extAx[0] / 2, extAx[1] / 2, extAx[2] / 2];
      var midAx = [(bmin[0] + bmax[0]) / 2, (bmin[1] + bmax[1]) / 2, (bmin[2] + bmax[2]) / 2];
      var cnormA = [mc[0] - midAx[0], mc[1] - midAx[1], mc[2] - midAx[2]];

      var f0 = Math.max(Math.min(cnormA[0] / halfAx[0], 1.5), -1.5);
      var f1 = Math.max(Math.min(cnormA[1] / halfAx[1], 1.5), -1.5);
      var f2 = Math.max(Math.min(cnormA[2] / halfAx[2], 1.5), -1.5);
      var maxExt = Math.max(extAx[0], extAx[1], extAx[2]);
      var f3 = Math.min(msz[0] / maxExt, 2);
      var f4 = Math.min(msz[1] / maxExt, 2);
      var f5 = Math.min(msz[2] / maxExt, 2);
      var f6 = nv[0], f7 = nv[1], f8 = nv[2];
      var f9 = Math.sqrt(f6 * f6 + f8 * f8);
      var hfrac = (mc[1] - bmin[1]) / (bmax[1] - bmin[1] + 1e-9);
      var f10 = Math.max(0, Math.min(1, hfrac));
      var volR = Math.min((msz[0] * msz[1] * msz[2]) / (extAx[0] * extAx[1] * extAx[2] + 1e-9), 1);
      var f11 = Math.log1p(volR * 10) / 3;

      var oh = pickOneHot(lab, rng);
      var horizMax = Math.max(msz[0], msz[2], 1e-6);
      var horizMin = Math.min(msz[0], msz[2]);
      var f20 = Math.min(msz[1] / horizMax, 4);
      var f21 = horizMax > 1e-6 ? horizMin / horizMax : 1;

      X.push([
        f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11,
        oh[0], oh[1], oh[2], oh[3], oh[4], oh[5], oh[6], oh[7],
        f20, f21
      ]);
      y.push(lab);
    }

    // Class-balanced per archetype: 60 × 4 × 10 = 2400 core samples (Phase 7b
    // bumped the archetype count from 6 → 10 by adding L/T/U/courtyard).
    var N_PER = 60;
    for (var ai = 0; ai < ARCH_KEYS.length; ai++) {
      for (var lab2 = 0; lab2 < 4; lab2++) {
        for (var kk = 0; kk < N_PER; kk++) emitSample(lab2, ARCH_KEYS[ai]);
      }
    }
    // Extra mullion / curtain-wall shards (walls) + sloped-roof facets (roofs) to
    // densify the ambiguous thin-vertical / sloped regions of feature space.
    for (var kk2 = 0; kk2 < 200; kk2++) emitSample(1, 'gable');
    for (var kk3 = 0; kk3 < 150; kk3++) emitSample(2, 'hipped');
    // A few ambiguous IFCSLAB-mid samples labelled interior — sits mid-building,
    // stops the MLP from reflex-labelling every horizontal-normal mesh as roof/floor.
    for (var kk4 = 0; kk4 < 120; kk4++) emitSample(0, 'mid_flat');
    // Phase 7b — densify concave wall samples so the MLP learns the inner-notch
    // pattern robustly. 40% of wall samples inside emitSample() will use the
    // notch-edge path on concave archetypes.
    for (var kk5 = 0; kk5 < 160; kk5++) emitSample(1, 'l_shape');
    for (var kk6 = 0; kk6 < 140; kk6++) emitSample(1, 't_shape');
    for (var kk7 = 0; kk7 < 140; kk7++) emitSample(1, 'u_shape');
    for (var kk8 = 0; kk8 < 120; kk8++) emitSample(1, 'courtyard');
    var lab, k, mc, ext, half, mid, cnorm, msz, nv, ln, maxExt, volR, side, i, j;

    // Phase 6: inject upsampled curated corrections BEFORE shuffle, so they're
    // mixed through the mini-batches rather than stacked at the tail. Each
    // correction is replicated `upsample` times; replicas get small gaussian
    // noise on the geometric features (f0..f11, f20, f21) to act as a
    // lightweight regulariser. The IFC one-hot (f12..f19) is left unchanged
    // so a correction against an IFC short-circuit decision stays meaningful.
    if (curatedExtras.length > 0) {
      for (var ex_i = 0; ex_i < curatedExtras.length; ex_i++) {
        var ex = curatedExtras[ex_i];
        if (!ex || !Array.isArray(ex.features) && !(ex.features && typeof ex.features.length === 'number')) continue;
        if (ex.features.length !== N_FEATURES) continue;
        if (ex.label !== 0 && ex.label !== 1 && ex.label !== 2 && ex.label !== 3) continue;
        for (var ex_u = 0; ex_u < upsample; ex_u++) {
          var noisy = new Array(N_FEATURES);
          for (var ex_f = 0; ex_f < N_FEATURES; ex_f++) {
            var v0 = +ex.features[ex_f];
            if (ex_u > 0 && noiseSigma > 0) {
              var geomIdx = (ex_f < 12) || ex_f === 20 || ex_f === 21;
              if (geomIdx) v0 += gauss() * noiseSigma;
            }
            noisy[ex_f] = v0;
          }
          X.push(noisy);
          y.push(ex.label | 0);
        }
      }
    }

    // Shuffle so mini-batches aren't all-one-class (training is per-archetype-per-label).
    for (var si = X.length - 1; si > 0; si--) {
      var sj = Math.floor(rng() * (si + 1));
      var tx = X[si]; X[si] = X[sj]; X[sj] = tx;
      var ty = y[si]; y[si] = y[sj]; y[sj] = ty;
    }
    var n = X.length;
    var s1 = Math.sqrt(2 / (N_FEATURES + N_HIDDEN));
    var s2 = Math.sqrt(2 / (N_HIDDEN + N_CLASSES));
    var W1 = new Array(N_FEATURES * N_HIDDEN);
    var b1 = new Array(N_HIDDEN);
    var W2 = new Array(N_HIDDEN * N_CLASSES);
    var b2 = new Array(N_CLASSES);
    for (i = 0; i < W1.length; i++) W1[i] = gauss() * s1;
    for (i = 0; i < N_HIDDEN; i++) b1[i] = 0;
    for (i = 0; i < W2.length; i++) W2[i] = gauss() * s2;
    for (i = 0; i < N_CLASSES; i++) b2[i] = 0;

    var lr = 0.08;
    var epochs = 10;
    var bs = 64;
    var ep, start, bi, x, yy, p, h, preH, dlog, dh, j2, k2, gW1, gb1, gW2, gb2, nb, t;

    for (ep = 0; ep < epochs; ep++) {
      for (start = 0; start < n; start += bs) {
        gW1 = new Array(W1.length);
        gb1 = new Array(N_HIDDEN);
        gW2 = new Array(W2.length);
        gb2 = new Array(N_CLASSES);
        for (i = 0; i < gW1.length; i++) gW1[i] = 0;
        for (i = 0; i < N_HIDDEN; i++) gb1[i] = 0;
        for (i = 0; i < gW2.length; i++) gW2[i] = 0;
        for (i = 0; i < N_CLASSES; i++) gb2[i] = 0;

        nb = Math.min(bs, n - start);
        for (bi = 0; bi < nb; bi++) {
          x = X[start + bi];
          yy = y[start + bi];
          preH = new Array(N_HIDDEN);
          for (j = 0; j < N_HIDDEN; j++) {
            t = b1[j];
            for (i = 0; i < N_FEATURES; i++) t += x[i] * W1[i * N_HIDDEN + j];
            preH[j] = t;
          }
          h = preH.map(function (v) { return relu(v); });
          var logits = new Array(N_CLASSES);
          for (k2 = 0; k2 < N_CLASSES; k2++) {
            t = b2[k2];
            for (j = 0; j < N_HIDDEN; j++) t += h[j] * W2[j * N_CLASSES + k2];
            logits[k2] = t;
          }
          p = softmax(logits);
          dlog = new Array(N_CLASSES);
          for (k2 = 0; k2 < N_CLASSES; k2++) dlog[k2] = p[k2] - (k2 === yy ? 1 : 0);
          for (k2 = 0; k2 < N_CLASSES; k2++) gb2[k2] += dlog[k2];
          for (j = 0; j < N_HIDDEN; j++) {
            for (k2 = 0; k2 < N_CLASSES; k2++) gW2[j * N_CLASSES + k2] += dlog[k2] * h[j];
          }
          dh = new Array(N_HIDDEN);
          for (j = 0; j < N_HIDDEN; j++) {
            t = 0;
            for (k2 = 0; k2 < N_CLASSES; k2++) t += dlog[k2] * W2[j * N_CLASSES + k2];
            dh[j] = t * (preH[j] > 0 ? 1 : 0);
          }
          for (j = 0; j < N_HIDDEN; j++) gb1[j] += dh[j];
          for (i = 0; i < N_FEATURES; i++) {
            for (j = 0; j < N_HIDDEN; j++) gW1[i * N_HIDDEN + j] += dh[j] * x[i];
          }
        }
        for (i = 0; i < W1.length; i++) W1[i] -= (lr * gW1[i]) / nb;
        for (j = 0; j < N_HIDDEN; j++) b1[j] -= (lr * gb1[j]) / nb;
        for (i = 0; i < W2.length; i++) W2[i] -= (lr * gW2[i]) / nb;
        for (k2 = 0; k2 < N_CLASSES; k2++) b2[k2] -= (lr * gb2[k2]) / nb;
      }
      lr *= 0.92;
    }

    return {
      version: 2,
      inputDim: N_FEATURES,
      hiddenDim: N_HIDDEN,
      numClasses: N_CLASSES,
      labels: ['interior', 'exterior_wall', 'roof', 'floor_ground'],
      W1: W1,
      W1Shape: [N_FEATURES, N_HIDDEN],
      b1: b1,
      W2: W2,
      W2Shape: [N_HIDDEN, N_CLASSES],
      b2: b2,
      trainAccuracy: 0
    };
  }

  function inferProbs(features) {
    if (!state.mlp) throw new Error('MeshClassifier: MLP not initialized');
    return mlpForward(features, state.mlp);
  }

  /**
   * IFC entity → semantic class + 8-dim category one-hot (AS/NZS 1170.2 Figure 5.2 envelope).
   * Returns { cls, conf, wallHint, roofHint, oneHot }:
   *   cls:      0|1|2|3 when the IFC type is definitive enough to short-circuit the MLP; null otherwise
   *   conf:     [0..1] short-circuit confidence (only meaningful when cls !== null)
   *   wallHint: legacy f12 scalar in [0..1] (kept for v1-weights compat)
   *   roofHint: legacy f13 scalar in [0..1]
   *   oneHot:   8-dim array used as f12..f19 in the v2 feature vector.
   *             Slots: [wall, curtain, opening, roof, slab, interior, beam-member, proxy-covering]
   *             Each is 1.0 for an exact match, 0 otherwise. Empty array when no IFC tag.
   */
  function ifcClassify(mesh, geomCtx) {
    var out = { cls: null, conf: 0, wallHint: 0, roofHint: 0, oneHot: [0,0,0,0,0,0,0,0] };
    var t = mesh && mesh.userData && mesh.userData.ifcLineType;
    if (typeof WebIFC === 'undefined' || t == null || t === undefined) return out;
    var W = WebIFC;
    var nm = function (k) { return typeof W[k] === 'number' ? W[k] : -1; };
    var oh = [0,0,0,0,0,0,0,0]; // wall, curtain, opening, roof, slab, interior, beam-member, proxy-cover
    try {
      // --- Definitive walls -----------------------------------------------
      if (t === nm('IFCWALL') || t === nm('IFCWALLSTANDARDCASE') || t === nm('IFCWALLTYPE')) {
        oh[0] = 1;
        return { cls: 1, conf: 0.95, wallHint: 1, roofHint: 0, oneHot: oh };
      }
      if (t === nm('IFCCURTAINWALL')) {
        oh[1] = 1;
        return { cls: 1, conf: 0.90, wallHint: 0.95, roofHint: 0, oneHot: oh };
      }
      if (t === nm('IFCWINDOW') || t === nm('IFCDOOR')) {
        oh[2] = 1;
        return { cls: 1, conf: 0.85, wallHint: 0.85, roofHint: 0, oneHot: oh };
      }
      // --- Definitive roofs -----------------------------------------------
      if (t === nm('IFCROOF') || t === nm('IFCROOFTYPE')) {
        oh[3] = 1;
        return { cls: 2, conf: 0.95, wallHint: 0, roofHint: 1, oneHot: oh };
      }
      // --- Definitive floors / footings -----------------------------------
      if (t === nm('IFCFOOTING')) {
        oh[7] = 1;
        return { cls: 3, conf: 0.85, wallHint: 0, roofHint: 0.2, oneHot: oh };
      }
      // --- Definitive interior (non-envelope) -----------------------------
      if (t === nm('IFCSTAIR') || t === nm('IFCSTAIRFLIGHT')
          || t === nm('IFCRAILING') || t === nm('IFCRAMP') || t === nm('IFCRAMPFLIGHT')) {
        oh[5] = 1;
        return { cls: 0, conf: 0.90, wallHint: 0, roofHint: 0, oneHot: oh };
      }
      if (t === nm('IFCSPACE')) {
        oh[5] = 1;
        return { cls: 0, conf: 0.95, wallHint: 0, roofHint: 0, oneHot: oh };
      }
      if (t === nm('IFCFURNISHINGELEMENT') || t === nm('IFCFURNITURE')
          || t === nm('IFCSYSTEMFURNITUREELEMENT')) {
        oh[5] = 1;
        return { cls: 0, conf: 0.85, wallHint: 0, roofHint: 0, oneHot: oh };
      }
      // --- IFCSLAB — disambiguate by vertical position in building bbox ---
      if (t === nm('IFCSLAB')) {
        oh[4] = 1;
        if (geomCtx && geomCtx.hfrac != null) {
          if (geomCtx.hfrac >= 0.82) return { cls: 2, conf: 0.80, wallHint: 0, roofHint: 0.95, oneHot: oh };
          if (geomCtx.hfrac <= 0.12) return { cls: 3, conf: 0.80, wallHint: 0, roofHint: 0.20, oneHot: oh };
        }
        return { cls: null, conf: 0, wallHint: 0.10, roofHint: 0.55, oneHot: oh };
      }
      // --- Weak signals (feed MLP, no short-circuit) ----------------------
      if (t === nm('IFCPLATE')) { oh[6] = 1; out.wallHint = 0.60; out.oneHot = oh; return out; }
      if (t === nm('IFCMEMBER')) { oh[6] = 1; out.wallHint = 0.30; out.oneHot = oh; return out; }
      if (t === nm('IFCBEAM')) { oh[6] = 1; out.wallHint = 0.20; out.oneHot = oh; return out; }
      if (t === nm('IFCCOLUMN')) { oh[6] = 1; out.wallHint = 0.20; out.oneHot = oh; return out; }
      if (t === nm('IFCCOVERING')) { oh[7] = 1; out.wallHint = 0.35; out.oneHot = oh; return out; }
      if (t === nm('IFCBUILDINGELEMENTPROXY')) { oh[7] = 1; out.wallHint = 0.20; out.oneHot = oh; return out; }
    } catch (e) { /* ignore */ }
    return out;
  }

  /** Legacy shim — returns [wallHint, roofHint] (v1 f12/f13 scalars). */
  function ifcHints(mesh, geomCtx) {
    var c = ifcClassify(mesh, geomCtx);
    return [c.wallHint, c.roofHint];
  }

  function extractFeatures(mesh, bboxBuilding, bCenter, bSize, mc, meshBox, nAccPre) {
    var bmin = bboxBuilding.min;
    var bmax = bboxBuilding.max;
    var ext = new THREE.Vector3(
      Math.max(bmax.x - bmin.x, 1e-6),
      Math.max(bmax.y - bmin.y, 1e-6),
      Math.max(bmax.z - bmin.z, 1e-6)
    );
    var half = ext.clone().multiplyScalar(0.5);
    var mid = new THREE.Vector3().addVectors(bmin, bmax).multiplyScalar(0.5);
    var cnorm = new THREE.Vector3().subVectors(mc, mid);
    var f0 = Math.max(Math.min(cnorm.x / half.x, 1.5), -1.5);
    var f1 = Math.max(Math.min(cnorm.y / half.y, 1.5), -1.5);
    var f2 = Math.max(Math.min(cnorm.z / half.z, 1.5), -1.5);

    var msz = new THREE.Vector3();
    meshBox.getSize(msz);
    var maxExt = Math.max(ext.x, ext.y, ext.z);
    var f3 = Math.min(msz.x / maxExt, 2);
    var f4 = Math.min(msz.y / maxExt, 2);
    var f5 = Math.min(msz.z / maxExt, 2);

    var nAcc = new THREE.Vector3();
    var normAttr = mesh.geometry.attributes.normal;
    var normalMatrix = new THREE.Matrix3();
    var nWorld = new THREE.Vector3();
    if (nAccPre && nAccPre.lengthSq() > 1e-12) {
      nAcc.copy(nAccPre);
    } else if (normAttr && normAttr.count) {
      normalMatrix.getNormalMatrix(mesh.matrixWorld);
      for (var ii = 0; ii < normAttr.count; ii++) {
        nWorld.fromBufferAttribute(normAttr, ii);
        nWorld.applyMatrix3(normalMatrix).normalize();
        nAcc.add(nWorld);
      }
      nAcc.divideScalar(normAttr.count);
      var alen = nAcc.length();
      if (alen > 1e-6) nAcc.divideScalar(alen);
    }
    var f6 = nAcc.x;
    var f7 = nAcc.y;
    var f8 = nAcc.z;
    var horiz = Math.sqrt(f6 * f6 + f8 * f8);
    var f9 = horiz;
    var hfrac = (mc.y - bmin.y) / (bmax.y - bmin.y + 1e-9);
    var f10 = Math.max(0, Math.min(1, hfrac));
    var volR = Math.min((msz.x * msz.y * msz.z) / (ext.x * ext.y * ext.z + 1e-9), 1);
    var f11 = Math.log1p(volR * 10) / 3;

    // v2 features: 8-dim IFC one-hot (f12..f19) + 2 aspect features (f20..f21).
    var ifc = ifcClassify(mesh, { hfrac: f10 });
    var oh = ifc.oneHot || [0,0,0,0,0,0,0,0];
    var f12 = oh[0], f13 = oh[1], f14 = oh[2], f15 = oh[3];
    var f16 = oh[4], f17 = oh[5], f18 = oh[6], f19 = oh[7];

    // Vertical aspect — tall-thin vs short-squat. Walls: high; floors/roofs: low.
    var horizMax = Math.max(msz.x, msz.z, 1e-6);
    var f20 = Math.min(msz.y / horizMax, 4);
    // Plan aspect — square (1) vs ribbon (→0). Roofs tend to be squarer; walls elongated.
    var horizMin = Math.min(msz.x, msz.z);
    var f21 = horizMax > 1e-6 ? horizMin / horizMax : 1;

    return new Float32Array([
      f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11,
      f12, f13, f14, f15, f16, f17, f18, f19,
      f20, f21
    ]);
  }

  /**
   * High-level classification: runs the IFC short-circuit first, falls back to the MLP.
   * Returns { probs:[p_int,p_wall,p_roof,p_floor], viaIfc, ifcClass, ifcConf, features }.
   *
   * The short-circuit blends the confident one-hot with a uniform prior at (1-conf) so
   * downstream ML_CONF thresholds still behave sensibly (conf 0.95 → p_winner ≈ 0.96).
   */
  function classifyMesh(mesh, bboxBuilding, bCenter, bSize, mc, meshBox, nAccPre) {
    var bmin = bboxBuilding.min;
    var bmax = bboxBuilding.max;
    var extY = Math.max(bmax.y - bmin.y, 1e-6);
    var hfrac = Math.max(0, Math.min(1, (mc.y - bmin.y) / extY));
    var ifc = ifcClassify(mesh, { hfrac: hfrac });

    // Always extract features so downstream active-learning can cache them
    // even when the IFC short-circuit wins (user may override the face-key and
    // we want the feature vector available for JSONL export).
    var feats = extractFeatures(mesh, bboxBuilding, bCenter, bSize, mc, meshBox, nAccPre);

    if (ifc.cls !== null && ifc.conf >= 0.80) {
      var p = [0.01, 0.01, 0.01, 0.01];
      p[ifc.cls] = 1 - 3 * 0.01;
      var conf = ifc.conf;
      // Blend with uniform prior so ML_CONF thresholds still compare meaningfully.
      var blended = [0, 0, 0, 0];
      for (var i = 0; i < 4; i++) blended[i] = conf * p[i] + (1 - conf) * 0.25;
      return {
        probs: blended,
        viaIfc: true,
        ifcClass: ifc.cls,
        ifcConf: ifc.conf,
        features: feats
      };
    }

    var probs = inferProbs(feats);
    return {
      probs: probs,
      viaIfc: false,
      ifcClass: null,
      ifcConf: 0,
      features: feats
    };
  }

  /**
   * Phase 3 — hull-aware post-processing.
   *
   * Refines a classifyMesh() result using geometric evidence from the building
   * bbox and (optional) footprint hull. Fires only when the IFC short-circuit
   * did NOT win AND the MLP prediction is below a confidence floor — IFC tags
   * always take precedence.
   *
   * Two conservative rules:
   *   (1) shell_promote — centroid within SHELL_NEAR of a bbox plan face (or
   *       within the hull band) AND outward normal points away from centre,
   *       but MLP labelled "interior": promote to wall/roof/floor based on
   *       normal direction + height gaps.
   *   (2) interior_demote — centroid deeper than SHELL_FAR from all plan
   *       faces AND MLP labelled wall/roof: demote to interior.
   *
   * Returns the same shape as classifyMesh() plus:
   *   postPass: { shellScore, planDist, geomCls, override }
   *
   * @param {object} classRes  — result from classifyMesh().
   * @param {THREE.Box3} bboxBuilding — full-building bbox.
   * @param {THREE.Vector3} bCenter — bbox centre.
   * @param {THREE.Vector3} mc — mesh centroid (world).
   * @param {THREE.Box3} meshBox — mesh bbox.
   * @param {THREE.Vector3} nAccPre — raw accumulated mean normal.
   * @param {number} hullEdgeDist — distance (world units) from mc XZ to footprint
   *     hull edge; pass Infinity when no hull context is available.
   */
  function applyHullRefinement(classRes, bboxBuilding, bCenter, mc, meshBox, nAccPre, hullEdgeDist) {
    var CONF_FLOOR = 0.72;   // MLP predictions below this are considered soft
    var SHELL_NEAR = 0.08;   // plan fraction — centroid within 8% of a face = on shell
    var SHELL_FAR  = 0.18;   // plan fraction — centroid deeper than 18% = interior
    var VERT_DOT   = 0.72;
    var HORIZ_DOT  = 0.55;

    var probs = classRes.probs.slice();
    var bestIdx = 0, bestP = probs[0];
    for (var i = 1; i < 4; i++) {
      if (probs[i] > bestP) { bestP = probs[i]; bestIdx = i; }
    }

    var out = {
      probs: probs,
      viaIfc: classRes.viaIfc,
      ifcClass: classRes.ifcClass,
      ifcConf: classRes.ifcConf,
      features: classRes.features,
      postPass: { shellScore: 0, planDist: 1, geomCls: null, override: null }
    };

    // IFC trumps geometry: never override a confident IFC short-circuit.
    if (classRes.viaIfc) return out;

    var bmin = bboxBuilding.min, bmax = bboxBuilding.max;
    var extX = Math.max(bmax.x - bmin.x, 1e-6);
    var extY = Math.max(bmax.y - bmin.y, 1e-6);
    var extZ = Math.max(bmax.z - bmin.z, 1e-6);
    var ux = (mc.x - bmin.x) / extX;
    var uz = (mc.z - bmin.z) / extZ;
    var planDist = Math.min(ux, 1 - ux, uz, 1 - uz);

    // Tighten plan distance using the footprint hull when provided.
    if (isFinite(hullEdgeDist) && hullEdgeDist >= 0) {
      var hullU = hullEdgeDist / Math.max(extX, extZ);
      if (hullU < planDist) planDist = hullU;
    }

    var roofGap = (bmax.y - (meshBox && meshBox.max ? meshBox.max.y : mc.y)) / extY;
    var floorGap = ((meshBox && meshBox.min ? meshBox.min.y : mc.y) - bmin.y) / extY;

    var shellScore;
    if (planDist <= SHELL_NEAR) shellScore = 1 - planDist / SHELL_NEAR;
    else if (planDist >= SHELL_FAR) shellScore = 0;
    else shellScore = 0.5 * (SHELL_FAR - planDist) / (SHELL_FAR - SHELL_NEAR);
    if (roofGap < 0.05) shellScore = Math.max(shellScore, 0.9);
    if (floorGap < 0.05) shellScore = Math.max(shellScore, 0.7);

    var n = outwardAlignedNormal(nAccPre, mc, bCenter);
    var ny = n.y, nx = n.x, nz = n.z;
    var absNy = Math.abs(ny);
    var horizLen = Math.hypot(nx, nz);

    // Horizontal outward dot: does the plan-normal actually point away from centre?
    var vx = mc.x - bCenter.x, vz = mc.z - bCenter.z;
    var vlen = Math.hypot(vx, vz) + 1e-6;
    var outDot = (nx * vx + nz * vz) / vlen;

    var geomCls = null;
    if (shellScore >= 0.75) {
      if (absNy >= VERT_DOT && ny > 0.3 && roofGap < 0.12) geomCls = 2;        // roof
      else if (absNy >= VERT_DOT && ny < -0.3 && floorGap < 0.12) geomCls = 3; // floor
      else if (horizLen >= HORIZ_DOT && outDot > 0) geomCls = 1;               // wall
    } else if (shellScore <= 0.12) {
      geomCls = 0;                                                             // interior
    }

    out.postPass.shellScore = shellScore;
    out.postPass.planDist = planDist;
    out.postPass.geomCls = geomCls;

    // Conservative overrides — only when the MLP is not confident.
    if (geomCls !== null && geomCls !== bestIdx && bestP < CONF_FLOOR) {
      if (geomCls !== 0 && bestIdx === 0 && shellScore >= 0.80) {
        // Shell-promote: interior → wall/roof/floor.
        for (var a = 0; a < 4; a++) probs[a] = 0.06;
        probs[geomCls] = 0.82;
        out.postPass.override = 'shell_promote';
      } else if (geomCls === 0 && (bestIdx === 1 || bestIdx === 2) && shellScore <= 0.12) {
        // Interior-demote: wall/roof → interior.
        for (var b = 0; b < 4; b++) probs[b] = 0.06;
        probs[0] = 0.82;
        out.postPass.override = 'interior_demote';
      }
    }

    out.probs = probs;
    return out;
  }

  /**
   * IFC meshes often have flipped / degenerate normals; align mean normal to point
   * outward from building bbox. If the input normal is near-zero (degenerate mesh
   * with cancelling face normals), synthesize one from centroid-to-mesh direction.
   */
  function outwardAlignedNormal(nAcc, mc, bCenter) {
    var n = nAcc.clone();
    var nLen = n.length();
    if (nLen < 1e-4) {
      // Degenerate: derive outward direction from geometry alone.
      var dx = mc.x - bCenter.x;
      var dz = mc.z - bCenter.z;
      var dy = mc.y - bCenter.y;
      var hl = Math.hypot(dx, dz);
      if (hl < 1e-6 && Math.abs(dy) < 1e-6) return new THREE.Vector3(0, 1, 0);
      // Strongly above/below centre → treat as horizontal (roof/floor).
      if (hl < 1e-6 || Math.abs(dy) > hl) {
        return new THREE.Vector3(0, dy >= 0 ? 1 : -1, 0);
      }
      return new THREE.Vector3(dx / hl, 0, dz / hl);
    }
    if (Math.abs(nLen - 1) > 1e-3) n.divideScalar(nLen);

    var outH = new THREE.Vector3(mc.x - bCenter.x, 0, mc.z - bCenter.z);
    var ay = Math.abs(n.y);
    if (ay < 0.85 && outH.lengthSq() > 1e-8) {
      outH.normalize();
      var nh = new THREE.Vector3(n.x, 0, n.z);
      if (nh.lengthSq() > 1e-8) {
        nh.normalize();
        if (nh.dot(outH) < 0) n.negate();
      }
    } else if (ay >= 0.5) {
      var out3 = new THREE.Vector3().subVectors(mc, bCenter);
      if (out3.lengthSq() > 1e-8) {
        out3.normalize();
        if (n.dot(out3) < 0) n.negate();
      }
    }
    return n;
  }

  /**
   * Pitch angle (degrees) of a surface normal relative to the vertical axis.
   * 0°  = perfectly horizontal roof face (normal points straight up/down)
   * 90° = vertical wall
   * Used to classify flat-vs-sloped roofs per AS/NZS 1170.2 Cl. 5.3.
   */
  function roofPitchFromNormal(n) {
    if (!n) return 0;
    var ay = Math.abs(n.y);
    var horizLen = Math.hypot(n.x, n.z);
    return Math.atan2(horizLen, Math.max(ay, 1e-6)) * (180 / Math.PI);
  }

  /**
   * Analytic wind keys from mean outward normal + wind (no bbox shell rules).
   * Phase 7c: roof face-key assignment uses the explicit pitch angle
   * (deg from horizontal) instead of a y-dot cutoff. Defaults:
   *   pitchDeg  < flatPitchDeg (5°)   → flat: bin by centroid
   *   pitchDeg ≥ flatPitchDeg         → sloped: bin by slope-facing direction
   * Phase 8a: wall face-keys use an axis-comparison (|dotW| vs |dotP|)
   * instead of the |dotW|≤dotWall band test. This guarantees any two walls
   * 90° apart in the horizontal plane receive different face-keys.
   * Phase 8c: when ctx.confidenceOut is provided (object), the function
   * also writes ctx.confidenceOut.lastConfidence ∈ [0, 1]:
   *   wall:        |aW − aP|  — distance from the 45° tie line
   *   roof flat:   (flatPitchDeg − pitchDeg) / flatPitchDeg
   *   roof ww/lw:  (|dw| − dotWall) / (1 − dotWall)
   *   roof cw:     (dotWall − |dw|) / dotWall
   * 0 ⇒ on a decision boundary (caller may want to defer to overrides),
   * 1 ⇒ unambiguous. Confidence is set to 0 when the function returns null.
   * Caller can override ctx.flatPitchDeg and ctx.dotWall (roof branch only).
   */
  function semanticToFaceKey(sem, nAcc, mc, ctx) {
    var DOT_W = 0.22;
    var FLAT_PITCH_DEG = 5;
    var windDir = ctx.windDir;
    var perpDir = ctx.perpDir;
    var bCenter = ctx.bCenter;
    var nH = ctx.nH;
    var confidenceOut = ctx && ctx.confidenceOut && typeof ctx.confidenceOut === 'object'
      ? ctx.confidenceOut : null;
    function setConf(c) {
      if (!confidenceOut) return;
      // Clamp to [0,1] defensively; numerical noise can land marginally outside.
      var v = c < 0 ? 0 : (c > 1 ? 1 : c);
      confidenceOut.lastConfidence = v;
    }
    /** Optional override of DOT_W (default 0.22). Used by the ROOF branch only:
     *  controls the |dw|≤band → roof_cw cutoff vs roof_ww/roof_lw. The wall
     *  branch ignores this since Phase 8a — see axis-comparison below. */
    var dotWall = ctx && typeof ctx.dotWall === 'number' ? ctx.dotWall : DOT_W;
    /** Optional override of flat-roof pitch threshold in degrees (default 5°). */
    var flatPitchDeg = ctx && typeof ctx.flatPitchDeg === 'number' ? ctx.flatPitchDeg : FLAT_PITCH_DEG;

    if (sem === SEM.interior || sem === SEM.floor) { setConf(0); return null; }

    var n = outwardAlignedNormal(nAcc, mc, bCenter);
    var horizLen = Math.hypot(n.x, n.z);

    if (sem === SEM.wall) {
      if (horizLen < 1e-4) { setConf(0); return null; }
      nH.set(n.x, 0, n.z);
      if (nH.lengthSq() < 1e-8) { setConf(0); return null; }
      nH.normalize();
      var dotW = nH.dot(windDir);
      var dotP = nH.dot(perpDir);
      // Phase 8a: axis-comparison instead of |dotW|≤band test.
      // Guarantees that any two walls 90° apart in the horizontal plane
      // receive different face-keys (they cannot both be 'windward', etc.).
      // The previous band test allowed e.g. walls at 135° and 225° to both
      // map to 'windward', which is geometrically impossible per AS/NZS Fig 5.2.
      // Tie-break at the exact 45° case (|dotW|==|dotP|): same-sign → wind axis,
      // opposite-sign → cross axis. This splits the diamond-rotated orthogonal
      // pair into one wind-axis wall and one cross-axis wall.
      // ctx.dotWall is retained for the roof branch only; walls now use the
      // axis-comparison unconditionally.
      var aW = Math.abs(dotW);
      var aP = Math.abs(dotP);
      var windAxis;
      if (aW > aP)      windAxis = true;
      else if (aW < aP) windAxis = false;
      else              windAxis = (dotW * dotP) >= 0;
      // Phase 8c: confidence = axis-margin. Pure axis (aW=1, aP=0) → 1.
      // 45° diamond tie (aW=aP=√½) → 0.
      setConf(Math.abs(aW - aP));
      if (windAxis) return dotW < 0 ? 'windward' : 'leeward';
      return dotP > 0 ? 'sidewall2' : 'sidewall1';
    }

    if (sem === SEM.roof) {
      var pitchDeg = roofPitchFromNormal(n);
      // Expose the pitch so downstream callers (pressure code, eval harness)
      // can consult it without recomputing. Key by face-group if ctx provides
      // a bucket; otherwise stash the most recent pitch on ctx.
      if (ctx && ctx.roofPitchOut && typeof ctx.roofPitchOut === 'object') {
        ctx.roofPitchOut.lastPitchDeg = pitchDeg;
      }

      // Near-flat roof — bin by centroid position relative to wind.
      if (pitchDeg < flatPitchDeg && n.y > 0.35) {
        var dx = mc.x - bCenter.x;
        var dz = mc.z - bCenter.z;
        var windProj = dx * windDir.x + dz * windDir.z;
        // Phase 8c: confidence = pitch margin from the flat threshold.
        // pitch=0 → 1.0; pitch→threshold → 0.
        setConf(flatPitchDeg > 0 ? (flatPitchDeg - pitchDeg) / flatPitchDeg : 1);
        return windProj > 0 ? 'roof_lw' : 'roof_ww';
      }

      if (horizLen > 1e-4) {
        // Sloped roof with a meaningful horizontal normal component.
        // AS/NZS 1170.2 Figure 5.2 faces: U (upwind), D (downwind), R (crosswind).
        //   dw = horizontal-component-unit ⋅ windDir:
        //     dw < -dotWall → slope faces into the wind     → U (roof_ww)
        //     dw >  dotWall → slope faces away from the wind → D (roof_lw)
        //     |dw| ≤ dotWall → slope faces crosswind         → R (roof_cw)
        // Shallow slopes (flatPitchDeg ≤ pitch < ~10°) are now correctly
        // routed here instead of falling into the centroid-bin branch above,
        // which was the Phase 7c fix.
        nH.set(n.x, 0, n.z).normalize();
        var dw = nH.dot(windDir);
        var adw = Math.abs(dw);
        // Phase 8c: piecewise confidence on distance to the dotWall boundary.
        if (adw > dotWall) {
          // ww or lw: 1 at adw=1 (slope perfectly faces wind), 0 at boundary.
          setConf(1 - dotWall > 0 ? (adw - dotWall) / (1 - dotWall) : 1);
        } else {
          // cw: 1 at adw=0 (perfectly cross-wind), 0 at boundary.
          setConf(dotWall > 0 ? (dotWall - adw) / dotWall : 1);
        }
        if (dw < -dotWall) return 'roof_ww';
        if (dw >  dotWall) return 'roof_lw';
        return 'roof_cw';
      }
      // Degenerate horizLen ≈ 0 and pitch above flat threshold: only happens
      // for sideways-facing roof segments that point straight down (n.y < 0),
      // which is non-physical for outward roof normals. Bin by centroid.
      var dx2 = mc.x - bCenter.x;
      var dz2 = mc.z - bCenter.z;
      var wp = dx2 * windDir.x + dz2 * windDir.z;
      // Degenerate path — confidence is low since the horizontal normal vanished.
      setConf(0);
      return wp > 0 ? 'roof_lw' : 'roof_ww';
    }

    setConf(0);
    return null;
  }

  /**
   * Accept weights JSON only when it matches the current feature/class layout.
   * v1 (14-d) files are rejected — a stale file on disk is ignored and we keep
   * the v2 weights trained in-browser at page load.
   */
  function loadWeightsJson(obj) {
    if (!obj || !obj.W1 || !obj.W2) return false;
    var ver = obj.version || 1;
    if (ver !== 2) return false;
    if (obj.inputDim !== N_FEATURES) {
      try { console.warn('MeshClassifier: weights inputDim=' + obj.inputDim + ' != ' + N_FEATURES + ', keeping embedded.'); } catch (e) {}
      return false;
    }
    if (obj.numClasses !== N_CLASSES) return false;
    state.mlp = {
      inputDim: obj.inputDim,
      hiddenDim: obj.hiddenDim,
      numClasses: obj.numClasses,
      W1: obj.W1,
      b1: obj.b1,
      W2: obj.W2,
      b2: obj.b2
    };
    return true;
  }

  function initMlp() {
    var w = trainEmbeddedMlp();
    loadWeightsJson(w);
    console.log('MeshClassifier: embedded MLP trained in-browser (synthetic data).');
  }

  // Phase 7a — persist retrained weights across reloads so user corrections stick.
  var PERSIST_KEY = 'meshClassifierWeights.v2';

  function saveWeightsToStorage() {
    try {
      if (typeof localStorage === 'undefined' || !state.mlp) return false;
      var m = state.mlp;
      var payload = {
        version: 2,
        inputDim: m.inputDim,
        hiddenDim: m.hiddenDim,
        numClasses: m.numClasses,
        W1: Array.prototype.slice.call(m.W1),
        b1: Array.prototype.slice.call(m.b1),
        W2: Array.prototype.slice.call(m.W2),
        b2: Array.prototype.slice.call(m.b2),
        savedAt: Date.now()
      };
      localStorage.setItem(PERSIST_KEY, JSON.stringify(payload));
      return true;
    } catch (e) {
      try { console.warn('MeshClassifier.saveWeightsToStorage failed:', e); } catch (_) {}
      return false;
    }
  }

  function loadWeightsFromStorage() {
    try {
      if (typeof localStorage === 'undefined') return false;
      var raw = localStorage.getItem(PERSIST_KEY);
      if (!raw) return false;
      var obj = JSON.parse(raw);
      return loadWeightsJson(obj);
    } catch (e) {
      return false;
    }
  }

  function clearPersistedWeights() {
    try {
      if (typeof localStorage !== 'undefined') localStorage.removeItem(PERSIST_KEY);
      return true;
    } catch (e) { return false; }
  }

  /**
   * Phase 6 — online retraining with captured corrections.
   *
   * Retrains the embedded MLP with a fresh synthetic archetype dataset PLUS
   * the supplied corrections (each upsampled with light gaussian noise for
   * regularisation). Corrections never replace archetypes — they augment.
   *
   * @param {Array<{features:number[22], label:0|1|2|3}>} rows — correction rows
   * @param {object} [opts] — forwarded to trainEmbeddedMlp; supports upsample,
   *     noiseSigma. Defaults: upsample=40, noiseSigma=0.02.
   * @returns {{ok:boolean, samples:number, weights:object|null}}
   */
  function retrainWithCorrections(rows, opts) {
    var valid = [];
    if (Array.isArray(rows)) {
      for (var r = 0; r < rows.length; r++) {
        var e = rows[r];
        if (!e || !e.features) continue;
        if (e.features.length !== N_FEATURES) continue;
        if (e.label !== 0 && e.label !== 1 && e.label !== 2 && e.label !== 3) continue;
        valid.push({ features: e.features, label: e.label | 0 });
      }
    }
    if (!valid.length) {
      console.warn('MeshClassifier.retrainWithCorrections: no valid rows');
      return { ok: false, samples: 0, weights: null };
    }
    var trainOpts = {
      curatedExtras: valid,
      upsample: opts && opts.upsample ? opts.upsample : 40,
      noiseSigma: opts && typeof opts.noiseSigma === 'number' ? opts.noiseSigma : 0.02
    };
    var w = trainEmbeddedMlp(trainOpts);
    var loaded = loadWeightsJson(w);
    if (!loaded) {
      console.warn('MeshClassifier.retrainWithCorrections: loadWeightsJson rejected fresh weights');
      return { ok: false, samples: valid.length, weights: null };
    }
    // Phase 7a: persist the freshly-retrained weights so the correction
    // survives a page reload. Best-effort — quota / privacy mode just warns.
    var persisted = saveWeightsToStorage();
    console.log('MeshClassifier: retrained with ' + valid.length + ' correction(s), upsample=' + trainOpts.upsample + (persisted ? ' [persisted]' : ''));
    return { ok: true, samples: valid.length, weights: w, persisted: persisted };
  }

  /** Train once at load so MLP is always available without network. Optional file overrides below. */
  initMlp();
  state.ready = true;

  // Phase 7a — persisted user-retrained weights take priority over the shipped
  // JSON baseline so corrections stick across reloads.
  var hasPersistedWeights = loadWeightsFromStorage();
  if (hasPersistedWeights) {
    try { console.log('MeshClassifier: loaded persisted retrained weights from localStorage.'); } catch (e) {}
  }

  fetch('mesh-classifier-weights.json', { cache: 'no-store' })
    .then(function (r) { return r.ok ? r.json() : Promise.reject(); })
    .then(function (j) {
      if (hasPersistedWeights) return; // user's retrained weights win over shipped baseline
      if (loadWeightsJson(j)) console.log('MeshClassifier: mesh-classifier-weights.json replaced embedded weights.');
    })
    .catch(function () { /* keep embedded */ });

  global.MeshClassifier = {
    SEM: SEM,
    ensureReady: function () {
      return Promise.resolve();
    },
    extractFeatures: extractFeatures,
    inferProbs: inferProbs,
    classifyMesh: classifyMesh,
    applyHullRefinement: applyHullRefinement,
    ifcClassify: ifcClassify,
    semanticToFaceKey: semanticToFaceKey,
    roofPitchFromNormal: roofPitchFromNormal,
    outwardAlignedNormal: outwardAlignedNormal,
    retrainWithCorrections: retrainWithCorrections,
    saveWeightsToStorage: saveWeightsToStorage,
    loadWeightsFromStorage: loadWeightsFromStorage,
    clearPersistedWeights: clearPersistedWeights,
    hasPersistedWeights: function () { return !!hasPersistedWeights; },
    isReady: function () { return !!state.ready && !!state.mlp; },
    usesMlp: function () { return !!state.mlp; }
  };
})(typeof window !== 'undefined' ? window : this);

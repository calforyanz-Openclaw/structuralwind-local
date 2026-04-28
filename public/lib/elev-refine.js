/**
 * Deterministic radial elevation refinement (noise reduction before Mh).
 * Layout: index 0 = site; then for each sub-direction si: nDist samples in order.
 * Does not replace Clause 4.4.2 — only smooths DEM jitter along each ray (median-of-3).
 */
(function (global) {
  function median3(a, b, c) {
    const x = Number(a);
    const y = Number(b);
    const z = Number(c);
    if (x > y) {
      if (y > z) return y;
      if (x > z) return z;
      return x;
    }
    if (x > z) return x;
    if (y > z) return z;
    return y;
  }

  /**
   * @param {number[]} elevations - flat batch from terrain detect
   * @param {number} nSub - sub-directions (16 or 32)
   * @param {number} nDist - distances per ray
   * @returns {number[]} new array (copy)
   */
  function refineRadialElevationsMedian(elevations, nSub, nDist) {
    if (!Array.isArray(elevations) || elevations.length < 1 + nSub * nDist) return elevations.slice();
    const out = elevations.slice();
    for (let si = 0; si < nSub; si++) {
      const base = 1 + si * nDist;
      for (let d = 1; d < nDist - 1; d++) {
        const i0 = base + d - 1;
        const i1 = base + d;
        const i2 = base + d + 1;
        out[i1] = median3(elevations[i0], elevations[i1], elevations[i2]);
      }
    }
    return out;
  }

  global.refineRadialElevationsMedian = refineRadialElevationsMedian;

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { refineRadialElevationsMedian, median3 };
  }
})(typeof self !== 'undefined' ? self : typeof global !== 'undefined' ? global : this);

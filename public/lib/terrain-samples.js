/**
 * Radial terrain sample distances (m) for Mt — AS/NZS 1170.2 profiles.
 * Full: original coarse chain + 15 m steps to 500 m for sharper crest/Lu resolution.
 */
(function (global) {
  const BASE_FULL = [
    30, 60, 90, 120, 150, 180, 210, 240, 270, 300,
    340, 380, 420, 470, 520,
    580, 650, 730, 820, 920,
    1050, 1200, 1400, 1650, 2000, 2500, 3200, 4200, 5000,
  ];
  const NEAR_15 = [];
  for (let d = 15; d <= 500; d += 15) NEAR_15.push(d);

  const FULL = [...new Set([...BASE_FULL, ...NEAR_15])].sort((a, b) => a - b);
  const FAST = [40, 80, 120, 180, 240, 300, 400, 500, 650, 850, 1100, 1500, 2200, 4000];

  function cwTerrainElevPointCount(fast) {
    const nDist = fast ? FAST.length : FULL.length;
    const nSub = fast ? 16 : 32;
    return 1 + nSub * nDist;
  }

  global.CW_SAMPLE_DISTANCES_FULL = FULL;
  global.CW_SAMPLE_DISTANCES_FAST = FAST;
  global.cwTerrainElevPointCount = cwTerrainElevPointCount;

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
      CW_SAMPLE_DISTANCES_FULL: FULL,
      CW_SAMPLE_DISTANCES_FAST: FAST,
      cwTerrainElevPointCount,
    };
  }
})(typeof self !== 'undefined' ? self : typeof global !== 'undefined' ? global : this);

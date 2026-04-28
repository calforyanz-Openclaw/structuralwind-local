/**
 * Hill-shape multiplier Mh — AS/NZS 1170.2 Cl 4.4.2
 * Shared by the browser app and Node tests.
 */
(function (global) {
  /**
   * @param {Array<{dist:number, elev:number}>} upProfile - upwind from site (dist 0 = site)
   * @param {Array<{dist:number, elev:number}>} downProfile - downwind (opposite ray), no duplicate at site
   * @param {number} siteElev
   * @param {number} z - reference height (building height, m)
   */
  function computeMhFromProfiles(upProfile, downProfile, siteElev, z) {
    const NONE = {
      Mh: 1,
      H: 0,
      Lu: 0,
      slope: 0,
      x: 0,
      L1: 0,
      L2: 0,
      crestElev: siteElev,
      crestDist: 0,
      isEsc: false,
    };
    if (upProfile.length < 3) return { ...NONE };

    let best = { ...NONE };

    function interp(d1, e1, d2, e2, eT) {
      if (Math.abs(e1 - e2) < 0.01) return (d1 + d2) / 2;
      return d1 + ((e1 - eT) / (e1 - e2)) * (d2 - d1);
    }

    /** Half-height crossing strictly upwind of crest (increasing dist). */
    function luWindwardSide(up, peakIdx, peakDist, halfH) {
      for (let j = peakIdx; j < up.length - 1; j++) {
        const a = up[j];
        const b = up[j + 1];
        const lo = Math.min(a.elev, b.elev);
        const hi = Math.max(a.elev, b.elev);
        if (halfH < lo - 1e-6 || halfH > hi + 1e-6) continue;
        if (Math.abs(a.elev - b.elev) < 1e-6) continue;
        const dCross = interp(a.dist, a.elev, b.dist, b.elev, halfH);
        if (dCross > peakDist + 0.25) return dCross - peakDist;
      }
      return 0;
    }

    /** Half-height crossing on lee side of crest (toward site, decreasing dist). */
    function luLeeSideOfCrest(up, peakIdx, peakDist, halfH) {
      for (let j = peakIdx; j > 0; j--) {
        const a = up[j - 1];
        const b = up[j];
        const lo = Math.min(a.elev, b.elev);
        const hi = Math.max(a.elev, b.elev);
        if (halfH < lo - 1e-6 || halfH > hi + 1e-6) continue;
        if (Math.abs(a.elev - b.elev) < 1e-6) continue;
        const dCross = interp(a.dist, a.elev, b.dist, b.elev, halfH);
        if (dCross < peakDist - 0.25) return peakDist - dCross;
      }
      if (peakIdx >= 1) {
        const b = up[1];
        const lo = Math.min(siteElev, b.elev);
        const hi = Math.max(siteElev, b.elev);
        if (halfH >= lo - 1e-6 && halfH <= hi + 1e-6 && Math.abs(siteElev - b.elev) >= 1e-6) {
          const dCross = interp(0, siteElev, b.dist, b.elev, halfH);
          if (dCross < peakDist) return peakDist - dCross;
        }
      }
      return 0;
    }

    /** When profile samples miss the half-height crossing, bound Lu from average slope (conservative). */
    function conservativeLuFallback(peakDist, H, ei, elevAtFoot) {
      const slopeApprox = peakDist > 1 ? Math.abs(ei - elevAtFoot) / peakDist : 0.12;
      const luFromSlope = H / 2 / Math.max(slopeApprox, 0.05);
      return Math.min(2500, Math.max(1, luFromSlope));
    }

    function tryCandidate(crestElev, xFromSite, H, Lu, leeDrop, siteDownwind) {
      if (H < 1 || Lu < 1) return;
      const slope = H / (2 * Lu);
      if (slope < 0.05) return;
      const L1 = Math.max(0.36 * Lu, 0.4 * H);
      const isEsc = leeDrop < 0.3 * H;
      const x = xFromSite;
      let L2;
      if (siteDownwind) {
        L2 = isEsc ? 10 * L1 : 4 * Lu;
      } else {
        L2 = 4 * Lu;
      }
      if (x >= L2) return;
      let Mh;
      if (slope <= 0.45) {
        Mh = 1 + (H / (3.5 * (z + L1))) * (1 - x / L2);
      } else if (siteDownwind && x <= L1) {
        Mh = 1 + 0.71 * (1 - x / L2);
      } else {
        Mh = 1 + (H / (3.5 * (z + L1))) * (1 - x / L2);
      }
      Mh = Math.max(Mh, 1.0);
      if (Mh > best.Mh) {
        best = {
          Mh,
          H,
          Lu,
          slope,
          x,
          L1,
          L2,
          crestElev,
          crestDist: x,
          isEsc,
        };
      }
    }

    for (let i = 1; i < upProfile.length; i++) {
      const ei = upProfile[i].elev;
      if (ei <= siteElev) continue;
      const peakDist = upProfile[i].dist;

      let windBase = ei;
      for (let j = i + 1; j < upProfile.length; j++) {
        if (upProfile[j].elev < windBase) windBase = upProfile[j].elev;
      }
      let leeBase = siteElev;
      for (let j = 1; j < i; j++) {
        if (upProfile[j].elev < leeBase) leeBase = upProfile[j].elev;
      }

      const H_wind = ei - windBase;
      const H_lee = ei - leeBase;
      const H = Math.max(H_wind, H_lee);
      const halfH = ei - H / 2;

      let Lu = 0;
      if (H_wind >= H_lee) {
        Lu = luWindwardSide(upProfile, i, peakDist, halfH);
      }
      if (Lu <= 0) {
        Lu = luLeeSideOfCrest(upProfile, i, peakDist, halfH);
      }
      if (Lu <= 0) {
        const elevFoot = H_lee >= H_wind ? leeBase : windBase;
        Lu = conservativeLuFallback(peakDist, H, ei, elevFoot);
      }
      Lu = Math.max(Lu, 1);

      const windwardDrop = H_wind;
      tryCandidate(ei, peakDist, H, Lu, windwardDrop, true);
    }

    {
      let minUpE = siteElev;
      for (let i = 1; i < upProfile.length; i++) {
        if (upProfile[i].elev < minUpE) minUpE = upProfile[i].elev;
      }
      let minDnE = siteElev;
      for (let i = 0; i < downProfile.length; i++) {
        if (downProfile[i].elev < minDnE) minDnE = downProfile[i].elev;
      }

      const H = siteElev - minUpE;
      if (H >= 1) {
        const halfH = siteElev - H / 2;
        let Lu = 0;
        for (let j = 1; j < upProfile.length; j++) {
          if (upProfile[j].elev <= halfH) {
            Lu =
              j > 1
                ? interp(
                    upProfile[j - 1].dist,
                    upProfile[j - 1].elev,
                    upProfile[j].dist,
                    upProfile[j].elev,
                    halfH,
                  )
                : upProfile[j].dist;
            break;
          }
        }
        if (Lu <= 0) {
          const lastE = upProfile[upProfile.length - 1].elev;
          const lastD = upProfile[upProfile.length - 1].dist;
          const drop = siteElev - lastE;
          Lu = drop > 0 ? (lastD * (H / 2)) / drop : lastD;
        }
        if (Lu <= 0) Lu = conservativeLuFallback(upProfile[upProfile.length - 1].dist, H, siteElev, minUpE);
        Lu = Math.max(Lu, 1);
        const leeDrop = siteElev - minDnE;
        tryCandidate(siteElev, 0, H, Lu, leeDrop, true);
      }
      const Hdn = siteElev - minDnE;
      if (Hdn >= 1) {
        const halfH = siteElev - Hdn / 2;
        let Lu = 0;
        for (let j = 0; j < downProfile.length; j++) {
          if (downProfile[j].elev <= halfH) {
            Lu =
              j > 0
                ? interp(
                    downProfile[j - 1].dist,
                    downProfile[j - 1].elev,
                    downProfile[j].dist,
                    downProfile[j].elev,
                    halfH,
                  )
                : interp(0, siteElev, downProfile[0].dist, downProfile[0].elev, halfH);
            break;
          }
        }
        if (Lu <= 0) {
          const lastE = downProfile[downProfile.length - 1].elev;
          const lastD = downProfile[downProfile.length - 1].dist;
          const drop = siteElev - lastE;
          Lu = drop > 0 ? (lastD * (Hdn / 2)) / drop : lastD;
        }
        if (Lu <= 0) Lu = conservativeLuFallback(downProfile[downProfile.length - 1].dist, Hdn, siteElev, minDnE);
        Lu = Math.max(Lu, 1);
        const leeDrop = siteElev - minUpE;
        tryCandidate(siteElev, 0, Hdn, Lu, leeDrop, true);
      }
    }

    return best;
  }

  global.computeMhFromProfiles = computeMhFromProfiles;

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { computeMhFromProfiles };
  }
})(typeof self !== 'undefined' ? self : typeof global !== 'undefined' ? global : this);

# Numeric Parity Report

Generated during live/local validation.

## Scope
Compared live `https://structuralwind.com/` against local `http://127.0.0.1:8018/` using direct Chrome DevTools evaluation on both pages.

## Local serving issue found and fixed during validation
Originally the local root page served `/public/index.html` but did not resolve relative assets like `/script.js` and `/style.css`, so `/` showed HTML without a runnable app.

This was fixed in `server.mjs` by adding a fallback that resolves missing static assets from `public/`.

Post-fix verification:
- `http://127.0.0.1:8018/` → 200
- `http://127.0.0.1:8018/script.js` → 200
- `http://127.0.0.1:8018/style.css` → 200
- local root now exposes `S` and `calc()` correctly

## Test method
For each case, live and local were forced to the same manual state:
- identical geometry
- identical region / terrain category / wind speed inputs
- `Ms = 1.0` in all directions
- `Mt = 1.0` in all directions
- `Mlee = 1.0` in all directions
- `TC_dir = terrainCat` in all directions
- then `calc()` was executed and key outputs were collected from `S.R`

## Case 1 — Gable, manual multipliers
Inputs:
- width 20 m
- depth 15 m
- height 6 m
- pitch 15°
- roofType gable
- terrainCat 3
- windSpeed 45 m/s
- svcVr 32 m/s
- loadCase A
- windAngle 315°
- region NZ1

Key results (live = local):
- `Vsit = 35.4825`
- `qz = 0.8370135`
- `windward_p = 0.7842498876`
- `leeward_p = 0.0309377376`
- `roof_ww_p = -0.2504920268`
- `roof_lw_p = -0.1197246924`
- `roofClause = Table 5.3(B)`

## Case 2 — Monoslope / crosswind, manual multipliers
Inputs:
- width 18 m
- depth 24 m
- height 8 m
- pitch 10°
- roofType monoslope
- parapet 0.3 m
- overhang 0.2 m
- terrainCat 2.5
- windSpeed 52 m/s
- svcVr 36 m/s
- loadCase C
- windAngle 45°
- mapBuildingAngle 30°
- region NZ2

Key results (live = local):
- `Vsit = 45.6534663416`
- `qz = 1.2505433934`
- `windward_p = 0.0324837459`
- `leeward_p = -1.1736653571`
- `roof_ww_p = -1.3647250738`
- `roof_lw_p = -1.1999267683`
- `roofClause = Table 5.3(A)`

## Conclusion
- For controlled manual-input cases, the core numeric calculation engine matches live exactly for the tested outputs.
- The main remaining risk is no longer the manual formula path; it is map-driven / detection-driven behavior (`Ms`, `Mt`, terrain detection, data-source parity).

## Recommended next step
Compare 1–2 real map-driven cases next, now that local root serving is fixed.

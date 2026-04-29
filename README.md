# structuralwind-local

Local-first clone workspace for structuralwind.com.

## Current status
- Public frontend mirrored into `public/`
- Local runnable app served from `server.mjs` at `http://127.0.0.1:8018/`
- Root-path asset loading fixed so `/` is a fully working local app, not just `/public/index.html`
- Firebase auth SDKs, billing/sign-in wording, and visible StructuralWind branding removed from the local app shell
- Saved-project flow now targets local `server.mjs` persistence with browser fallback
- Local `/api/buildings-hybrid` and `/api/elevation-batch` routes wired for terrain/topography detection, with elevation fallback providers added after Open-Meteo rate-limit failures
- Live parity tooling created under `scripts/` and `notes/` for repeatable validation against `https://structuralwind.com/`
- Controlled manual numeric parity validated: 2 hand-crafted cases matched exactly
- Bulk numeric parity validated: 100 randomized controlled cases matched exactly (`notes/bulk-numeric-parity.md`)
- Real-site NZ comparison started for map/data-source parity, with remaining divergence concentrated in terrain/elevation-driven detection rather than the formula core
- Wind-flow / pressure visualization upgraded:
  - pressure-driven pseudo-CFD style flow visualization
  - dynamic flow arrows
  - pressure direction arrows for positive pressure vs suction/uplift
  - enlarged 3D labels for readability

## Validation artifacts
- `notes/live-parity-report.md` — live/local route and content parity notes
- `notes/numeric-parity-report.md` — controlled-case numeric parity notes
- `notes/bulk-numeric-parity.md` / `.json` — 100-case randomized parity run
- `notes/real-site-compare.md` / `.json` — real-address live/local comparison samples
- `scripts/live_parity_check.py` — live/local parity harness
- `scripts/bulk_numeric_parity.js` — randomized numeric parity harness
- `scripts/real_site_compare.js` — real-address comparison harness
- `scripts/bulk_real_site_compare.js` — wider NZ site-coverage comparison harness

## Local preview
- Static mirror: `http://127.0.0.1:8017/public/index.html`
- Local app server: `http://127.0.0.1:8018/`

## Next steps
1. Finish the 100+ NZ real-site live/local comparison pass and summarize the biggest parameter divergences
2. Tighten map/data-source parity, especially `terrainDataSource`, `detectedSiteElev`, `TC_dir`, `Mt`, and downstream `qz` / pressure outputs
3. Keep refining the 3D engineering visualization so flow, suction, uplift, and pressure magnitude are easier to read at a glance
4. Expand local persistence beyond projects if needed (templates/activity history/export metadata)

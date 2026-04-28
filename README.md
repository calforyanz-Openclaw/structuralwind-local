# structuralwind-local

Local-first clone workspace for structuralwind.com.

## Current status
- Public frontend mirrored into `public/`
- Local static server started for verification
- Local Node shim server added in `server.mjs`
- Initial parity gap notes created under `notes/`

## Local preview
- Static mirror: `http://127.0.0.1:8017/public/index.html`
- Local app server: `http://127.0.0.1:8018/`

## Next steps
1. Wire real local implementations for elevation/building endpoints
2. Replace remote auth/billing/persistence with local equivalents
3. Build validation harness against live site
4. Run numerical parity checks with shared test cases

# structuralwind-local

Local-first clone workspace for structuralwind.com.

## Current status
- Public frontend mirrored into `public/`
- Local static server started for verification
- Local Node shim server added in `server.mjs`
- Initial parity gap notes created under `notes/`
- Firebase auth SDKs and billing/sign-in UI removed from the local app shell
- Saved-project flow now targets local `server.mjs` persistence with browser fallback

## Local preview
- Static mirror: `http://127.0.0.1:8017/public/index.html`
- Local app server: `http://127.0.0.1:8018/`

## Next steps
1. Finish replacing remaining remote/shared/cloud-only UI flows with local equivalents or explicit local-build messaging
2. Build validation harness against live site
3. Run numerical parity checks with shared test cases
4. Expand local persistence beyond projects if needed (templates/activity history/export metadata)

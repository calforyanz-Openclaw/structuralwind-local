# Live Parity Report

- Generated: 2026-04-28T19:59:17.272288+00:00
- Live base: `https://structuralwind.com`
- Local base: `http://127.0.0.1:8018`

## index
- live url: `https://structuralwind.com/`
- local url: `http://127.0.0.1:8018/`
- live sha16: `4aad0e1cf465d7fc`
- local sha16: `e8c2a7dce9a29d7d`
- identical: `False`
- live title: `StructuralWind — Professional Wind Analysis`
- local title: `StructuralWind — Professional Wind Analysis`
- live script refs: `17`
- local script refs: `13`
- missing script refs in local: `4`
- added script refs in local: `0`

## about
- live url: `https://structuralwind.com/about.html`
- local url: `http://127.0.0.1:8018/public/about.html`
- live sha16: `ffaa7cee5e4ad979`
- local sha16: `d88c01c42c762bed`
- identical: `False`
- live title: `About — StructuralWind`
- local title: `About — StructuralWind`
- live script refs: `0`
- local script refs: `0`
- missing script refs in local: `0`
- added script refs in local: `0`

## script
- live url: `https://structuralwind.com/script.js`
- local url: `http://127.0.0.1:8018/public/script.js`
- live sha16: `7091873496ac8ff8`
- local sha16: `ffcaf68ce4cd5958`
- identical: `False`
- live api routes found: `18`
- local api routes found: `8`
- routes only in live script: `['/api/activity-log', '/api/activity-log?projectId=', '/api/api-keys', '/api/check-subscription', '/api/create-checkout-session', '/api/create-portal-session', '/api/enterprise-enquiry', '/api/invite-member', '/api/shared-projects', '/api/templates']`
- routes only in local script: `[]`
- live cloud string counts: `{"firebase": 53, "sign in with google": 0, "sign in with microsoft": 0, "create-checkout-session": 1, "create-portal-session": 1, "shared-projects": 9, "api-keys": 6, "enterprise-enquiry": 5}`
- local cloud string counts: `{"firebase": 21, "sign in with google": 0, "sign in with microsoft": 0, "create-checkout-session": 0, "create-portal-session": 0, "shared-projects": 3, "api-keys": 3, "enterprise-enquiry": 0}`

## Findings
- Local index has removed Firebase SDK includes while live still serves them.
- Local auth overlay is localized away from cloud sign-in.
- About page now explicitly documents local-only behavior.
- Shared-project route references still exist in local script; behavior is stubbed but code remains.
- API key overlay code still exists in local script; behavior is stubbed but not removed.

## Next recommended checks
- Remove dead cloud-only overlay/rendering code paths now that local stubs are in place.
- Use browser screenshots to compare the live and local landing pages visually.
- Build one or two shared numeric test cases and compare the calculated outputs between live and local.

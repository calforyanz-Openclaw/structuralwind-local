# StructuralWind Local Clone – Parity Gap (initial)

## Already mirrored
- Exact HTML/CSS/JS assets from public site
- Same public CDN dependencies referenced by the original page
- Same browser-side localStorage/indexedDB helpers

## Backend/API routes detected
- `fetch('/api/activity-log`
- `fetch('/api/activity-log?projectId=`
- `fetch('/api/api-keys`
- `fetch('/api/api-keys/`
- `fetch('/api/check-subscription`
- `fetch('/api/create-checkout-session`
- `fetch('/api/create-portal-session`
- `fetch('/api/delete-project`
- `fetch('/api/enterprise-enquiry`
- `fetch('/api/invite-member`
- `fetch('/api/list-projects`
- `fetch('/api/load-project?id=`
- `fetch('/api/rvt-status?urn=`
- `fetch('/api/save-project`
- `fetch('/api/shared-projects`
- `fetch('/api/templates`
- `fetch('/api/templates/`

## Likely available locally already
- Core UI layout
- Browser-side wind calculation logic
- 3D rendering / visualization
- PDF/export front-end hooks
- Guest-mode elevation/building detection fallback paths
- Local project save/load via localStorage

## Must be reimplemented for full parity
- Firebase auth flows
- Firestore-backed persistence
- Team/shared project flows
- Activity log
- API key management
- Stripe checkout / billing portal
- Any Netlify functions not exposed in client code
- Cloud IFC AI refinement endpoint

## Validation target
- UI pixel parity against live site
- Numerical parity against same inputs on live site
- Endpoint-by-endpoint parity for all accessible user flows

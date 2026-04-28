# Bulk Numeric Parity Report

- Cases: 100
- Seed: 20260429
- Mismatches: 0
- Tolerance: 1e-9

## Coverage
- Randomized geometry, roof types, load cases, regions, terrain categories, openings, wind angles
- Randomized directional arrays for `Md`, `TC_dir`, `Ms`, `Mt`, `Mlee`
- Compared live and local via direct DevTools evaluation of `calc()` outputs

## Result
All 100 generated cases matched exactly within tolerance.

## Notes
- This validates the numeric engine for controlled/manual state injection.
- Next risk surface remains map-driven auto-detection and external data parity.

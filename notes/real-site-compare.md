# Real Site Compare Report

Fixed geometry for all real-site comparisons:
- width 20 m, depth 15 m, height 6 m, pitch 15°, roof gable
- windSpeed 45 m/s, svcVr 32 m/s, loadCase A, windAngle 315°

| Site | Region (live/local) | Source (live/local) | Elev (live/local) | TC_dir head (live/local) | Vsit live | Vsit local | ΔVsit | ΔVsit % | Mt live | Mt local | ΔMt | Ms live | Ms local | ΔMs | qz live | qz local | Δqz | Δroof WW p | Δwindward p |
|---|---|---|---:|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| Auckland CBD | NZ1/NZ1 | live/unknown | 36/null | [3.9,3.93,3.93] / [3.9,3.93,3.93] | 28.325536 | 27.665802 | 0.659734 | -2.329% | 1.185000 | 1.157400 | 0.027600 | 0.740000 | 0.740000 | 0.000000 | 0.481402 | 0.459238 | 0.022164 | -0.006633 | 0.020766 |
| Wellington CBD | NZ2/NZ2 | live/live | 9/9 | [3.87,3.53,2.65] / [3.87,3.33,2.65] | 25.857563 | 25.939956 | -0.082393 | 0.319% | 1.074500 | 1.074500 | 0.000000 | 0.710000 | 0.710000 | 0.000000 | 0.463033 | 0.511430 | -0.048397 | 0.014484 | -0.045346 |
| Christchurch CBD | NZ3/NZ3 | live/live | 10.9/10.9 | [3.38,3.75,3.75] / [2.94,3.75,3.75] | 29.708170 | 29.708170 | 0.000000 | 0.000% | 1.025300 | 1.025300 | 0.000000 | 0.710000 | 0.710000 | 0.000000 | 1.171343 | 1.284112 | -0.112769 | 0.033748 | -0.105660 |

## Notes
- Differences here are map/data-source path differences, not formula-core differences.
- `terrainDataSource` may differ even when outputs match or nearly match because one side can reuse cached terrain/elevation data.

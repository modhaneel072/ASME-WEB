# Build Line WebGL implementation report

## Outcome

The Codex branch implements the complete browser-native Three.js/GSAP motion engine for “Built at Iowa — The Build Line.” One progress value drives the camera, scene weights, assemblies, lighting, atmosphere, Build Line reveal, hotspots, and final frame. The implementation uses the real scrollable document; it does not install a fake scrollbar, lock the page during the experience, load Spline, or preload the legacy 192-frame sequence.

The engine was also tested against the current Claude HTML, CSS, JSON, and shell helper through a local read-only overlay server. The successful path displayed the WebGL canvas above the fallback with `data-webgl="available"`, loaded the real repository content, and produced no browser console errors. Two canonical team-ID changes in Claude-owned JSON remain release-blocking; see `CODEX_REQUESTS.md`.

## Architecture

The runtime flow is:

1. `main.js` claims the DOM contract, keeps the static fallback visible, selects an initial quality, and starts content/vendor loading.
2. `scroll-controller.js` measures each rendered chapter and maps its physical scroll interval onto its declared `data-start`/`data-end` range. GSAP `quickTo` adds 180 ms of restrained mass without replacing native scroll.
3. `camera-path.js` samples one inspectable position/target/roll/FOV path. `scene-controller.js` applies scene-local progress to every mechanical state and renders one valid frame before the shell switches from `pending` to `available`.
4. `quality-manager.js` applies DPR/geometry/effect profiles and observes sustained frame pacing with warm-up, cooldown, and hysteresis.
5. `asset-loader.js` owns optional/lazy assets, content normalization, timeouts, color space, cloning, and disposal. Workshop textures are requested only on approach to Robotic Dog/People.
6. Input, hotspot, debug, fallback, and cleanup controllers remain separate. Page hide, BFCache return, lost WebGL context, shader failure, missing content, and in-flight navigation disposal all have defined paths.

All motion is calculated from normalized scroll progress. There are no independent idle animation clocks that can drift away from the document.

## Modules created

| File | Responsibility |
| --- | --- |
| `static/js/landing/main.js` | Bootstrap, events/datasets, vendor loading, lazy chapter handoff, fallback, teardown |
| `static/js/landing/scroll-controller.js` | Native chapter-to-range mapping, GSAP/ScrollTrigger smoothing, reversible/direct scrolling |
| `static/js/landing/scene-controller.js` | Renderer, all nine scenes, lighting, mechanical progress, observers, metrics, disposal |
| `static/js/landing/camera-path.js` | Central keyframed camera position, target, roll, FOV, easing, and scene weights |
| `static/js/landing/asset-loader.js` | JSON/texture/GLTF loading, optional/required policy, timeouts, lazy limits, disposal |
| `static/js/landing/quality-manager.js` | High/medium/low/fallback profiles and sustained frame-rate adaptation |
| `static/js/landing/input-controller.js` | Skip, optional sound state, panel pause/focus, Escape, visibility, BFCache, reduced motion |
| `static/js/landing/team-interactions.js` | One projected, pointer-safe hotspot per canonical team and `landing:teampanel` events |
| `static/js/landing/materials.js` | Shared named engineering materials |
| `static/js/landing/debug.js` | Opt-in progress/FPS/quality/assets/camera/draw metrics |
| `static/shaders/landing/build-line-shader.js` | Efficient curve-UV Build Line reveal |

The requested material names are present: `machinedAluminum`, `darkSteel`, `mattePolymer`, `printedPolymer`, `rubber`, `glassHousing`, `buildLineGold`, `draftingLine`, and `stressOverlay`.

## Libraries and versions

- Three.js `0.185.1`, locally vendored as production ES modules with its MIT license.
- Three.js `GLTFLoader`, `BufferGeometryUtils`, and `SkeletonUtils` from `0.185.1`, lazy/available for suitable GLB assets.
- GSAP `3.15.0` and ScrollTrigger `3.15.0`, locally vendored with upstream headers and notice.
- No Lenis, Spline, React, bundler, or new Node build system was added.
- DRACO/KTX2 decoder injection points exist, but decoder binaries were not shipped because no selected landing asset requires them.

## Scene implementation status

| Range | Scene | Implemented state |
| --- | --- | --- |
| `0.00–0.08` | Intro | Near-black drafting grid, ticks, restrained camera, gold datum draw |
| `0.08–0.18` | Design | Sketch/axis, wireframe, solid hub, bearings, bolt pattern, exploded assembly |
| `0.18–0.31` | IAM3D | Generic additive rover mechanism, rails, brackets, axles, wheels, layer lines |
| `0.31–0.43` | Autonomous Plane | Student UAV silhouette assembled from airfoil, ribs, spar, control surfaces, propeller, electronics volume |
| `0.43–0.56` | MATE ROV | Controlled pool/test course, frame, thrusters, camera housing, tether, manipulator, task ring |
| `0.56–0.68` | Student Design | Crank/link/slider cycle, active load path, brief controlled overload, revised geometry |
| `0.68–0.82` | Robotic Dog | Ten-step exposed quadruped assembly, actuator-like joints, fasteners, tray/wiring/sensor, settle and gait |
| `0.82–0.92` | People | Workshop doorway and lazy real-photo planes with centered cover crop and restrained parallax |
| `0.92–1.00` | Join | Motion settles; Build Line forms a quiet mechanical CTA frame |

The IAM3D, Plane, ROV, Student Design, and Robotic Dog geometry is intentionally generic procedural implementation geometry. Comments in `scene-controller.js` explicitly state that it is not a digital twin or representation of the teams’ current machines.

## Assets used

Only repository photography that truthfully fits the People/workshop transition is loaded into WebGL:

| Asset | Dimensions | Bytes | Policy |
| --- | ---: | ---: | --- |
| `static/images/gallery/electronics_lab.jpg` | 1600×1067 | 247,911 | Low/medium/high lazy layer |
| `static/images/gallery/team_presenting.jpg` | 1067×1600 | 191,226 | Low/medium/high lazy portrait layer |
| `static/images/makeathon/IMG_3507.jpg` | 1600×1200 | 342,344 | Medium/high lazy layer |
| `static/images/makeathon/IMG_5970.jpg` | 1600×900 | 172,480 | High-only lazy layer |

The low/medium/high lazy photo payloads are 439,137 / 781,481 / 953,961 raw bytes respectively. None is part of the initial scene load. Texture UVs use an aspect-preserving centered cover crop instead of stretching people.

No model was copied or optimized because the repository has no truthful, landing-ready Plane, MATE ROV, SDC, or Robotic Dog model. The authentic IAM3D arm simulator assets are roughly 17.3 MiB/~363k triangles for the 19 loaded STL parts before the larger unused chassis assets; shipping them without merge/decimation would have violated the landing budget. `static/models/landing` therefore remains empty.

Explicitly not used: the Spline humanoid, humanoid GLB/poster, synthetic rover frame sequence, qyrox/Veo video, synthetic reference image, beach HDR, oversized ASME logo texture, and unrelated/generic project images.

## Performance measurements

Measurements were taken on the final code in the Codex in-app Chromium browser against a local HTTP server at 1440×900, forced medium quality. The browser exposed a high-refresh `requestAnimationFrame` cadence but not stable hardware/GPU identity, so the FPS sample is reported as an observation rather than a device-wide claim.

| Scene midpoint | Draw calls | Triangles |
| --- | ---: | ---: |
| Intro | 3 | 2,560 |
| Design | 19 | 3,200 |
| IAM3D | 31 | 3,316 |
| Autonomous Plane | 18 | 2,790 |
| MATE ROV | 54 | 5,118 |
| Student Design | 8 | 2,748 |
| Robotic Dog | 31 | 3,636 |
| People | 12 | 2,612 |
| Join | 4 | 2,564 |

- Observed debug cadence during final merged-shell sampling: 169.8–240.4 callbacks/second. This is not a mobile benchmark and is not presented as guaranteed production FPS.
- Peak measured medium scene cost: 54 draws and 5,118 triangles (MATE ROV).
- Initial on-disk JavaScript payload, before HTTP compression: 989,849 bytes. This consists of 121,409 bytes of landing engine/shader source, 750,938 bytes of initial Three modules, and 117,502 bytes of GSAP/ScrollTrigger.
- Three loader/util addons add 164,175 raw bytes only when imported; GLTFLoader is dynamically imported on first GLTF request.
- No post-processing pass, bloom, depth-of-field, chromatic aberration, or full-screen underwater effect is present.
- The old 192-frame sequence (about 21 MiB compressed and roughly 1.3 GiB if retained decoded) is never requested.

The local server measurements do not include production compression/CDN latency. A representative physical mobile-device profile is still required before claiming the 30+ FPS mobile target across hardware.

## Quality behavior

| Mode | DPR cap | Main reductions |
| --- | ---: | --- |
| High | 1.75 | 220-segment Build Line, limited shadows, 72 restrained underwater points, four photo layers |
| Medium | 1.35 | 160-segment Build Line, no dynamic shadows, 30 points, lower water transmission, three photo layers |
| Low | 1.0 | 100-segment Build Line, simplified ribs/detail, no shadows/particles/transmission, two photo layers |
| Fallback | 1.0 | Canvas hidden; polished DOM/photo shell and all links remain active |

Selection considers WebGL2, viewport, DPR, mobile UA, hardware concurrency, and device memory when available. Runtime observation waits through a six-second warm-up, evaluates sustained samples, and enforces a 15-second change cooldown. High and medium can step down on ordinary sustained poor pacing. Low intentionally retains a stable 30 FPS cadence and enters fallback only after three severe windows, preventing quality thrash and premature mobile fallback.

A downgrade rebuilds the Build Line geometry, adjusts water transmission, shadows, particles, visible detail, DPR, and future lazy photo count. Shader compile failure, lost context, or sustained severe low-mode pacing activates the same terminal fallback.

## Browser and responsive testing

The final engine was tested both in the isolated contract harness and against the current Claude shell/data through a local overlay. The latter is the closest available pre-merge integration test and did not modify Claude-owned files.

| Viewport | Result |
| --- | --- |
| 1440×900 | WebGL available, fallback removed, real content loaded, no horizontal overflow |
| 1280×720 | Canvas 1265×720, navigation contained, no horizontal overflow |
| 768×1024 | Canvas 753×1024, navigation contained, no horizontal overflow |
| 390×844 | Low quality active, canvas 375×844, native vertical scroll, no horizontal overflow, hotspot and controls contained |

Verified behaviors:

- Incremental forward and backward native scrolling; chapter-height mapping stayed on declared progress even where the mobile People chapter rendered much taller than its nominal timeline share.
- Mouse-wheel/trackpad-style increments, resize, responsive reflow, Skip Experience, final CTA focus, and four final routes.
- One projected hotspot per scene; canonical event payload; IAM3D shell panel content; Escape/close/backdrop; focus restoration; renderer pause and resume.
- Sound control visible and muted by default even with no sound assets.
- Forced fallback: `data-webgl="fallback"`, WebGL hidden, fallback visible, no Spline/old-frame requests, all four CTA links retained.
- Missing optional asset, missing content fallback, in-flight disposal, and lazy-quality limiting.
- Navigation away/back restored the same scene with WebGL available; persisted page hide pauses and page show resumes.
- Normal merged-shell path produced no browser console errors or uncaught promise rejections.

The available browser controller did not expose real network throttling, real touch hardware, or browser chrome zoom. Missing-asset tests cover the network failure path; 125%-equivalent effective viewport resizing and the four required responsive sizes cover layout scaling. Physical touch, browser zoom, slow-network, and representative mobile GPU checks remain release-device QA rather than claimed automation coverage.

## Accessibility and reduced motion

- The browser remains the scroller. Wheel, touch, arrows, Space, Page Up/Down, Home, and End are never globally prevented.
- Canvas and WebGL layers are `aria-hidden` and pointer-inert; navigation and CTA links stay above them.
- Skip moves directly to the final CTA and focuses its first action.
- Team panel opening pauses rendering. Engine and shell-driven closes stay synchronized through observed panel state, including backdrop and DOM-only team links.
- Visibility loss pauses the render loop; BFCache return resumes it; navigation disposal releases observers, renderer, geometry, materials, and textures.
- Reduced motion removes camera inertia/roll, parallax, wheel/propeller rotation, particles, overload pulse, gait, and progressive rapid assembly. Scenes present direct, stable assembled states while native scroll and content remain intact.
- Audio is optional, muted by default, and loads no audio payload. The control state remains functional for future gesture-gated assets.

## Tests

`node --test tests/landing/*.test.mjs` passes 21 tests covering:

- exact continuous scene ranges and camera interpolation;
- scene weight normalization;
- quality selection, DPR caps, 30 FPS low retention, severe fallback hysteresis, and event payloads;
- content failure and Claude array/key normalization;
- optional assets, background-preload loader isolation, lazy limits, and dispose-while-inflight behavior;
- native scrolling and non-proportional chapter mapping;
- canonical team hotspot events;
- skip, sound, panel/Escape/shell-close state, visibility, and BFCache behavior;
- representative reduced-motion mechanical invariants.

All 11 landing JavaScript/shader modules also pass `node --check`.

The Flask application was run with an in-memory SQLite URL and integrations disabled. `/`, `/healthz`, `main.js`, the Three module, and GSAP each returned HTTP 200 with the expected content type. Backend, database, templates, CSS, portal, and admin files were not edited.

## Known limitations and real media still needed

- Real, approved Autonomous Plane, MATE ROV, Student Design, and Robotic Dog photography/models are still needed. Current procedural geometry is an honest editorial stand-in.
- A landing-optimized, merged/decimated IAM3D model could replace the generic rover mechanism after technical/art review.
- No approved audio exists, so the toggle controls architecture/state only; it does not synthesize or autoplay placeholder sound.
- No production device lab was available. Physical iOS/Android touch, browser zoom, throttled network, and average mobile GPU profiling remain launch QA.
- Current Claude JSON still assigns two legacy team keys and two misleading/weak images; exact fixes are in `CODEX_REQUESTS.md`.
- Some procedural primitive detail is created at initial quality and hidden rather than rebuilt on later downgrade; the largest profile-dependent geometry/effect costs do rebuild or disable.

## Integration instructions

1. Apply every item in `docs/landing/CODEX_REQUESTS.md` on the Claude branch.
2. Merge the Claude shell branch and `codex/landing-webgl` through the normal review flow; do not copy the old inline rover/Spline implementation back into the template.
3. Confirm `landing.html` loads `static/css/landing.css`, the relocated shell helper, and `static/js/landing/main.js` as an ES module.
4. Verify the merged page with `?landingDebug=1`, then remove the query for normal presentation. Use `?landingFallback=1` for the explicit fallback check.
5. Re-run the 21 Node tests, Flask smoke checks, all four viewports, the five canonical team panels, and a real mobile-device pass.

No push, merge, deployment, backend migration, or database mutation is part of this branch.

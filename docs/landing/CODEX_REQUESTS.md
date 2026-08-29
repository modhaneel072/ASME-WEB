# Claude-owned landing requests

These are the only remaining changes required in Claude-owned files before the two landing branches are merged.

## 1. Canonicalize the Plane and ROV team keys

In `static/data/landing-content.json`:

- Change the Autonomous Plane object in `teams[]` from `"id": "plane"` to `"id": "autonomous-plane"`.
- Change the MATE ROV object in `teams[]` from `"id": "rov"` to `"id": "mate-rov"`.
- Change the matching `chapters[].team` references from `"plane"` to `"autonomous-plane"` and from `"rov"` to `"mate-rov"`.
- Keep the chapter/WebGL scene names `plane` and `rov`; those are the required timeline scene keys.

This is release-blocking. The WebGL event contract correctly emits `autonomous-plane` and `mate-rov`, while the current shell panel lookup still expects `plane` and `rov`.

## 2. Use the exact Robotic Dog status

In the Robotic Dog team object in `static/data/landing-content.json`:

- Add `"status": "Selective team — interview required"`.
- Change `interview_note` to the same exact sentence case: `"Selective team — interview required"`.

The template currently renders `interview_note`, so changing both fields keeps the data source and visible copy exact.

## 3. Correct two media assignments

In `static/data/landing-content.json`:

- Change both Student Design Competition references to `images/gallery/competition.jpg` from `kind: "photo"` to the existing `kind: "placeholder"` form with `src` and `alt` set to `null`, and caption `STUDENT DESIGN — PROJECT MEDIA IN PROGRESS`. That photo is visibly an IAM3D Challenge shirt and must not be presented as SDC media.
- Replace the People-strip entry `images/makeathon/IMG_5848.jpg` with `images/makeathon/IMG_3507.jpg`, then update its alt text/caption to describe the real maker-lab scene. The current frame is dominated by beverage cans rather than engineering work.

## 4. Move the shell helper out of the Codex-owned directory and align event fields

The brief assigns all of `static/js/landing/*` to Codex. Move the current shell helper from `static/js/landing/ui.js` to `static/js/landing-shell.js`, and update the deferred script URL in `templates/site/landing.html` accordingly.

While moving that file, make these exact event-consumer changes:

- In the `landing:scenechange` listener, read `event.detail.current` instead of `event.detail.scene`.
- In the `landing:qualitychange` listener, read `event.detail.current` instead of `event.detail.quality`.
- In the `landing:error` listener, set `document.documentElement.dataset.webgl` to `fallback`, or leave that dataset entirely under `main.js`; do not set it to `unavailable`.

No other HTML, CSS, Flask, database, portal, or admin change is requested.

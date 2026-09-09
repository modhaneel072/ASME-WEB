# Reference recording — observations

**Recording status: NOT SUPPLIED.** No file matching
`MaintainX and 1 more page - Personal - Microsoft Edge 2026-09-08 14-58-54.mp4` (or any
other product recording) was present in the workspace, the uploads folder, or
`Downloads` when this build ran, and `ffmpeg`/`ffprobe` are not installed on the
build machine. No frames were extracted.

Everything below is therefore taken from the master specification's own written
description of the recording (§3, §6, §7) and is labelled accordingly. Nothing here is
OBSERVED first-hand.

| Pattern | Label | Source |
|---|---|---|
| Fixed left sidebar with uppercase group labels (Setup / Work / Optimize / Manage) | INFERRED | spec §3, §4 |
| Slim setup banner across the top of the workspace | INFERRED | spec §3, §7.2 |
| Large page title left; search + primary split button right; filter chips beneath | INFERRED | spec §3, §7.3 |
| Master-detail split ~40/60 or 42/58 | INFERRED | spec §9.1, §40 |
| Right-side create pane ≈58 % of main content width, list stays visible | INFERRED | spec §10.3 |
| To Do / Done tabs, sort selector, saved filters ("My Filters") | INFERRED | spec §10.1, §26 |
| Empty state: small illustration, one-line title, direct CTA link | INFERRED | spec §10.2, §36 |
| Two-column report card grid | INFERRED | spec §3, §25 |
| Light borders, minimal shadows, restrained motion | INFERRED | spec §6.1, §6.6 |
| Colour tokens, type scale, spacing scale, control heights | PROPOSED | spec §6.2–6.5 (the spec proposes these values; they were not measured from footage) |

## What was built from this

The app shell, headers, filter bar, master-detail layout and create pane follow the
INFERRED patterns above with the PROPOSED measurements from §6. When a recording is
supplied, run the extraction described in the master prompt §4, fill
`design-measurements.md` with measured values, and diff them against the tokens in
`apps/ops-web/src/ui/tokens.css`.

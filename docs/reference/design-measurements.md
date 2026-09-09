# Design measurements

No recording was supplied (see `video-observations.md`), so there are no measured
values. The values below are the spec's PROPOSED numbers (§6.3–6.5) and are what the
tokens in `apps/ops-web/src/ui/tokens.css` implement.

| Element | Value | Label |
|---|---|---|
| Desktop sidebar width | 264 px (collapsed 64 px) | PROPOSED |
| Setup banner height | 40 px | PROPOSED |
| Page header row | 64–72 px | PROPOSED |
| Standard control height | 40 px (compact 32 px) | PROPOSED |
| Filter chip height | 36 px | PROPOSED |
| Table row height | 48–56 px | PROPOSED |
| Content padding | 24 px desktop, 16 px tablet | PROPOSED |
| Sidebar row | 40 px high, 16 px horizontal padding, nested indent 20–24 px | PROPOSED |
| Active nav row | pale blue background + 3 px primary left accent | PROPOSED |
| Page title | 32 px / 700–750 / 1.15 | PROPOSED |
| Section heading | 20–22 px / 650–700 | PROPOSED |
| Body | 14 px / 400–450 | PROPOSED |
| Label | 13 px / 550–600 | PROPOSED |
| Table text | 13 px | PROPOSED |
| Spacing scale | 4, 8, 12, 16, 20, 24, 32, 40, 48 | PROPOSED |
| Radii | 4 / 6 / 8 / pill | PROPOSED |
| Motion | hover 100–140 ms; pane 180 ms ease-out, opacity + 8 px x; popover 120–160 ms scale .98→1 | PROPOSED |
| Split proportions | list 40–44 % / detail 56–60 %; create pane ≈58 % of main content | PROPOSED |

Acceptance viewports (spec §9): 1440×900 (primary), 1366×768, 768×1024, 390×844.
Playwright screenshots for these live under `apps/ops-web/e2e/screenshots/` after
`npm run test:e2e`.

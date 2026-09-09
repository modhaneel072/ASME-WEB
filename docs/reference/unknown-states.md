# Unknown states

States the spec itself lists as not exposed by the recording (§42), plus what this build
could not verify because no recording was available at all.

## Not evidenced by any reference (built from the spec text only)

- Every hover, focus, validation, loading, success and failure state.
- Work-order detail with realistic data in every card.
- The exact composition of the setup banner and its dismissal affordance.
- Sidebar collapsed state.
- Mobile / tablet layouts.
- Notification centre, settings screens, user/team management forms.

## Decisions taken in the absence of evidence

| Question | Decision | Label |
|---|---|---|
| What happens on tablet when a list and a detail cannot fit side by side? | Detail becomes a stacked route with a back control that restores the list scroll position (spec §35) | PROPOSED |
| Where do "Critical" priority and "Blocked" live? | Critical is a fifth segment only shown to roles holding `work_order.set_critical`; Blocked is a flag with a badge, never a status (spec §10.4, §10.6) | PROPOSED |
| Which view is the default for Work Orders? | Panel view, To Do tab, sorted "Due date: soonest first" | PROPOSED |
| Does the setup banner persist per device or per user? | Per user (stored on the membership) so it survives devices; admins can reopen setup from Settings (spec §7.2) | INFERRED |
| Modal vs pane for create? | Right-side pane for work orders and projects; small dialogs only for confirmations (cancel, delete) | INFERRED |

When a recording arrives, verify each row above and change the label to OBSERVED or
adjust the implementation.

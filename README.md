# ASME @ UIowa Web Platform

Single coherent Flask app with:
- Public marketing site
- Member/team-leader/admin portal
- Inventory checkout + returns
- Print request queue
- Google Calendar viewing + slot-based scheduling
- NFC attendance + kiosk login entry flows

## Roles
- `member`
- `team_leader`
- `admin`

Server-side RBAC is enforced on portal routes.

## Key routes

### Public
- `/` Home
- `/about` (alias: `/who-we-are`)
- `/exec` (alias: `/executive-team`)
- `/projects`
- `/projects/<slug>`
- `/join`
- `/contact` (alias: `/socials`)
- `/login`
- `/signup`
- `/forgot-password`

### Portal member
- `/portal/member`
- `/portal/member/inventory`
- `/portal/member/inventory/<id>`
- `/portal/member/my-items` (alias: `/portal/member/checkouts`)
- `/portal/member/prints`
- `/portal/member/calendar`
- `/portal/member/profile`

### Team leader / admin scheduling
- `/portal/member/schedule`
- `/portal/leader/schedule` (alias)

### Admin
- `/portal/admin`
- `/portal/admin/members`
- `/portal/admin/attendance`
- `/portal/admin/inventory`
- `/portal/admin/prints`
- `/portal/admin/members/import-roster` (POST)
- `/portal/admin/members/import-credentials/<filename>` (GET)
- `/portal/admin/settings`

Hidden admin routes (kept alive for compatibility) redirect to `/portal/admin`:
- `/portal/admin/nfc`
- `/portal/admin/calendar`
- `/portal/admin/exports`
- `/portal/admin/content`
- `/portal/admin/announcements`

### Shared NFC URL tag flows
- `/kiosk` shared login tag entry (login/signup -> inventory)
- `/checkin` shared attendance tag flow
- `/checkin/select`
- `/checkin/success`

## Legacy ops compatibility

Set `ASME_ENABLE_LEGACY_OPS=1` to keep old ops pages active.

Default is `ASME_ENABLE_LEGACY_OPS=0`, so legacy routes redirect into portal equivalents:
- `/dashboard` -> `/portal`
- `/attendance` -> `/portal/admin/attendance`
- `/inventory` -> `/portal/member/inventory`
- `/prints` -> `/portal/member/prints`
- `/activity` -> `/portal/admin/inventory`
- `/settings` -> `/portal/admin/settings`
- `/scan` -> `/kiosk`
- `/my-items` -> `/portal/member/my-items`
- `/admin/nfc` -> `/portal/admin/nfc`
- `/calendar` -> `/portal/member/calendar`
- `/app` -> `/kiosk`

## Google Calendar scheduling setup

Required env vars for slot-based scheduling:
- `GOOGLE_SERVICE_ACCOUNT_JSON` (raw JSON string OR absolute path to JSON file)
- `GOOGLE_CALENDAR_ID_ROBOTICS`
- `GOOGLE_CALENDAR_ID_FLUIDS`
- `GOOGLE_CALENDAR_TIMEZONE` (default `America/Chicago`)
- `CALENDAR_SCHEDULING_DAYS` (default `14`)
- `CALENDAR_WORK_HOURS_START` (default `08:00`)
- `CALENDAR_WORK_HOURS_END` (default `22:00`)

Scheduling uses Google Calendar `freebusy.query` to compute available slots and `events.insert` to create meetings.

## Core env vars

See `.env.example`. Most important:
- `ASME_SECRET_KEY`
- `ASME_DATABASE_URL`
- `ASME_CALENDAR_PROVIDER=google`
- `ASME_GOOGLE_CALENDAR_EMBED_URL` (optional if room IDs are configured; embed can be generated)
- `ASME_ENABLE_LEGACY_OPS=0`
- `ASME_SESSION_IDLE_MINUTES=240`
- `ASME_ADMIN_SESSION_IDLE_MINUTES=30`
- `ASME_SESSION_COOKIE_SECURE=1` (set `0` for plain-http local testing)

## Bulk roster import

On **Admin -> Members / Roles**, use **Bulk Import Roster**:
- Upload roster PDF.
- Accounts are created/updated in bulk.
- Username format: first letter of first name + last name.
- Password format: first 2 letters of first name + first letter of last name + 5 digits.
- A credentials CSV is generated for download.

Member login accepts **email or username**.

## Quick start (Windows PowerShell)
```powershell
cd C:\Users\modha\Desktop\asme-inventory
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
python db_init.py
python app.py
```

Open `http://127.0.0.1:5000`.

## Smoke checks
After startup, verify:
- `/`
- `/login`
- `/portal/member`
- `/portal/admin`
- `/kiosk`
- `/checkin`

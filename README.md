# ASME @ UIowa Web Platform

Modern Flask full-stack app with:
- Public marketing site
- Role-based member/team leader/admin portal
- Inventory checkout/return
- NFC tag mapping + attendance
- 3D print request queue
- Calendar and meeting workflows

## Tech stack
- Python + Flask
- Flask-SQLAlchemy (SQLite by default, Postgres-compatible URI)
- Jinja templates + custom CSS

## Roles
- `guest` (not logged in)
- `member`
- `team_leader`
- `admin`

## Pages
- Public: `/`, `/who-we-are`, `/executive-team`, `/projects`, `/projects/<slug>`, `/join`, `/sponsors`, `/contact`, `/login`, `/admin-login`, `/signup`, `/forgot-password`
- Portal router: `/portal`
- Member:
  - `/portal/member` (dashboard tiles)
  - `/portal/member/inventory`
  - `/portal/member/items/<id>`
  - `/portal/member/checkouts`
  - `/portal/member/prints`
  - `/portal/member/prints/<id>`
  - `/portal/member/calendar`
  - `/portal/member/help`
  - `/portal/member/profile`
- Team leader: `/portal/team`
- Admin:
  - `/portal/admin` (KPI dashboard)
  - `/portal/admin/members`
  - `/portal/admin/nfc`
  - `/portal/admin/attendance`
  - `/portal/admin/inventory`
  - `/portal/admin/prints`
  - `/portal/admin/announcements`
- Existing ops backend remains available at `/dashboard`, `/inventory`, `/attendance`, `/prints`, `/calendar`

## Database models
- `users`
- `nfc_tags`
- `items` (inventory items)
- `transactions` (inventory checkout/return records)
- `print_requests`
- `events`
- `attendance_records`
- `projects`
- `contact_messages`
- `announcements`
- `audit_logs`

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

Open: `http://127.0.0.1:5000`

## Demo credentials
After `python db_init.py`, default seeded users use:
- password: `ChangeMe123!`
- sample emails:
  - `jordan@uiowa.edu` (admin)
  - `morgan@uiowa.edu` (team_leader)
  - `avery@uiowa.edu` (member)

## Environment variables
See `.env.example`.

## Render deploy notes
- Build command: `pip install -r requirements.txt`
- Start command: `python app.py`
- Set `ASME_DATABASE_URL` to a valid SQLAlchemy URL (or keep SQLite for testing).

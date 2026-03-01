# Outlook Calendar + Email Cancellation Setup

## 1) Choose the mailbox and calendars
1. Decide which mailbox owns events (example: `asme@yourtenant.edu`).
   - Use a Microsoft 365 mailbox in your Entra tenant.
   - Gmail/Yahoo/iCloud mailboxes cannot be used with Microsoft Graph Calendar API.
   - For app-only (client credentials) flow, avoid `common`, `organizations`, and `consumers` as tenant values.
2. In Outlook, use one calendar for all rooms or separate calendars for:
   - `ASME Robotics Room`
   - `ASME Fluids Lab`
3. Copy each Outlook calendar ID (Graph API ID) if you want per-room calendars.

## 2) Create Azure app credentials (for website -> Microsoft Graph)
1. Open Azure Portal -> Microsoft Entra ID -> App registrations.
2. Create a new app registration.
3. Create a client secret and copy its value.
4. Under API permissions, add Microsoft Graph **Application** permission:
   - `Calendars.ReadWrite`
5. Grant admin consent for your tenant.

## 3) Configure environment values
Edit `instance/print_commands.env` and set:

```env
ASME_OUTLOOK_TENANT_ID=your-tenant-id
ASME_OUTLOOK_CLIENT_ID=your-app-client-id
ASME_OUTLOOK_CLIENT_SECRET=your-app-client-secret
ASME_OUTLOOK_CALENDAR_USER=asme@yourtenant.edu

ASME_OUTLOOK_CALENDAR_ROBOTICS_ID=optional_calendar_id
ASME_OUTLOOK_CALENDAR_FLUIDS_ID=optional_calendar_id
ASME_OUTLOOK_CALENDAR_TZ=Central Standard Time

ASME_SMTP_HOST=smtp.office365.com
ASME_SMTP_PORT=587
ASME_SMTP_USER=asme@yourtenant.edu
ASME_SMTP_PASS=your_smtp_password_or_app_password
ASME_CANCEL_NOTIFY_TO=asme@yourtenant.edu
```

If you use one shared calendar for both rooms:

```env
ASME_OUTLOOK_CALENDAR_ID=your_shared_calendar_id
```

## 4) Optional live calendar embed on `/calendar`
Set:

```env
ASME_OUTLOOK_CALENDAR_EMBED_URL=your_outlook_embed_url
```

If this is not set, booking/cancel sync still works, but the iframe will not display.

## 5) Restart and verify
1. Restart app: `python app.py`.
2. Open `/calendar`.
3. Run sync doctor:
   - `python scripts/outlook_sync_doctor.py`
   - Result should show: `READY: Outlook sync can create real events.`
4. Book a test meeting and verify it appears in Outlook.
5. Click `Cancel Request`, approve from email, and verify it is removed.

# Google Calendar + Gmail Cancellation Setup

## 1) Create/choose your Google Calendar
1. Log in as `uiowaasme@gmail.com`.
2. In Google Calendar, create one calendar (or two):
   - `ASME Robotics Room`
   - `ASME Fluids Lab`
3. Copy each calendar ID from:
   - Calendar `Settings and sharing` -> `Integrate calendar` -> `Calendar ID`.

## 2) Create a Google service account key (for website -> Calendar API)
1. Go to Google Cloud Console.
2. Create/select a project.
3. Enable `Google Calendar API`.
4. Create a Service Account.
5. Create a JSON key and download it.
6. Place the JSON file at:
   - `instance/google_service_account.json`

## 3) Share calendar(s) with the service account
1. Open each calendar in Google Calendar settings.
2. Under `Share with specific people`, add the service account email
   (looks like `name@project-id.iam.gserviceaccount.com`).
3. Grant permission: `Make changes to events`.

## 4) Configure app env values
Edit `instance/print_commands.env` and set:

```env
ASME_GCAL_SERVICE_ACCOUNT_FILE=instance/google_service_account.json
ASME_GOOGLE_CALENDAR_ROBOTICS_ID=your_robotics_calendar_id@group.calendar.google.com
ASME_GOOGLE_CALENDAR_FLUIDS_ID=your_fluids_calendar_id@group.calendar.google.com
ASME_GOOGLE_CALENDAR_TZ=America/Chicago

ASME_SMTP_HOST=smtp.gmail.com
ASME_SMTP_PORT=587
ASME_SMTP_USER=uiowaasme@gmail.com
ASME_SMTP_PASS=your_gmail_app_password
ASME_CANCEL_NOTIFY_TO=uiowaasme@gmail.com
```

If you use one shared calendar for both rooms, set:

```env
ASME_GOOGLE_CALENDAR_ID=your_shared_calendar_id@group.calendar.google.com
```

## 5) Gmail app password (required for email confirmations)
1. In `uiowaasme@gmail.com`, enable 2-Step Verification.
2. Create an `App Password` for Mail.
3. Put that password in `ASME_SMTP_PASS`.

## 6) Restart app and verify
1. Restart server: `python app.py`.
2. Open `/calendar`.
3. In `Calendar Sync Status`, verify:
   - service file found
   - calendar IDs set
   - SMTP ready

## 7) How workflow now works
1. User books meeting on website -> app creates Google Calendar event.
2. User clicks `Cancel Request` -> app emails `uiowaasme@gmail.com`.
3. Admin clicks `Confirm cancellation` link in email -> app removes Google event and local meeting.

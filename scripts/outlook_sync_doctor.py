import os
import sys
from pathlib import Path
from urllib.parse import quote

# Allow running from project root or from scripts/ directly.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import (
    _http_json_request,
    get_outlook_access_token,
    get_outlook_calendar_id_for_room,
    validate_outlook_sync_config,
)


def print_section(title):
    print(f"\n== {title} ==")


def main():
    print("ASME Outlook Sync Doctor")

    errors, warnings = validate_outlook_sync_config()
    print_section("Config Validation")
    if errors:
        print("Status: FAILED")
        for err in errors:
            print(f"- {err}")
    else:
        print("Status: OK")

    if warnings:
        print_section("Warnings")
        for warn in warnings:
            print(f"- {warn}")

    if errors:
        return 1

    token, token_error = get_outlook_access_token()
    print_section("OAuth Token")
    if token_error:
        print(f"Status: FAILED\n- {token_error}")
        return 1
    print("Status: OK")

    mailbox_user = (os.environ.get("ASME_OUTLOOK_CALENDAR_USER") or "").strip()
    encoded_user = quote(mailbox_user, safe="")
    headers = {"Authorization": f"Bearer {token}"}

    status, payload, error = _http_json_request(
        method="GET",
        url=f"https://graph.microsoft.com/v1.0/users/{encoded_user}/calendars?$top=10",
        headers=headers,
        retries=1,
    )
    print_section("Mailbox Access")
    if status != 200:
        msg = payload.get("error", {}).get("message") if isinstance(payload.get("error"), dict) else None
        print(f"Status: FAILED\n- {msg or error or 'unable to list calendars'}")
        return 1

    calendars = payload.get("value") or []
    print(f"Status: OK\n- Found {len(calendars)} calendar(s) for {mailbox_user}.")
    if calendars:
        print("- Available calendars:")
        for cal in calendars:
            cal_name = (cal.get("name") or "Unnamed").strip()
            cal_id = (cal.get("id") or "").strip()
            print(f"  * {cal_name} | id={cal_id}")

    configured_ids = {}
    default_id = (os.environ.get("ASME_OUTLOOK_CALENDAR_ID") or "").strip()
    robotics_id = get_outlook_calendar_id_for_room("Robotics Room")
    fluids_id = get_outlook_calendar_id_for_room("Fluids Lab")
    if default_id:
        configured_ids["default"] = default_id
    if robotics_id:
        configured_ids["robotics"] = robotics_id
    if fluids_id:
        configured_ids["fluids"] = fluids_id

    if configured_ids:
        print_section("Configured Calendar IDs")
        for label, cal_id in configured_ids.items():
            encoded_cal_id = quote(cal_id, safe="")
            st, pay, err = _http_json_request(
                method="GET",
                url=f"https://graph.microsoft.com/v1.0/users/{encoded_user}/calendars/{encoded_cal_id}",
                headers=headers,
                retries=1,
            )
            if st == 200:
                cal_name = (pay.get("name") or "Unnamed").strip()
                print(f"- {label}: OK ({cal_name})")
            else:
                msg = pay.get("error", {}).get("message") if isinstance(pay.get("error"), dict) else None
                print(f"- {label}: FAILED ({msg or err or 'not found'})")
                return 1

    print_section("Result")
    print("READY: Outlook sync can create real events.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

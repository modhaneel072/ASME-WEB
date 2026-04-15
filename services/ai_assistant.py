"""
ASME @ UIowa — Admin Analytics Assistant (Claude with tool use).

Exposes a set of SQLAlchemy-backed tools that Claude can call to answer
analytics questions about members, inventory, prints, attendance, etc.
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import func

# Lazy import to avoid circular imports at module load time.
def _db():
    from app import db
    return db

def _models():
    from models import (
        User, Member, Item, Transaction, PrintRequest,
        Event, AttendanceRecord, Announcement, ContactMessage,
    )
    return {
        "User": User, "Member": Member, "Item": Item,
        "Transaction": Transaction, "PrintRequest": PrintRequest,
        "Event": Event, "AttendanceRecord": AttendanceRecord,
        "Announcement": Announcement, "ContactMessage": ContactMessage,
    }


# ── TOOL DEFINITIONS (what Claude sees) ─────────────────────────────────
TOOLS = [
    {
        "name": "count_active_members",
        "description": "Return the number of active members and a breakdown by role (admin, team_leader, member).",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_pending_prints",
        "description": "List pending 3D print requests awaiting admin approval or currently printing. Returns up to 25 results.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_print_queue_stats",
        "description": "Get aggregate stats on 3D print requests: total, by status, by printer.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_overdue_checkouts",
        "description": "List inventory items currently checked out past their due date.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_active_checkouts",
        "description": "List all items currently checked out (status=OUT), with who has them and when.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "search_inventory",
        "description": "Search inventory items by name, category, or location keyword. Returns name, category, location, available_qty, total_qty.",
        "input_schema": {
            "type": "object",
            "properties": {"query": {"type": "string", "description": "Search keyword"}},
            "required": ["query"],
        },
    },
    {
        "name": "most_checked_out_items",
        "description": "Return the top N most-frequently checked-out items of all time.",
        "input_schema": {
            "type": "object",
            "properties": {"limit": {"type": "integer", "description": "Number of items", "default": 10}},
            "required": [],
        },
    },
    {
        "name": "get_recent_signups",
        "description": "Members created within the last N days.",
        "input_schema": {
            "type": "object",
            "properties": {"days": {"type": "integer", "description": "Lookback window in days", "default": 30}},
            "required": [],
        },
    },
    {
        "name": "get_upcoming_events",
        "description": "List scheduled events from now into the future. Up to 10 results.",
        "input_schema": {
            "type": "object",
            "properties": {"limit": {"type": "integer", "description": "Max events", "default": 10}},
            "required": [],
        },
    },
    {
        "name": "get_attendance_stats",
        "description": "Attendance stats across all events or for a specific event id. Returns total_checkins, unique_members, event breakdown.",
        "input_schema": {
            "type": "object",
            "properties": {"event_id": {"type": "integer", "description": "Optional specific event id"}},
            "required": [],
        },
    },
    {
        "name": "get_inactive_members",
        "description": "Members who haven't logged in within the last N days (based on last_login_at).",
        "input_schema": {
            "type": "object",
            "properties": {"days": {"type": "integer", "description": "Days since last login", "default": 30}},
            "required": [],
        },
    },
    {
        "name": "get_unread_contact_messages",
        "description": "Contact form messages with status 'new' (unreviewed).",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "search_members",
        "description": "Find members/users by partial name, email, or username.",
        "input_schema": {
            "type": "object",
            "properties": {"query": {"type": "string", "description": "Name, email, or username fragment"}},
            "required": ["query"],
        },
    },
    {
        "name": "get_dashboard_summary",
        "description": "Quick overview: active member count, pending prints, items checked out, overdue count, today's attendance, unread messages.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
]


# ── TOOL IMPLEMENTATIONS ─────────────────────────────────────────────────
def tool_count_active_members() -> dict:
    M = _models()
    User = M["User"]
    rows = (
        _db().session.query(User.role, func.count(User.id))
        .filter(User.is_active.is_(True))
        .group_by(User.role)
        .all()
    )
    return {
        "total": sum(c for _, c in rows),
        "by_role": {r or "member": c for r, c in rows},
    }


def tool_get_pending_prints() -> dict:
    M = _models()
    PrintRequest = M["PrintRequest"]
    User = M["User"]
    rows = (
        _db().session.query(PrintRequest, User.name)
        .outerjoin(User, User.id == PrintRequest.user_id)
        .filter(PrintRequest.status.in_(["submitted", "approved", "printing"]))
        .order_by(PrintRequest.created_at.desc())
        .limit(25)
        .all()
    )
    return {
        "count": len(rows),
        "items": [
            {
                "id": pr.id,
                "submitted_by": name or "Unknown",
                "printer": pr.printer_type,
                "status": pr.status,
                "filament": pr.filament,
                "priority": pr.priority,
                "deadline": pr.deadline.isoformat() if pr.deadline else None,
                "created_at": pr.created_at.strftime("%Y-%m-%d %H:%M"),
            }
            for pr, name in rows
        ],
    }


def tool_get_print_queue_stats() -> dict:
    M = _models()
    PrintRequest = M["PrintRequest"]
    by_status = dict(
        _db().session.query(PrintRequest.status, func.count(PrintRequest.id))
        .group_by(PrintRequest.status).all()
    )
    by_printer = dict(
        _db().session.query(PrintRequest.printer_type, func.count(PrintRequest.id))
        .group_by(PrintRequest.printer_type).all()
    )
    return {
        "total": sum(by_status.values()),
        "by_status": by_status,
        "by_printer": by_printer,
    }


def tool_get_overdue_checkouts() -> dict:
    M = _models()
    Transaction, Item, Member = M["Transaction"], M["Item"], M["Member"]
    today = datetime.now().date()
    rows = (
        _db().session.query(Transaction, Item.name, Member.name)
        .join(Item, Item.id == Transaction.item_id)
        .outerjoin(Member, Member.id == Transaction.member_id)
        .filter(Transaction.status == "OUT", Transaction.due_date != None, Transaction.due_date < today)  # noqa: E711
        .limit(50)
        .all()
    )
    return {
        "count": len(rows),
        "items": [
            {
                "transaction_id": tx.id,
                "item": iname,
                "member": mname or "Unknown",
                "qty": tx.qty,
                "due_date": tx.due_date.isoformat() if tx.due_date else None,
                "days_overdue": (today - tx.due_date).days if tx.due_date else 0,
            }
            for tx, iname, mname in rows
        ],
    }


def tool_get_active_checkouts() -> dict:
    M = _models()
    Transaction, Item, Member = M["Transaction"], M["Item"], M["Member"]
    rows = (
        _db().session.query(Transaction, Item.name, Member.name)
        .join(Item, Item.id == Transaction.item_id)
        .outerjoin(Member, Member.id == Transaction.member_id)
        .filter(Transaction.status == "OUT")
        .order_by(Transaction.checkout_time.desc())
        .limit(50)
        .all()
    )
    return {
        "count": len(rows),
        "items": [
            {
                "item": iname,
                "member": mname or "Unknown",
                "qty": tx.qty,
                "checkout_time": (tx.checkout_time or tx.timestamp).strftime("%Y-%m-%d %H:%M"),
                "due_date": tx.due_date.isoformat() if tx.due_date else None,
            }
            for tx, iname, mname in rows
        ],
    }


def tool_search_inventory(query: str) -> dict:
    M = _models()
    Item = M["Item"]
    q = f"%{query.lower()}%"
    rows = (
        Item.query.filter(Item.active.is_(True))
        .filter(
            func.lower(Item.name).like(q)
            | func.lower(Item.category).like(q)
            | func.lower(Item.location).like(q)
        )
        .limit(25).all()
    )
    return {
        "count": len(rows),
        "items": [
            {
                "name": i.name,
                "category": i.category,
                "location": i.location,
                "available": i.available_qty,
                "total": i.total_qty,
            }
            for i in rows
        ],
    }


def tool_most_checked_out_items(limit: int = 10) -> dict:
    M = _models()
    Transaction, Item = M["Transaction"], M["Item"]
    rows = (
        _db().session.query(Item.name, func.count(Transaction.id).label("c"))
        .join(Transaction, Transaction.item_id == Item.id)
        .filter(Transaction.action == "checkout")
        .group_by(Item.id, Item.name)
        .order_by(func.count(Transaction.id).desc())
        .limit(limit).all()
    )
    return {"items": [{"name": n, "checkouts": c} for n, c in rows]}


def tool_get_recent_signups(days: int = 30) -> dict:
    M = _models()
    User = M["User"]
    cutoff = datetime.now() - timedelta(days=days)
    rows = (
        User.query.filter(User.created_at >= cutoff)
        .order_by(User.created_at.desc()).limit(100).all()
    )
    return {
        "count": len(rows),
        "window_days": days,
        "members": [
            {
                "name": u.name, "email": u.email, "role": u.role,
                "joined": u.created_at.strftime("%Y-%m-%d") if u.created_at else None,
            }
            for u in rows
        ],
    }


def tool_get_upcoming_events(limit: int = 10) -> dict:
    M = _models()
    Event = M["Event"]
    rows = (
        Event.query.filter(Event.status != "cancelled", Event.start_time >= datetime.now())
        .order_by(Event.start_time.asc()).limit(limit).all()
    )
    return {
        "count": len(rows),
        "events": [
            {
                "id": e.id, "title": e.title,
                "location": e.location, "status": e.status,
                "start": e.start_time.strftime("%Y-%m-%d %H:%M") if e.start_time else None,
                "end": e.end_time.strftime("%Y-%m-%d %H:%M") if e.end_time else None,
            }
            for e in rows
        ],
    }


def tool_get_attendance_stats(event_id: int | None = None) -> dict:
    M = _models()
    AttendanceRecord, Event = M["AttendanceRecord"], M["Event"]
    q = _db().session.query(AttendanceRecord)
    if event_id is not None:
        q = q.filter(AttendanceRecord.event_id == event_id)
    total = q.count()
    unique_members = q.with_entities(AttendanceRecord.member_id).distinct().count()

    if event_id is None:
        # overall breakdown per event
        breakdown = (
            _db().session.query(Event.title, func.count(AttendanceRecord.id))
            .join(AttendanceRecord, AttendanceRecord.event_id == Event.id)
            .group_by(Event.id, Event.title)
            .order_by(func.count(AttendanceRecord.id).desc()).limit(10).all()
        )
        return {
            "total_checkins": total,
            "unique_members": unique_members,
            "top_events": [{"event": t, "checkins": c} for t, c in breakdown],
        }

    event = Event.query.get(event_id)
    return {
        "event": event.title if event else f"Event #{event_id}",
        "total_checkins": total,
        "unique_members": unique_members,
    }


def tool_get_inactive_members(days: int = 30) -> dict:
    M = _models()
    User = M["User"]
    cutoff = datetime.now() - timedelta(days=days)
    rows = (
        User.query.filter(User.is_active.is_(True))
        .filter((User.last_login_at == None) | (User.last_login_at < cutoff))  # noqa: E711
        .order_by(User.last_login_at.asc().nullsfirst() if hasattr(User.last_login_at.asc(), 'nullsfirst') else User.last_login_at.asc())
        .limit(100).all()
    )
    return {
        "count": len(rows),
        "window_days": days,
        "members": [
            {
                "name": u.name, "email": u.email, "role": u.role,
                "last_login": u.last_login_at.strftime("%Y-%m-%d") if u.last_login_at else "Never",
            }
            for u in rows
        ],
    }


def tool_get_unread_contact_messages() -> dict:
    M = _models()
    ContactMessage = M["ContactMessage"]
    rows = (
        ContactMessage.query.filter_by(status="new")
        .order_by(ContactMessage.created_at.desc()).limit(50).all()
    )
    return {
        "count": len(rows),
        "messages": [
            {
                "id": m.id, "name": m.name, "email": m.email,
                "subject": m.subject, "kind": m.kind,
                "created": m.created_at.strftime("%Y-%m-%d %H:%M") if m.created_at else None,
            }
            for m in rows
        ],
    }


def tool_search_members(query: str) -> dict:
    M = _models()
    User = M["User"]
    q = f"%{query.lower()}%"
    rows = (
        User.query.filter(
            func.lower(User.name).like(q)
            | func.lower(User.email).like(q)
            | func.lower(User.username).like(q)
        ).limit(25).all()
    )
    return {
        "count": len(rows),
        "members": [
            {
                "name": u.name, "email": u.email, "username": u.username,
                "role": u.role, "active": bool(u.is_active),
                "last_login": u.last_login_at.strftime("%Y-%m-%d") if u.last_login_at else "Never",
            }
            for u in rows
        ],
    }


def tool_get_dashboard_summary() -> dict:
    M = _models()
    User = M["User"]
    PrintRequest = M["PrintRequest"]
    Transaction = M["Transaction"]
    AttendanceRecord = M["AttendanceRecord"]
    ContactMessage = M["ContactMessage"]

    today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    today = today_start.date()

    return {
        "active_members": User.query.filter(User.is_active.is_(True)).count(),
        "pending_prints": PrintRequest.query.filter(PrintRequest.status.in_(["submitted", "approved", "printing"])).count(),
        "items_checked_out": Transaction.query.filter(Transaction.status == "OUT").count(),
        "overdue_checkouts": Transaction.query.filter(
            Transaction.status == "OUT", Transaction.due_date != None, Transaction.due_date < today  # noqa: E711
        ).count(),
        "checkins_today": AttendanceRecord.query.filter(AttendanceRecord.checkin_time >= today_start).count(),
        "unread_messages": ContactMessage.query.filter_by(status="new").count(),
    }


TOOL_MAP = {
    "count_active_members": tool_count_active_members,
    "get_pending_prints": tool_get_pending_prints,
    "get_print_queue_stats": tool_get_print_queue_stats,
    "get_overdue_checkouts": tool_get_overdue_checkouts,
    "get_active_checkouts": tool_get_active_checkouts,
    "search_inventory": tool_search_inventory,
    "most_checked_out_items": tool_most_checked_out_items,
    "get_recent_signups": tool_get_recent_signups,
    "get_upcoming_events": tool_get_upcoming_events,
    "get_attendance_stats": tool_get_attendance_stats,
    "get_inactive_members": tool_get_inactive_members,
    "get_unread_contact_messages": tool_get_unread_contact_messages,
    "search_members": tool_search_members,
    "get_dashboard_summary": tool_get_dashboard_summary,
}


SYSTEM_PROMPT = """You are ASKme Analytics, the AI assistant for admins of ASME at the University of Iowa student chapter.

Answer questions using ONLY the tools provided. Never invent data or stats.
If the tools don't have what's needed, say so and suggest which page in the portal can help.

Style guide:
- Be concise. Admins are busy.
- Lead with the answer. Put numbers and names up front.
- Use compact lists when returning multiple rows.
- Never say "I'll now call tool X" — just call it.
- If a question spans multiple areas (e.g., "who has overdue items and also hasn't logged in"), call multiple tools.

You have access to live data about: members, inventory, 3D print requests, checkouts, events, attendance, and contact messages."""


def run_assistant(user_message: str, history: list | None = None, max_steps: int = 6) -> dict:
    """
    Run a single conversation turn with Claude. Returns {"reply": str, "tool_calls": [...]}.
    `history` is the prior list of {"role": "user"|"assistant", "content": ...} messages.
    """
    try:
        from anthropic import Anthropic
    except ImportError:
        return {"reply": "The anthropic Python package is not installed. Run: pip install anthropic", "tool_calls": []}

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return {
            "reply": "ANTHROPIC_API_KEY is not set. Add it to your .env file to enable the analytics assistant.",
            "tool_calls": [],
        }

    client = Anthropic(api_key=api_key)
    messages = list(history or [])
    messages.append({"role": "user", "content": user_message})

    tool_calls_summary = []

    for _ in range(max_steps):
        resp = client.messages.create(
            model=os.environ.get("ANTHROPIC_MODEL", "claude-haiku-4-5-20251001"),
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages,
        )

        if resp.stop_reason == "tool_use":
            assistant_content = resp.content
            messages.append({"role": "assistant", "content": assistant_content})

            tool_results = []
            for block in resp.content:
                if block.type == "tool_use":
                    name = block.name
                    args = block.input or {}
                    tool_calls_summary.append({"name": name, "args": args})
                    fn = TOOL_MAP.get(name)
                    try:
                        result = fn(**args) if fn else {"error": f"Unknown tool {name}"}
                    except Exception as e:
                        result = {"error": f"{type(e).__name__}: {e}"}
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": json.dumps(result, default=str)[:4000],
                    })
            messages.append({"role": "user", "content": tool_results})
            continue

        # end of turn — collect text
        text = "".join(b.text for b in resp.content if getattr(b, "type", None) == "text")
        return {"reply": text or "(no response)", "tool_calls": tool_calls_summary, "history": messages}

    return {"reply": "Stopped after too many tool steps.", "tool_calls": tool_calls_summary}

from asme import events
from asme.extensions import db
from asme.models import Member, User
from asme.services import roster


ENTRIES = [
    {"name": "Grace Hopper", "first_name": "Grace", "last_name": "Hopper", "email": "ghopper@uiowa.edu"},
    {"name": "Alan Turing", "first_name": "Alan", "last_name": "Turing", "email": ""},
]


def test_import_creates_members_and_users(app, captured_events):
    result = roster.import_roster_entries(ENTRIES)
    db.session.commit()
    roster.emit_import_events(result)
    assert result["created_members"] == 2 and result["created_users"] == 2
    assert User.query.filter_by(email="ghopper@uiowa.edu").first().username == "ghopper"
    # missing email gets a placeholder domain
    turing = User.query.filter(User.name == "Alan Turing").first()
    assert turing.email.endswith("@asme.local")
    assert sum(1 for name, _ in captured_events if name == events.USER_CREATED) == 2
    assert any(name == events.CHAPTER_CHANGED for name, _ in captured_events)


def test_import_is_idempotent(app):
    roster.import_roster_entries(ENTRIES)
    db.session.commit()
    second = roster.import_roster_entries(ENTRIES)
    db.session.commit()
    assert second["created_members"] == 0 and second["created_users"] == 0 and second["updated_users"] == 2
    assert User.query.count() == 2 and Member.query.count() == 2


def test_import_never_downgrades_admin(app, users):
    result = roster.import_roster_entries([{"name": "Ada Admin", "first_name": "Ada", "last_name": "Admin", "email": "ada@uiowa.edu"}])
    db.session.commit()
    assert User.query.filter_by(email="ada@uiowa.edu").first().role == "admin"
    assert result["updated_users"] == 1


def test_credentials_csv_matches_applied_passwords(app):
    roster.import_roster_entries(ENTRIES)
    db.session.commit()
    result = roster.import_roster_entries(ENTRIES, reset_existing_passwords=False)
    db.session.commit()
    assert all(row["password"] == "" for row in result["credentials"])


def test_parse_pdf_lines(monkeypatch):
    class Page:
        def extract_text(self):
            return "First Name Last Name Email\nGrace Hopper ghopper@uiowa.edu Yes\nAlan Turing\n"

    class Reader:
        def __init__(self, *_a, **_k):
            self.pages = [Page()]

    monkeypatch.setattr(roster, "PdfReader", Reader)
    rows = roster.parse_roster_pdf_entries(b"%PDF")
    assert rows == [
        {"name": "Grace Hopper", "first_name": "Grace", "last_name": "Hopper", "email": "ghopper@uiowa.edu"},
        {"name": "Alan Turing", "first_name": "Alan", "last_name": "Turing", "email": ""},
    ]

import os

from asme.config import Settings


def test_defaults_are_sane(monkeypatch):
    for key in list(os.environ):
        if key.startswith("ASME_") or key.startswith("GOOGLE_") or key.startswith("CALENDAR_"):
            monkeypatch.delenv(key, raising=False)
    cfg = Settings.from_env()
    assert cfg.env == "production"
    assert cfg.database_url.startswith("sqlite")
    assert cfg.member_session_idle_minutes == 240
    assert cfg.admin_session_idle_minutes == 30
    assert cfg.calendar_provider == "google"
    assert cfg.onboarding_enforce is False
    assert cfg.auto_migrate is False  # production never auto-migrates


def test_validate_flags_default_secret_in_production(monkeypatch):
    cfg = Settings.from_env(env="production", secret_key="asme-dev-secret")
    problems = cfg.validate()
    assert any("ASME_SECRET_KEY" in p for p in problems)


def test_validate_flags_half_configured_outlook():
    cfg = Settings.from_env(env="development", outlook_tenant_id="abc", outlook_client_id="", outlook_client_secret="", outlook_calendar_user="")
    assert any("Outlook" in p for p in cfg.validate())


def test_env_parsing(monkeypatch):
    monkeypatch.setenv("ASME_ENV", "development")
    monkeypatch.setenv("ASME_SESSION_IDLE_MINUTES", "15")
    monkeypatch.setenv("ASME_ONBOARDING_ENFORCE", "yes")
    monkeypatch.setenv("ASME_SESSION_COOKIE_SAMESITE", "none")
    monkeypatch.setenv("ASME_ADMIN_EMAILS", "A@x.com, b@y.org")
    cfg = Settings.from_env()
    assert cfg.env == "development"
    assert cfg.member_session_idle_minutes == 15
    assert cfg.onboarding_enforce is True
    assert cfg.session_cookie_samesite == "None"
    assert cfg.admin_emails == frozenset({"a@x.com", "b@y.org"})
    assert cfg.auto_migrate is True

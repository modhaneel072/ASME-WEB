"""Compatibility shim: models now live in ``asme.models``.

Scripts that ``from models import db, User`` keep working.
"""

from asme.models import *  # noqa: F401,F403
from asme.models import __all__  # noqa: F401

"""Compatibility shim: the assistant moved to ``asme.integrations.assistant``."""

from asme.integrations.assistant import *  # noqa: F401,F403
from asme.integrations.assistant import SYSTEM_PROMPT, TOOL_MAP, TOOLS, run_assistant  # noqa: F401

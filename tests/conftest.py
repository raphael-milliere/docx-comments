"""Shared fixtures: keep the suite hermetic against machine state."""

import pytest

from docx_comments import system_author


@pytest.fixture(autouse=True)
def _hermetic_author_env(monkeypatch):
    """No test may depend on the developer's Office install or env vars."""
    monkeypatch.delenv("DOCX_COMMENTS_AUTHOR_DOCX", raising=False)
    monkeypatch.setattr(
        system_author, "_system_office_user_info", lambda: (None, None)
    )

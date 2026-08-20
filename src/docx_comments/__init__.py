"""
python-docx-comments: Complete Word document comment manipulation.

Distributed on PyPI as ``python-docx-comments``; imported as ``docx_comments``.

This module provides full OOXML comment support including:
- Adding anchored comments to specific text ranges
- Replying to existing comments (threaded)
- Marking comments as resolved
- Full Word Online compatibility
"""

from importlib.metadata import PackageNotFoundError, version

from docx_comments.exceptions import CommentNotFoundError, PersonNotFoundError
from docx_comments.manager import CommentManager, PersonSpec
from docx_comments.models import CommentContent, CommentInfo, CommentThread, PersonInfo

try:
    __version__ = version("python-docx-comments")
except PackageNotFoundError:  # pragma: no cover - local checkout without metadata
    __version__ = "0.0.0"
__all__ = [
    "CommentManager",
    "CommentContent",
    "CommentThread",
    "CommentInfo",
    "PersonInfo",
    "PersonSpec",
    "CommentNotFoundError",
    "PersonNotFoundError",
]

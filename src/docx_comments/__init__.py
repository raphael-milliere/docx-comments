"""
docx-comments: Complete Word document comment manipulation.

This module provides full OOXML comment support including:
- Adding anchored comments to specific text ranges
- Replying to existing comments (threaded)
- Marking comments as resolved
- Full Word Online compatibility
"""

from importlib.metadata import PackageNotFoundError, version

from docx_comments.manager import CommentManager, PersonSpec
from docx_comments.models import CommentInfo, CommentThread, PersonInfo

try:
    __version__ = version("docx-comments")
except PackageNotFoundError:  # pragma: no cover - local checkout without metadata
    __version__ = "0.0.0"
__all__ = ["CommentManager", "CommentThread", "CommentInfo", "PersonInfo", "PersonSpec"]

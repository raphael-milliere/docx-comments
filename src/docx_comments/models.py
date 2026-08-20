"""Data models for comment information."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Optional, Sequence, Tuple, Union

RunSpec = Union[str, Tuple[str, Dict[str, bool]]]
"""A comment run: plain text, or (text, {"bold"/"italic"/"underline": True})."""

ParagraphSpec = Union[str, Sequence[RunSpec]]
"""A comment paragraph: plain text or a sequence of runs."""

CommentContent = Union[str, Sequence[ParagraphSpec]]
"""Comment content: a plain string (one paragraph) or paragraph sequence."""


@dataclass
class CommentInfo:
    """Information about a single comment."""

    comment_id: Optional[str]
    """Unique comment ID (w:id attribute). None for schema-invalid comments
    that lack w:id (the library still lists them)."""

    para_id: str
    """Paragraph ID linking to extended/ids parts (w14:paraId). Empty string
    ("") when the comment has no identifiable paragraph id."""

    text: str
    """Comment text content."""

    author: str
    """Comment author name."""

    initials: Optional[str] = None
    """Author initials (optional)."""

    timestamp: Optional[datetime] = None
    """Comment creation timestamp."""

    parent_para_id: Optional[str] = None
    """Parent paragraph ID for replies (w15:paraIdParent)."""

    is_resolved: bool = False
    """Whether comment is marked as done (w15:done)."""

    durable_id: Optional[str] = None
    """Durable ID for persistence (w16cid:durableId)."""

    @property
    def is_reply(self) -> bool:
        """Check if this comment is a reply to another comment."""
        return self.parent_para_id is not None


@dataclass
class PersonInfo:
    """Information about a person entry in people.xml."""

    author: str
    """Person author name (w15:person/@w15:author)."""

    provider_id: Optional[str] = None
    """Presence provider ID (w15:presenceInfo/@w15:providerId)."""

    user_id: Optional[str] = None
    """Presence user ID (w15:presenceInfo/@w15:userId)."""

    @property
    def has_presence(self) -> bool:
        """Check if presence metadata is present."""
        return bool(self.provider_id and self.user_id)


@dataclass
class CommentThread:
    """A comment thread with root comment and replies."""

    root: CommentInfo
    """The root (parent) comment of the thread."""

    replies: list[CommentInfo] = field(default_factory=list)
    """List of reply comments in chronological order."""

    @property
    def all_comments(self) -> list[CommentInfo]:
        """All comments in the thread (root + replies)."""
        return [self.root] + self.replies

    @property
    def is_resolved(self) -> bool:
        """Check if the thread is resolved (root comment is done)."""
        return self.root.is_resolved

    @property
    def reply_count(self) -> int:
        """Number of replies in the thread."""
        return len(self.replies)

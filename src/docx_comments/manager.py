"""Main CommentManager class for Word document comment manipulation."""

from __future__ import annotations

import os
import random
import uuid
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Any, Iterator, Optional, Union

from lxml import etree

from docx_comments.anchors import CommentAnchor
from docx_comments.exceptions import CommentNotFoundError
from docx_comments.models import (
    CommentContent,
    CommentInfo,
    CommentThread,
    PersonInfo,
)
from docx_comments.system_author import _default_person_from_system
from docx_comments.xml_parts import (
    CommentsExtendedPart,
    CommentsExtensiblePart,
    CommentsIdsPart,
    CommentsPart,
    PeoplePart,
    ensure_comment_parts,
    ensure_mc_ignorable,
    validate_xml_text,
)

# Backwards-compatible alias: this helper used to live in this module.
_validate_xml_text = validate_xml_text

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# OOXML Namespaces
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
NS_XML = "http://www.w3.org/XML/1998/namespace"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

PersonSpec = Union[PersonInfo, str, dict[str, Any], bool]

# Private RNG so consumer calls to random.seed() cannot make id sequences
# collide across comments.
_rng = random.Random()

# w:id values are ST_DecimalNumber; the OOXML ecosystem (Word, the Open XML
# SDK, python-docx, LibreOffice) treats them as 32-bit signed integers, so
# generated ids must stay within that range.
_MAX_ID = 0x7FFFFFFE


def _qn(ns: str, name: str) -> str:
    """Create qualified name with namespace."""
    return f"{{{ns}}}{name}"


def _generate_id() -> str:
    """Generate a random comment ID (positive 32-bit integer as string)."""
    return str(_rng.randint(1, _MAX_ID))


def _generate_long_hex_id() -> str:
    """Generate an 8-hex-digit ST_LongHexNumber within the valid range."""
    return f"{_rng.randint(1, _MAX_ID):08X}"


def _format_utc(dt: datetime) -> str:
    """Format a timezone-aware datetime in UTC."""
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_comment_date(date_str: Optional[str]) -> Optional[datetime]:
    """Parse a comment date string into a tz-aware datetime."""
    if not date_str:
        return None
    try:
        if date_str.endswith("Z"):
            return datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        parsed = datetime.fromisoformat(date_str)
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except ValueError:
        return None


def _coerce_comment_id(comment_id: Union[int, str]) -> str:
    """Accept python-docx native int ids alongside this library's str ids."""
    if isinstance(comment_id, bool):
        raise TypeError("comment_id must be a str or int, not bool")
    if isinstance(comment_id, int):
        return str(comment_id)
    if not isinstance(comment_id, str):
        raise TypeError(
            f"comment_id must be a str or int, got {type(comment_id).__name__}"
        )
    return comment_id


class CommentManager:
    """
    Manager for Word document comments.

    Provides complete comment manipulation including:
    - Adding anchored comments to specific text ranges
    - Replying to existing comments (threaded)
    - Marking comments as resolved
    - Full Word Online compatibility

    Example:
        >>> from docx import Document
        >>> from docx_comments import CommentManager, PersonInfo
        >>>
        >>> doc = Document("document.docx")
        >>> mgr = CommentManager(doc)
        >>>
        >>> # Add comment
        >>> comment_id = mgr.add_comment(
        ...     paragraph=doc.paragraphs[0],
        ...     text="Review this",
        ...     author=PersonInfo(author="Reviewer")
        ... )
        >>>
        >>> # Reply to comment
        >>> reply_id = mgr.reply_to_comment(
        ...     comment_id,
        ...     "Fixed",
        ...     PersonInfo(author="Author")
        ... )
        >>>
        >>> doc.save("reviewed.docx")
    """

    def __init__(self, document: Document, auto_migrate: bool = False) -> None:
        """
        Initialize CommentManager with a python-docx Document.

        Comment parts are created lazily on the first mutating operation, so
        a manager used only for reading leaves the document untouched.

        Args:
            document: A python-docx Document instance.
            auto_migrate: Whether to backfill missing comment metadata on init
                (this creates the comment parts if they are missing).
        """
        self._document = document
        self._comments_handler: Optional[CommentsPart] = None
        if auto_migrate:
            self.migrate_comment_metadata()

    def _ensure_parts(self) -> None:
        """Ensure all required comment parts exist in the document."""
        ensure_comment_parts(self._document)

    def _comments_part(self) -> CommentsPart:
        """Get the cached comments part handler."""
        if self._comments_handler is None:
            self._comments_handler = CommentsPart(self._document)
        return self._comments_handler

    @property
    def _comments_xml(self) -> etree._Element:
        """Get the comments.xml root element."""
        return self._comments_part().xml

    def _save_comments(self) -> None:
        """Save changes to comments.xml."""
        if self._comments_handler is not None:
            self._comments_handler._save()

    def _comment_id_exists(self, comment_id: str) -> bool:
        """Check whether comments.xml contains a comment with this id."""
        for elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            if elem.get(_qn(NS_W, "id")) == comment_id:
                return True
        return False

    def _anchor_roots(self) -> list[etree._Element]:
        """XML roots that can carry anchors (body, headers/footers, notes)."""
        return list(CommentAnchor(self._document)._iter_anchor_roots())

    def _new_comment_id(self) -> str:
        """Generate a comment id unique within the document."""
        used: set[str] = set()
        for elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            value = elem.get(_qn(NS_W, "id"))
            if value:
                used.add(value)
        # Anchors may reference ids with no comment element (other tools);
        # avoid colliding with those too, in every anchor-bearing part.
        for root in self._anchor_roots():
            for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
                for elem in root.iter(_qn(NS_W, tag)):
                    value = elem.get(_qn(NS_W, "id"))
                    if value:
                        used.add(value)
        while True:
            candidate = _generate_id()
            if candidate not in used:
                return candidate

    def _used_long_hex_ids(self) -> set[str]:
        """Collect ST_LongHexNumber values already present in the document.

        Covers comment paraId/textId, paragraph paraIds in every story part
        (body, headers/footers, footnotes/endnotes), threading entries,
        durable ids, and commentsExtensible entries: new ids are drawn
        outside this pool so they cannot collide.
        """
        used: set[str] = set()
        for comment_elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            for para in comment_elem.findall(_qn(NS_W, "p")):
                for attr in ("paraId", "textId"):
                    value = para.get(_qn(NS_W14, attr))
                    if value:
                        used.add(value.upper())
        for root in self._anchor_roots():
            for para in root.iter(_qn(NS_W, "p")):
                value = para.get(_qn(NS_W14, "paraId"))
                if value:
                    used.add(value.upper())
        used.update(
            pid.upper()
            for pid in CommentsExtendedPart(self._document).get_threading_info()
        )
        for para_id, durable_id in CommentsIdsPart(self._document).get_durable_ids().items():
            used.add(para_id.upper())
            used.add(durable_id.upper())
        used.update(
            durable.upper()
            for durable in CommentsExtensiblePart(self._document).get_extensible_info()
        )
        return used

    @staticmethod
    def _new_long_hex_id(used: set[str]) -> str:
        """Draw a fresh hex id outside `used`, adding it to the set."""
        while True:
            candidate = _generate_long_hex_id()
            if candidate not in used:
                used.add(candidate)
                return candidate

    @staticmethod
    def _primary_para_id(
        para_ids: list[str],
        threading: dict[str, dict],
        durable_ids: dict[str, str],
    ) -> Optional[str]:
        """The paraId that keys a comment's satellite metadata.

        Word keys commentsExtended/commentsIds to the LAST paragraph of a
        comment; prefer a paraId with a threading entry, then one with a
        durable id, then the last paragraph.
        """
        for pid in reversed(para_ids):
            if pid in threading:
                return pid
        for pid in reversed(para_ids):
            if pid in durable_ids:
                return pid
        return para_ids[-1] if para_ids else None

    def _comment_index(
        self,
    ) -> tuple[list[CommentInfo], dict[str, CommentInfo], dict[str, CommentInfo]]:
        comments = list(self.list_comments())
        by_id = {c.comment_id: c for c in comments if c.comment_id is not None}
        by_para_id = {c.para_id: c for c in comments if c.para_id}
        return comments, by_id, by_para_id

    def _root_for(
        self, comment: CommentInfo, by_para_id: dict[str, CommentInfo]
    ) -> CommentInfo:
        current = comment
        seen: set[str] = set()
        while current.parent_para_id and current.parent_para_id in by_para_id:
            if current.parent_para_id in seen:
                break
            seen.add(current.parent_para_id)
            current = by_para_id[current.parent_para_id]
        return current

    @staticmethod
    def _thread_key(comment: CommentInfo) -> str:
        # Degenerate comments may lack both ids; key them by object identity
        # so they form singleton threads instead of aliasing to one key.
        return comment.para_id or comment.comment_id or f"__anon_{id(comment)}"

    def _thread_comments_for(self, comment_id: str) -> list[CommentInfo]:
        comments, by_id, by_para_id = self._comment_index()
        target = by_id.get(comment_id)
        if target is None:
            raise CommentNotFoundError(f"Comment {comment_id} not found")
        root = self._root_for(target, by_para_id)
        root_key = self._thread_key(root)
        return [
            comment
            for comment in comments
            if self._thread_key(self._root_for(comment, by_para_id)) == root_key
        ]

    def _collect_comment_para_ids(self) -> set[str]:
        para_ids: set[str] = set()
        for comment_elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            for para in comment_elem.findall(_qn(NS_W, "p")):
                para_id = para.get(_qn(NS_W14, "paraId"))
                if para_id:
                    para_ids.add(para_id)
        return para_ids

    def _cleanup_orphan_metadata(self, valid_para_ids: set[str]) -> None:
        ext_part = CommentsExtendedPart(self._document)
        ids_part = CommentsIdsPart(self._document)
        extensible_part = CommentsExtensiblePart(self._document)

        orphan_para_ids: set[str] = set()
        for elem in list(ext_part.xml):
            if etree.QName(elem).localname != "commentEx":
                continue
            para_id = elem.get(_qn(NS_W15, "paraId"))
            if para_id and para_id not in valid_para_ids:
                orphan_para_ids.add(para_id)

        for elem in list(ids_part.xml):
            if etree.QName(elem).localname != "commentId":
                continue
            para_id = elem.get(_qn(NS_W16CID, "paraId"))
            if para_id and para_id not in valid_para_ids:
                orphan_para_ids.add(para_id)

        removed_durable_ids: set[str] = set()
        for para_id in orphan_para_ids:
            ext_part.remove_comment_ex(para_id)
            durable_id = ids_part.remove_comment_id(para_id)
            if durable_id:
                removed_durable_ids.add(durable_id)

        for durable_id in removed_durable_ids:
            extensible_part.remove_comment_extensible(durable_id)

    def _detach_orphan_replies(self, valid_para_ids: set[str]) -> None:
        # Read the raw threading entries (not list_comments, which normalizes
        # dangling parents away) so the stale w15:paraIdParent attributes are
        # actually removed from the saved XML.
        ext_part = CommentsExtendedPart(self._document)
        for para_id, info in ext_part.get_threading_info().items():
            parent = info.get("parent_para_id")
            if parent and parent not in valid_para_ids:
                ext_part.set_parent(para_id, None)

    def migrate_comment_metadata(self) -> None:
        """
        Backfill missing comment metadata in existing documents.

        Ensures:
        - w14:paraId and w14:textId on comment paragraphs
        - commentsExtended.xml entries (commentEx)
        - commentsIds.xml entries (durableId)
        - commentsExtensible.xml entries (commentExtensible)

        Also removes metadata entries that no longer match any comment
        paragraph and detaches replies whose parent no longer exists.
        """
        ensure_comment_parts(self._document)

        ext_part = CommentsExtendedPart(self._document)
        ids_part = CommentsIdsPart(self._document)
        extensible_part = CommentsExtensiblePart(self._document)
        threading = ext_part.get_threading_info()
        durable_ids = ids_part.get_durable_ids()
        extensible_info = extensible_part.get_extensible_info()

        comment_elems = self._comments_xml.findall(_qn(NS_W, "comment"))
        updated_comments = False

        # When exactly one comment paragraph lost its w14:paraId and exactly
        # one metadata entry is unclaimed, reuse the old paraId — but only
        # when independent evidence (matching dateUtc) ties the entry to this
        # comment, so orphan metadata from a deleted comment cannot be
        # grafted onto an unrelated one.
        existing_para_ids: set[str] = set()
        missing_paras: list[etree._Element] = []
        for comment_elem in comment_elems:
            for para in comment_elem.findall(_qn(NS_W, "p")):
                pid = para.get(_qn(NS_W14, "paraId"))
                if pid:
                    existing_para_ids.add(pid)
                else:
                    missing_paras.append(para)
        unclaimed = {pid for pid in threading if pid not in existing_para_ids}
        unclaimed.update(pid for pid in durable_ids if pid not in existing_para_ids)
        if len(missing_paras) == 1 and len(unclaimed) == 1:
            candidate = unclaimed.pop()
            para = missing_paras[0]
            owner = para.getparent()
            while owner is not None and etree.QName(owner).localname != "comment":
                owner = owner.getparent()
            if owner is not None and self._para_id_reuse_corroborated(
                owner, candidate, durable_ids, extensible_info
            ):
                para.set(_qn(NS_W14, "paraId"), candidate)
                updated_comments = True

        used_hex_ids = self._used_long_hex_ids()

        for comment_elem in comment_elems:
            para_ids = []
            for para in comment_elem.findall(_qn(NS_W, "p")):
                para_id = para.get(_qn(NS_W14, "paraId"))
                if not para_id:
                    para_id = self._new_long_hex_id(used_hex_ids)
                    para.set(_qn(NS_W14, "paraId"), para_id)
                    updated_comments = True
                para_ids.append(para_id)

                text_id = para.get(_qn(NS_W14, "textId"))
                if not text_id:
                    text_id = self._new_long_hex_id(used_hex_ids)
                    para.set(_qn(NS_W14, "textId"), text_id)
                    updated_comments = True

            if not para_ids:
                continue

            primary_para_id = None
            for pid in reversed(para_ids):
                if pid in threading:
                    primary_para_id = pid
                    break
            if primary_para_id is None:
                for pid in reversed(para_ids):
                    if pid in durable_ids:
                        primary_para_id = pid
                        break
            if primary_para_id is None:
                primary_para_id = para_ids[-1]

            if primary_para_id not in threading:
                ext_part.add_comment_ex(
                    para_id=primary_para_id, parent_para_id=None, done=False
                )
                threading[primary_para_id] = {
                    "parent_para_id": None,
                    "done": False,
                }

            if primary_para_id not in durable_ids:
                durable_ids[primary_para_id] = self._new_long_hex_id(used_hex_ids)
                ids_part.add_comment_id(
                    para_id=primary_para_id,
                    durable_id=durable_ids[primary_para_id],
                )

            durable_id = durable_ids.get(primary_para_id)
            ext_entry = extensible_info.get(durable_id) if durable_id else None
            if durable_id and (
                durable_id not in extensible_info
                or not (ext_entry or {}).get("date_utc")
            ):
                date_str = comment_elem.get(_qn(NS_W, "date"))
                timestamp = _parse_comment_date(date_str)
                date_utc = _format_utc(timestamp) if timestamp else None
                extensible_part.add_comment_extensible(
                    durable_id=durable_id,
                    date_utc=date_utc,
                )

        if updated_comments:
            # Backfilled w14 attributes must be declared ignorable on roots
            # that do not already say so (e.g. foreign comments.xml parts).
            ensure_mc_ignorable(self._comments_xml)
            self._save_comments()

        # Repair leftovers: metadata keyed to paraIds that no longer match a
        # comment paragraph, and reply links to parents that no longer exist.
        valid_para_ids = self._collect_comment_para_ids()
        self._cleanup_orphan_metadata(valid_para_ids)
        self._detach_orphan_replies(valid_para_ids)

    @staticmethod
    def _para_id_reuse_corroborated(
        comment_elem: etree._Element,
        para_id: str,
        durable_ids: dict[str, str],
        extensible_info: dict[str, dict],
    ) -> bool:
        """Check that an unclaimed paraId's metadata belongs to this comment.

        The entry's durable id must carry a dateUtc equal to the comment's
        own w:date; cardinality alone (one missing paragraph, one unclaimed
        entry) does not prove correspondence.
        """
        durable_id = durable_ids.get(para_id)
        if not durable_id:
            return False
        date_utc = (extensible_info.get(durable_id) or {}).get("date_utc")
        if not date_utc:
            return False
        timestamp = _parse_comment_date(comment_elem.get(_qn(NS_W, "date")))
        return timestamp is not None and _format_utc(timestamp) == date_utc

    @staticmethod
    def _comment_text(comment_elem: etree._Element) -> str:
        """Extract a comment's text, preserving breaks, tabs, and paragraphs."""

        def has_ancestor(elem: etree._Element, tag: str, stop: etree._Element) -> bool:
            parent = elem.getparent()
            while parent is not None and parent is not stop:
                if parent.tag == tag:
                    return True
                parent = parent.getparent()
            return False

        paragraphs: list[str] = []
        for para in comment_elem.iter(_qn(NS_W, "p")):
            # Paragraphs nested inside another paragraph (text boxes) are
            # covered by the outer paragraph's run walk; visiting them here
            # too would double-count their text.
            if has_ancestor(para, _qn(NS_W, "p"), comment_elem):
                continue
            # Block-level mc:AlternateContent duplicates whole paragraphs in
            # its Fallback branch; extract only the Choice copy.
            if has_ancestor(para, _qn(NS_MC, "Fallback"), comment_elem):
                continue
            pieces: list[str] = []
            for run in para.iter(_qn(NS_W, "r")):
                # mc:AlternateContent carries the same content twice
                # (mc:Choice and mc:Fallback); extract only the Choice copy.
                if has_ancestor(run, _qn(NS_MC, "Fallback"), para):
                    continue
                for child in run:
                    localname = etree.QName(child).localname
                    if localname == "t":
                        if child.text:
                            pieces.append(child.text)
                    elif localname in ("br", "cr"):
                        pieces.append("\n")
                    elif localname == "tab":
                        pieces.append("\t")
            paragraphs.append("".join(pieces))
        return "\n".join(paragraphs)

    def list_comments(self) -> Iterator[CommentInfo]:
        """
        List all comments in the document.

        Yields:
            CommentInfo objects for each comment.
        """
        # Collect comments from comments.xml
        comments_data: list[dict] = []

        for comment_elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            comment_id = comment_elem.get(_qn(NS_W, "id"))
            author = comment_elem.get(_qn(NS_W, "author"), "")
            initials = comment_elem.get(_qn(NS_W, "initials"))
            date_str = comment_elem.get(_qn(NS_W, "date"))

            text = self._comment_text(comment_elem)

            # Collect paraIds from all comment paragraphs (some comments span multiple paragraphs)
            para_ids = []
            for para in comment_elem.findall(_qn(NS_W, "p")):
                para_id = para.get(_qn(NS_W14, "paraId"))
                if para_id:
                    para_ids.append(para_id)

            # Parse timestamp (OOXML uses UTC, normalize all to tz-aware)
            timestamp = _parse_comment_date(date_str)

            comments_data.append(
                {
                    "comment_id": comment_id,
                    "para_ids": para_ids,
                    "text": text,
                    "author": author,
                    "initials": initials,
                    "timestamp": timestamp,
                }
            )

        # Get threading info from commentsExtended.xml
        ext_part = CommentsExtendedPart(self._document)
        threading = ext_part.get_threading_info()

        # Get durable IDs from commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        durable_ids = ids_part.get_durable_ids()

        known_para_ids: set[str] = set()
        for info in comments_data:
            known_para_ids.update(info["para_ids"])

        # Build CommentInfo objects
        for info in comments_data:
            # "" when no paraId is identifiable (never a key in the satellite
            # parts, so the lookups below fall through to their defaults).
            para_id = (
                self._primary_para_id(info["para_ids"], threading, durable_ids) or ""
            )

            thread_info = threading.get(para_id, {})
            parent_para_id = thread_info.get("parent_para_id")
            if parent_para_id and parent_para_id not in known_para_ids:
                # Dangling link (parent removed by another tool): treat the
                # comment as the thread root it effectively is.
                parent_para_id = None
            yield CommentInfo(
                comment_id=info["comment_id"],
                para_id=para_id,
                text=info["text"],
                author=info["author"],
                initials=info["initials"],
                timestamp=info["timestamp"],
                parent_para_id=parent_para_id,
                is_resolved=thread_info.get("done", False),
                durable_id=durable_ids.get(para_id),
            )

    def get_comment_threads(self) -> list[CommentThread]:
        """
        Get all comment threads (grouped by root comment).

        Returns:
            List of CommentThread objects.
        """
        comments = list(self.list_comments())

        # Index comments by para_id for parent traversal
        by_para_id = {c.para_id: c for c in comments if c.para_id}

        # Build threads by walking parent chains (supports reply-to-reply)
        threads_by_root: dict[str, CommentThread] = {}
        for comment in comments:
            root = self._root_for(comment, by_para_id)
            root_key = self._thread_key(root)
            thread = threads_by_root.get(root_key)
            if thread is None:
                thread = CommentThread(root=root, replies=[])
                threads_by_root[root_key] = thread

            if comment is not root:
                thread.replies.append(comment)

        # Sort replies by timestamp (use tz-aware min for comparison)
        min_dt = datetime.min.replace(tzinfo=timezone.utc)
        for thread in threads_by_root.values():
            thread.replies.sort(key=lambda c: c.timestamp or min_dt)

        return list(threads_by_root.values())

    def get_comment(self, comment_id: Union[int, str]) -> CommentInfo:
        """Return the CommentInfo for a single comment.

        Raises:
            CommentNotFoundError: If no comment has this id.
        """
        comment_id = _coerce_comment_id(comment_id)
        _, by_id, _ = self._comment_index()
        info = by_id.get(comment_id)
        if info is None:
            raise CommentNotFoundError(f"Comment {comment_id} not found")
        return info

    def get_thread(self, comment_id: Union[int, str]) -> CommentThread:
        """Return the thread containing a comment (root or reply).

        Raises:
            CommentNotFoundError: If no comment has this id.
        """
        comment_id = _coerce_comment_id(comment_id)
        for thread in self.get_comment_threads():
            if any(c.comment_id == comment_id for c in thread.all_comments):
                return thread
        raise CommentNotFoundError(f"Comment {comment_id} not found")

    def get_comment_paragraph(
        self, comment_id: Union[int, str]
    ) -> Optional[Paragraph]:
        """Paragraph containing the comment's anchor.

        Returns None when the comment exists but has no range anchors.

        Raises:
            CommentNotFoundError: If no comment has this id.
        """
        comment_id = _coerce_comment_id(comment_id)
        if not self._comment_id_exists(comment_id):
            raise CommentNotFoundError(f"Comment {comment_id} not found")
        return CommentAnchor(self._document).find_paragraph_with_comment(comment_id)

    def get_anchored_text(self, comment_id: Union[int, str]) -> Optional[str]:
        """The document text the comment is anchored to.

        Returns None when the comment has no commentRangeStart/End pair
        (reference-only or anchor-less comments).

        Raises:
            CommentNotFoundError: If no comment has this id.
        """
        comment_id = _coerce_comment_id(comment_id)
        if not self._comment_id_exists(comment_id):
            raise CommentNotFoundError(f"Comment {comment_id} not found")
        return CommentAnchor(self._document).get_anchored_text(comment_id)

    def get_authors(self) -> dict[str, str]:
        """
        Get all unique authors who have commented on this document.

        Returns:
            Dict mapping author name to initials, e.g. {"Sun, Ting": "ST"}
        """
        authors: dict[str, str] = {}
        for comment in self.list_comments():
            if not comment.author:
                continue
            if comment.author not in authors:
                authors[comment.author] = comment.initials or ""
            elif not authors[comment.author] and comment.initials:
                # Prefer first non-empty initials when available
                authors[comment.author] = comment.initials
        return authors

    def get_document_author(self) -> tuple[str, Optional[str]]:
        """
        Get the document owner's name and initials.

        Uses document core properties for the author name, then looks up
        initials from existing comments by that author.

        Returns:
            Tuple of (author_name, initials). author_name is always a string
            but may be empty ("") if no author is set in document properties.
            Initials may be None if the document owner hasn't made any comments.
        """
        author = self._document.core_properties.author or ""
        if not author:
            # Fallback to last_modified_by
            author = self._document.core_properties.last_modified_by or ""

        # Look for initials in existing comments
        initials = None
        for comment in self.list_comments():
            if comment.author == author and comment.initials:
                initials = comment.initials
                break

        return author, initials

    def get_people(self) -> list[PersonInfo]:
        """
        List people entries from word/people.xml.

        Returns:
            List of PersonInfo entries. Empty if people.xml is absent.
        """
        people_part = PeoplePart(self._document)
        return people_part.get_people()

    def get_person(self, author: str) -> PersonInfo:
        """
        Get a single person entry by author name.

        Args:
            author: Author name to look up in people.xml.

        Returns:
            PersonInfo if found.

        Raises:
            KeyError: If no matching person is found.
        """
        people_part = PeoplePart(self._document)
        return people_part.get_person(author)

    def ensure_person(
        self, author: str, presence: Optional[dict[str, str]] = None
    ) -> PersonInfo:
        """
        Ensure a people.xml entry exists for an author.

        Args:
            author: Author name to match w:comment/@w:author.
            presence: Optional presence metadata with provider_id/user_id.

        Returns:
            PersonInfo for the ensured entry.
        """
        people_part = PeoplePart(self._document)
        return people_part.ensure_person(author, presence)

    def _parse_author_spec(self, author: PersonInfo) -> tuple[str, Optional[dict[str, str]]]:
        if not isinstance(author, PersonInfo):
            raise TypeError("author must be a PersonInfo")

        author_name = author.author
        if not author_name:
            raise ValueError("author must be non-empty")

        presence = None
        if author.provider_id and author.user_id:
            presence = {
                "provider_id": author.provider_id,
                "user_id": author.user_id,
            }
        elif author.provider_id or author.user_id:
            raise ValueError("author presence must include provider_id and user_id")

        return author_name, presence

    _ALLOWED_RUN_FORMATS = frozenset({"bold", "italic", "underline"})

    def _normalize_content(
        self, content: CommentContent
    ) -> list[list[tuple[str, dict]]]:
        """Normalize comment content to paragraphs of (text, format) runs.

        Validates types, format keys, and XML-legality up front so callers
        can raise before mutating anything.
        """
        if isinstance(content, str):
            _validate_xml_text(content, "comment text")
            return [[(content, {})]]
        paragraphs: list[list[tuple[str, dict]]] = []
        for para_spec in content:
            if isinstance(para_spec, str):
                _validate_xml_text(para_spec, "comment text")
                paragraphs.append([(para_spec, {})])
                continue
            runs: list[tuple[str, dict]] = []
            for run_spec in para_spec:
                fmt: dict = {}
                if isinstance(run_spec, str):
                    run_text = run_spec
                else:
                    run_text, fmt = run_spec
                    if not isinstance(run_text, str) or not isinstance(fmt, dict):
                        raise TypeError(
                            "run specs must be str or (str, dict) tuples"
                        )
                    unknown = set(fmt) - self._ALLOWED_RUN_FORMATS
                    if unknown:
                        raise ValueError(
                            f"unsupported run formatting keys: {sorted(unknown)}"
                        )
                    fmt = dict(fmt)
                _validate_xml_text(run_text, "comment text")
                runs.append((run_text, fmt))
            paragraphs.append(runs)
        if not paragraphs:
            raise ValueError("comment content must have at least one paragraph")
        return paragraphs

    def _get_default_author_person(
        self,
        docx_path: Optional[str] = None,
        include_presence: bool = False,
        strict_docx: bool = False,
    ) -> tuple[PersonInfo, Optional[str]]:
        """
        Internal helper to resolve a default author PersonInfo.

        Preference order:
        1) DOCX file from `docx_path` or env var DOCX_COMMENTS_AUTHOR_DOCX
        2) System Office user info (macOS plist / Windows registry)
        3) Current document core properties

        Returns:
            (PersonInfo, initials)

        Raises:
            ValueError: If no author can be resolved.
        """
        person, initials = _default_person_from_system(
            docx_path=docx_path,
            include_presence=include_presence,
            strict_docx=strict_docx,
        )
        if person:
            return person, initials

        if strict_docx and (docx_path or os.environ.get("DOCX_COMMENTS_AUTHOR_DOCX")):
            raise ValueError("default author DOCX did not yield an author")

        author_name = self._document.core_properties.author or ""
        if not author_name:
            author_name = self._document.core_properties.last_modified_by or ""
        if author_name:
            return PersonInfo(author=author_name), None

        raise ValueError("no default author could be resolved")

    def get_default_author_person(
        self,
        docx_path: Optional[str] = None,
        include_presence: bool = False,
        strict_docx: bool = False,
    ) -> tuple[PersonInfo, Optional[str]]:
        """
        Resolve a default author PersonInfo.

        Args:
            docx_path: Optional path to a DOCX file used as the author source.
            include_presence: Whether to include presence metadata from people.xml.
            strict_docx: If True and a DOCX source is provided (or env var set),
                raise when the DOCX cannot provide an author, without falling back.
                A DOCX with multiple people entries triggers a warning and falls back.

        Returns:
            (PersonInfo, initials)
        """
        return self._get_default_author_person(
            docx_path=docx_path,
            include_presence=include_presence,
            strict_docx=strict_docx,
        )

    def merge_people_from(
        self, source: Document, include_presence: bool = False
    ) -> list[PersonInfo]:
        """
        Merge people entries from another document.

        Args:
            source: Document to import people.xml entries from.
            include_presence: Whether to copy presence metadata.

        Returns:
            List of PersonInfo entries added to this document.
        """
        source_part = PeoplePart(source)
        target_part = PeoplePart(self._document)
        return target_part.merge_from(source_part, include_presence)

    def _resolve_person_spec(
        self,
        author: str,
        person: Optional[PersonSpec],
    ) -> Optional[tuple[str, Optional[dict[str, str]]]]:
        """Validate a person spec without mutating the document.

        Returns the (author, presence) pair to ensure in people.xml, or None
        when no entry should be created. All validation errors are raised
        here so callers can validate before touching the package.
        """
        if person is None or person is False:
            return None

        if isinstance(person, bool):
            return (author, None)

        presence: Optional[dict[str, str]] = None
        person_author = author

        if isinstance(person, PersonInfo):
            person_author = person.author
            if person.provider_id and person.user_id:
                presence = {
                    "provider_id": person.provider_id,
                    "user_id": person.user_id,
                }
            elif person.provider_id or person.user_id:
                raise ValueError("presence must include provider_id and user_id")
        elif isinstance(person, str):
            person_author = person
        elif isinstance(person, dict):
            if "author" in person and isinstance(person["author"], str):
                person_author = person["author"]
            raw_presence = person.get("presence")
            if isinstance(raw_presence, dict):
                presence = raw_presence  # type: ignore[assignment]
            else:
                provider_id = person.get("provider_id") or person.get("providerId")
                user_id = person.get("user_id") or person.get("userId")
                if provider_id and user_id:
                    presence = {
                        "provider_id": str(provider_id),
                        "user_id": str(user_id),
                    }
                elif provider_id or user_id:
                    raise ValueError("presence must include provider_id and user_id")
        else:
            raise TypeError("person must be a bool, str, dict, or PersonInfo")

        if person_author != author:
            raise ValueError("person author must match comment author to link identity")

        _validate_xml_text(person_author, "person author")
        if presence is not None:
            provider_id, user_id = PeoplePart._normalize_presence(presence)
            _validate_xml_text(provider_id, "presence provider_id")
            _validate_xml_text(user_id, "presence user_id")

        return (person_author, presence)

    def _apply_person_spec(
        self, plan: Optional[tuple[str, Optional[dict[str, str]]]]
    ) -> None:
        if plan is not None:
            self.ensure_person(plan[0], plan[1])

    def add_comment(
        self,
        paragraph: Paragraph,
        text: CommentContent,
        author: PersonInfo,
        initials: Optional[str] = None,
        start_run: int = 0,
        end_run: Optional[int] = None,
        person: Optional[PersonSpec] = None,
        timestamp: Optional[datetime] = None,
    ) -> str:
        """
        Add a new anchored comment to a paragraph.

        Args:
            paragraph: The paragraph to comment on (must belong to this
                manager's document).
            text: Comment content. A plain str is one paragraph; newlines
                and tabs are preserved (encoded as w:br / w:tab so Word
                renders them). A sequence of paragraphs is also accepted,
                each a str or a sequence of runs, where a run is a str or
                (text, format) with format keys "bold"/"italic"/"underline",
                e.g. ``[[("urgent", {"bold": True}), " please fix"], "thanks"]``.
            author: PersonInfo instance.
            initials: Author initials (optional).
            start_run: Index of first run to anchor (default: 0). Python-style
                negative indices are accepted; out-of-range indices raise.
            end_run: Index of last run to anchor (default: last run).
            person: Optional people.xml entry to link author identity.
                Accepts True (ensure an entry for the comment author), a str
                author name or PersonInfo (must match the comment author), or
                a dict with optional "author" and presence keys
                ("provider_id"/"user_id" or a "presence" dict). None/False
                leave people.xml untouched.
            timestamp: Optional creation time. Naive datetimes are
                interpreted as local time; None uses the current time.

        Returns:
            The comment ID of the new comment.

        Raises:
            ValueError: If the paragraph belongs to another document, the
                text contains characters not allowed in XML, end_run precedes
                start_run, or the person spec is invalid.
            IndexError: If start_run/end_run are out of range.
        """
        # Validate everything before mutating anything.
        author_name, author_presence = self._parse_author_spec(author)
        content = self._normalize_content(text)
        _validate_xml_text(author_name, "author")
        _validate_xml_text(initials, "initials")

        anchor = CommentAnchor(self._document)
        anchor.validate_anchor_target(paragraph, start_run, end_run)

        person_spec = person
        if person_spec is None and author_presence:
            person_spec = {"presence": author_presence}
        elif person_spec is True and author_presence:
            person_spec = {"presence": author_presence}
        person_plan = self._resolve_person_spec(author_name, person_spec)

        self._ensure_parts()

        comment_id = self._new_comment_id()
        used_hex_ids = self._used_long_hex_ids()
        para_id = self._new_long_hex_id(used_hex_ids)
        text_id = self._new_long_hex_id(used_hex_ids)
        durable_id = self._new_long_hex_id(used_hex_ids)

        self._apply_person_spec(person_plan)

        # 1. Add to comments.xml
        timestamp = self._add_comment_xml(
            comment_id=comment_id,
            para_id=para_id,
            text_id=text_id,
            content=content,
            author=author_name,
            initials=initials,
            timestamp=timestamp,
            used_hex_ids=used_hex_ids,
        )

        # 2. Add anchors to document.xml
        anchor.add_anchors(
            paragraph=paragraph,
            comment_id=comment_id,
            start_run=start_run,
            end_run=end_run,
        )

        # 3. Add to commentsExtended.xml (root comment, no parent)
        ext_part = CommentsExtendedPart(self._document)
        ext_part.add_comment_ex(para_id=para_id, parent_para_id=None, done=False)

        # 4. Add to commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        ids_part.add_comment_id(para_id=para_id, durable_id=durable_id)

        # 5. Add to commentsExtensible.xml (modern comments metadata)
        extensible_part = CommentsExtensiblePart(self._document)
        extensible_part.add_comment_extensible(
            durable_id=durable_id,
            date_utc=_format_utc(timestamp),
        )

        return comment_id

    def reply_to_comment(
        self,
        parent_id: Union[int, str],
        text: CommentContent,
        author: PersonInfo,
        initials: Optional[str] = None,
        person: Optional[PersonSpec] = None,
        timestamp: Optional[datetime] = None,
    ) -> str:
        """
        Reply to an existing comment.

        Args:
            parent_id: Comment ID of the parent comment.
            text: Reply content. Newlines and tabs are preserved. Accepts
                the same forms as add_comment: a plain str, or a sequence
                of paragraphs of (text, format) runs, e.g.
                ``[[("agreed", {"italic": True})]]``.
            author: PersonInfo instance.
            initials: Author initials (optional).
            person: Optional people.xml entry to link author identity (see
                add_comment for the accepted forms).
            timestamp: Optional creation time. Naive datetimes are
                interpreted as local time; None uses the current time.

        Returns:
            The comment ID of the reply.

        Raises:
            CommentNotFoundError: If the parent comment is not found.
            TypeError: If author is not a PersonInfo (or str once Task 12
                lands) or parent_id is neither str nor int.
            ValueError: If the text/author/initials contain characters not
                allowed in XML, the person spec is invalid, or the parent
                comment has no anchors in the document.
        """
        parent_id = _coerce_comment_id(parent_id)
        author_name, author_presence = self._parse_author_spec(author)
        content = self._normalize_content(text)
        _validate_xml_text(author_name, "author")
        _validate_xml_text(initials, "initials")

        person_spec = person
        if person_spec is None and author_presence:
            person_spec = {"presence": author_presence}
        elif person_spec is True and author_presence:
            person_spec = {"presence": author_presence}
        person_plan = self._resolve_person_spec(author_name, person_spec)

        # Validate the parent id before mutating anything (including the
        # metadata migration below).
        if not self._comment_id_exists(parent_id):
            raise CommentNotFoundError(f"Parent comment {parent_id} not found")

        self._ensure_parts()

        # Find parent comment's para_id and resolve root for compatibility.
        comments = list(self.list_comments())
        parent_comment = next((c for c in comments if c.comment_id == parent_id), None)

        if parent_comment is None or not parent_comment.para_id:
            # The parent exists but lacks metadata (e.g. created by another
            # tool); backfill it and retry.
            self.migrate_comment_metadata()
            comments = list(self.list_comments())
            parent_comment = next((c for c in comments if c.comment_id == parent_id), None)
            if parent_comment is None or not parent_comment.para_id:
                raise CommentNotFoundError(f"Parent comment {parent_id} not found")

        parent_para_id = parent_comment.para_id
        parent_parent_para_id = parent_comment.parent_para_id

        by_para_id = {c.para_id: c for c in comments if c.para_id}
        root_comment = self._root_for(parent_comment, by_para_id)

        # Word UI doesn't support nested replies; attach to the root comment.
        effective_parent_para_id = root_comment.para_id or parent_para_id
        effective_parent_parent_para_id = root_comment.parent_para_id

        anchor = CommentAnchor(self._document)
        anchor_parent_id = root_comment.comment_id or parent_id

        # Validate the anchor location before writing the reply so a failure
        # cannot leave an anchor-less comment behind. A reference run alone is
        # a legal anchor (range markers are optional per ECMA-376 §17.13.4).
        _, parent_start, parent_end, parent_ref = anchor._find_anchor_elements(
            anchor_parent_id
        )
        if (parent_start is None or parent_end is None) and parent_ref is None:
            raise ValueError(f"Could not find anchors for comment {anchor_parent_id}")

        comment_id = self._new_comment_id()
        used_hex_ids = self._used_long_hex_ids()
        para_id = self._new_long_hex_id(used_hex_ids)
        text_id = self._new_long_hex_id(used_hex_ids)
        durable_id = self._new_long_hex_id(used_hex_ids)

        self._apply_person_spec(person_plan)

        # 1. Add to comments.xml
        timestamp = self._add_comment_xml(
            comment_id=comment_id,
            para_id=para_id,
            text_id=text_id,
            content=content,
            author=author_name,
            initials=initials,
            timestamp=timestamp,
            used_hex_ids=used_hex_ids,
        )

        # 2. Add anchors at the root comment location for Word threading compatibility.
        anchor.add_anchors_at_comment(
            parent_comment_id=anchor_parent_id,
            new_comment_id=comment_id,
        )

        # 3. Ensure parent exists in commentsExtended.xml, then add reply link
        ext_part = CommentsExtendedPart(self._document)
        threading = ext_part.get_threading_info()
        if parent_para_id not in threading:
            ext_part.add_comment_ex(
                para_id=parent_para_id,
                parent_para_id=parent_parent_para_id,
                done=False,
            )
        if effective_parent_para_id not in threading and effective_parent_para_id != parent_para_id:
            ext_part.add_comment_ex(
                para_id=effective_parent_para_id,
                parent_para_id=effective_parent_parent_para_id,
                done=False,
            )
        # The reply inherits the thread's current resolution state so a
        # resolved thread stays consistently resolved (Word keeps all
        # members' done flags in sync).
        inherited_done = threading.get(effective_parent_para_id, {}).get("done", False)
        ext_part.add_comment_ex(
            para_id=para_id,
            parent_para_id=effective_parent_para_id,
            done=inherited_done,
        )

        # 4. Add to commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        ids_part.add_comment_id(para_id=para_id, durable_id=durable_id)

        # 5. Add to commentsExtensible.xml (modern comments metadata)
        extensible_part = CommentsExtensiblePart(self._document)
        extensible_part.add_comment_extensible(
            durable_id=durable_id,
            date_utc=_format_utc(timestamp),
        )

        return comment_id

    def resolve_comment(self, comment_id: Union[int, str]) -> None:
        """
        Mark a comment's thread as resolved.

        Resolution is thread-scoped, matching Word: every comment in the
        thread is marked done.

        Args:
            comment_id: Any comment ID within the thread.

        Raises:
            ValueError: If comment not found.
        """
        self.set_comment_resolved(comment_id, True)

    def unresolve_comment(self, comment_id: Union[int, str]) -> None:
        """
        Mark a comment's thread as unresolved (thread-scoped, like Word).

        Args:
            comment_id: Any comment ID within the thread.

        Raises:
            ValueError: If comment not found.
        """
        self.set_comment_resolved(comment_id, False)

    def set_comment_resolved(self, comment_id: Union[int, str], resolved: bool) -> None:
        """
        Set the resolved status for a comment's thread.

        Word treats resolution as thread-scoped and marks every member done,
        so this updates the whole thread. Missing metadata (e.g. comments
        created by other tools) is backfilled first.

        Args:
            comment_id: Any comment ID within the thread.
            resolved: True to resolve, False to unresolve.

        Raises:
            ValueError: If comment not found.
        """
        comment_id = _coerce_comment_id(comment_id)
        if not self._comment_id_exists(comment_id):
            raise CommentNotFoundError(f"Comment {comment_id} not found")

        self.migrate_comment_metadata()

        thread_comments = self._thread_comments_for(comment_id)
        ext_part = CommentsExtendedPart(self._document)
        threading = ext_part.get_threading_info()
        for comment in thread_comments:
            if comment.para_id and comment.para_id in threading:
                ext_part.set_done(comment.para_id, done=resolved)

    def delete_comment(self, comment_id: Union[int, str]) -> None:
        """
        Delete a single comment.

        Replies remain in the document but are detached from the deleted
        parent (their anchors stay at the original location).

        Args:
            comment_id: The comment ID to delete.

        Raises:
            ValueError: If comment not found (checked before any mutation).
        """
        comment_id = _coerce_comment_id(comment_id)
        if not self._comment_id_exists(comment_id):
            raise CommentNotFoundError(f"Comment {comment_id} not found")

        self.migrate_comment_metadata()

        # Remove anchors first so a failure cannot leave anchors referencing
        # a comment that no longer exists.
        anchor = CommentAnchor(self._document)
        anchor.remove_anchors(comment_id)

        removed_para_ids = self._comments_part().remove_comment(comment_id) or []

        # Remove comment metadata entries.
        deleted_para_ids = {pid for pid in removed_para_ids if pid}
        self._cleanup_comment_metadata(deleted_para_ids)

        remaining_para_ids = self._collect_comment_para_ids()
        self._cleanup_orphan_metadata(remaining_para_ids)
        self._detach_orphan_replies(remaining_para_ids)

    def delete_thread(self, comment_id: Union[int, str]) -> None:
        """
        Delete an entire comment thread (root + replies).

        Args:
            comment_id: Any comment ID within the thread.

        Raises:
            ValueError: If comment not found (checked before any mutation).
        """
        comment_id = _coerce_comment_id(comment_id)
        if not self._comment_id_exists(comment_id):
            raise CommentNotFoundError(f"Comment {comment_id} not found")

        self.migrate_comment_metadata()
        thread_comments = self._thread_comments_for(comment_id)

        handler = self._comments_part()
        anchor = CommentAnchor(self._document)
        deleted_para_ids: set[str] = set()
        seen_ids: set[str] = set()

        try:
            for comment in thread_comments:
                cid = comment.comment_id
                if not cid or cid in seen_ids:
                    continue
                seen_ids.add(cid)
                anchor.remove_anchors(cid)
                removed_para_ids = handler.remove_comment(cid)
                if removed_para_ids:
                    deleted_para_ids.update(pid for pid in removed_para_ids if pid)
        finally:
            # Always purge metadata for whatever was removed, even if a
            # malformed thread member interrupted the loop.
            self._cleanup_comment_metadata(deleted_para_ids)
            remaining_para_ids = self._collect_comment_para_ids()
            self._cleanup_orphan_metadata(remaining_para_ids)
            self._detach_orphan_replies(remaining_para_ids)

    def edit_comment(
        self,
        comment_id: Union[int, str],
        text: CommentContent,
        author: Optional[str] = None,
        initials: Optional[str] = None,
        timestamp: Optional[datetime] = None,
    ) -> None:
        """Replace a comment's text in place.

        The comment id, primary paraId, durable id, threading (replies),
        resolution state, and document anchors are all preserved; only the
        content (and optionally author/initials/date) changes. w14:textId
        is refreshed, matching Word's text-revision semantics.

        Args:
            comment_id: The comment to edit.
            text: New comment content (same forms as add_comment): a plain
                str, or a sequence of paragraphs of (text, format) runs,
                e.g. ``["revised:", [("keep", {"bold": True}), " as is"]]``.
            author: Optional new author name.
            initials: Optional new initials.
            timestamp: Optional new date (naive = local time); also updates
                the commentsExtensible dateUtc entry.

        Raises:
            CommentNotFoundError: If no comment has this id.
            ValueError: If text/author/initials contain characters not
                allowed in XML, or author is empty.
        """
        comment_id = _coerce_comment_id(comment_id)
        content = self._normalize_content(text)
        if author is not None:
            if not author:
                raise ValueError("author must be non-empty")
            _validate_xml_text(author, "author")
        if initials is not None:
            _validate_xml_text(initials, "initials")
        if timestamp is not None and timestamp.tzinfo is None:
            timestamp = timestamp.astimezone()

        comment_elem = None
        for elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            if elem.get(_qn(NS_W, "id")) == comment_id:
                comment_elem = elem
                break
        if comment_elem is None:
            raise CommentNotFoundError(f"Comment {comment_id} not found")

        threading = CommentsExtendedPart(self._document).get_threading_info()
        durable_ids = CommentsIdsPart(self._document).get_durable_ids()
        para_ids = [
            para.get(_qn(NS_W14, "paraId"))
            for para in comment_elem.findall(_qn(NS_W, "p"))
            if para.get(_qn(NS_W14, "paraId"))
        ]
        primary = self._primary_para_id(para_ids, threading, durable_ids)

        used_hex_ids = self._used_long_hex_ids()
        if primary is None:
            primary = self._new_long_hex_id(used_hex_ids)
        text_id = self._new_long_hex_id(used_hex_ids)

        rsid_r = uuid.uuid4().hex[:8].upper()
        rsid_default = uuid.uuid4().hex[:8].upper()
        rsid_rpr = uuid.uuid4().hex[:8].upper()

        new_paras = []
        for index, runs in enumerate(content):
            is_last = index == len(content) - 1
            # The primary paraId (and fresh textId) stay on the LAST
            # paragraph, where Word keys the satellite metadata.
            new_paras.append(
                self._build_comment_paragraph(
                    runs,
                    para_id=primary if is_last else self._new_long_hex_id(used_hex_ids),
                    text_id=text_id if is_last else self._new_long_hex_id(used_hex_ids),
                    rsid_r=rsid_r,
                    rsid_default=rsid_default,
                    rsid_rpr=rsid_rpr,
                    include_annotation_ref=index == 0,
                )
            )

        for para in comment_elem.findall(_qn(NS_W, "p")):
            comment_elem.remove(para)
        for new_para in new_paras:
            comment_elem.append(new_para)

        if author is not None:
            comment_elem.set(_qn(NS_W, "author"), author)
        if initials is not None:
            comment_elem.set(_qn(NS_W, "initials"), initials)
        if timestamp is not None:
            comment_elem.set(
                _qn(NS_W, "date"), timestamp.isoformat(timespec="seconds")
            )
            durable = durable_ids.get(primary)
            if durable:
                CommentsExtensiblePart(self._document).set_date_utc(
                    durable, _format_utc(timestamp)
                )

        # Only the primary paraId survives the rebuild (non-last paragraphs
        # get fresh ids); metadata keyed to any other old paraId is orphaned.
        dropped = {pid for pid in para_ids if pid != primary}
        self._cleanup_comment_metadata(dropped)

        self._save_comments()

    def move_comment(
        self,
        comment_id: Union[int, str],
        paragraph: Paragraph,
        start_run: int = 0,
        end_run: Optional[int] = None,
    ) -> None:
        """
        Move a standalone comment's anchor to a new paragraph.

        Word keeps a thread's anchors co-located, so comments that are part
        of a thread with replies must be moved with move_thread() instead.

        Args:
            comment_id: The comment ID to move.
            paragraph: Paragraph to anchor the comment to (must belong to
                this manager's document).
            start_run: Index of first run to anchor.
            end_run: Index of last run to anchor (default: last run).

        Raises:
            ValueError: If comment not found, the comment belongs to a thread
                with replies, or the paragraph/run indices are invalid.
            IndexError: If start_run/end_run are out of range.
        """
        comment_id = _coerce_comment_id(comment_id)
        _, by_id, _ = self._comment_index()
        if comment_id not in by_id:
            raise CommentNotFoundError(f"Comment {comment_id} not found")

        thread_comments = self._thread_comments_for(comment_id)
        if len({c.comment_id for c in thread_comments if c.comment_id}) > 1:
            raise ValueError(
                f"Comment {comment_id} belongs to a thread with replies; "
                "use move_thread() to keep the thread's anchors co-located"
            )

        anchor = CommentAnchor(self._document)
        # Resolve the target span before removing the old anchors: removal
        # can delete reference runs and shift the run indices the caller
        # computed. The whole-paragraph default is index-free, so it is
        # re-resolved after removal instead (an in-place move would otherwise
        # pin the comment's own reference run as an endpoint).
        default_span = start_run == 0 and end_run is None
        span = anchor.plan_anchor_span(paragraph, start_run, end_run)
        if not default_span:
            anchor.ensure_span_survives_removal(span, {comment_id})
        anchor.remove_anchors(comment_id)
        if default_span:
            span = anchor.plan_anchor_span(paragraph, start_run, end_run)
        anchor.add_anchors_at_span(paragraph, span, comment_id)

    def move_thread(
        self,
        comment_id: Union[int, str],
        paragraph: Paragraph,
        start_run: int = 0,
        end_run: Optional[int] = None,
    ) -> None:
        """
        Move an entire comment thread (root + replies) to a new paragraph.

        Args:
            comment_id: Any comment ID within the thread.
            paragraph: Paragraph to anchor the thread to (must belong to this
                manager's document).
            start_run: Index of first run to anchor (root comment).
            end_run: Index of last run to anchor (root comment).

        Raises:
            ValueError: If comment not found, or the paragraph/run indices
                are invalid.
            IndexError: If start_run/end_run are out of range.
        """
        comment_id = _coerce_comment_id(comment_id)
        thread_comments = self._thread_comments_for(comment_id)
        by_para_id = {c.para_id: c for c in thread_comments if c.para_id}
        target = next(
            (comment for comment in thread_comments if comment.comment_id == comment_id),
            None,
        )
        if target is None:
            raise CommentNotFoundError(f"Comment {comment_id} not found")
        root = self._root_for(target, by_para_id)
        if not root.comment_id:
            raise ValueError(
                f"thread containing comment {comment_id} has a root without a "
                "w:id; the thread cannot be moved"
            )

        anchor = CommentAnchor(self._document)
        # Resolve the target span before removing the old anchors (removal
        # can delete reference runs and shift run indices); the index-free
        # whole-paragraph default is re-resolved after removal instead.
        default_span = start_run == 0 and end_run is None
        span = anchor.plan_anchor_span(paragraph, start_run, end_run)
        thread_ids = {c.comment_id for c in thread_comments if c.comment_id}
        if not default_span:
            anchor.ensure_span_survives_removal(span, thread_ids)

        seen_ids: set[str] = set()
        for comment in thread_comments:
            cid = comment.comment_id
            if not cid or cid in seen_ids:
                continue
            seen_ids.add(cid)
            anchor.remove_anchors(cid)

        if default_span:
            span = anchor.plan_anchor_span(paragraph, start_run, end_run)
        anchor.add_anchors_at_span(paragraph, span, root.comment_id)

        # Re-anchor replies at the root comment location.
        re_anchored = {root.comment_id}
        for comment in thread_comments:
            cid = comment.comment_id
            if not cid or cid in re_anchored:
                continue
            re_anchored.add(cid)
            anchor.add_anchors_at_comment(
                parent_comment_id=root.comment_id,
                new_comment_id=cid,
            )

    def _cleanup_comment_metadata(self, para_ids: set[str]) -> None:
        if not para_ids:
            return

        ext_part = CommentsExtendedPart(self._document)
        ids_part = CommentsIdsPart(self._document)
        extensible_part = CommentsExtensiblePart(self._document)
        durable_ids = ids_part.get_durable_ids()

        for para_id in para_ids:
            ext_part.remove_comment_ex(para_id)
            ids_part.remove_comment_id(para_id)
            durable_id = durable_ids.get(para_id)
            if durable_id:
                extensible_part.remove_comment_extensible(durable_id)

    @staticmethod
    def _append_text_content(
        para: etree._Element, text: str, rsid_rpr: str, fmt: Optional[dict] = None
    ) -> None:
        """Append the comment text as runs, preserving whitespace fidelity.

        Newlines become w:br and tabs w:tab (literal \\n/\\t inside w:t are
        collapsed by Word); chunks with leading/trailing whitespace get
        xml:space="preserve" so Word does not trim them. `fmt` may request
        bold/italic/underline run properties.
        """
        run = etree.SubElement(para, _qn(NS_W, "r"))
        run.set(_qn(NS_W, "rsidRPr"), rsid_rpr)
        if fmt:
            rpr = etree.SubElement(run, _qn(NS_W, "rPr"))
            if fmt.get("bold"):
                etree.SubElement(rpr, _qn(NS_W, "b"))
            if fmt.get("italic"):
                etree.SubElement(rpr, _qn(NS_W, "i"))
            if fmt.get("underline"):
                u = etree.SubElement(rpr, _qn(NS_W, "u"))
                u.set(_qn(NS_W, "val"), "single")

        def append_chunk(chunk: str) -> None:
            t = etree.SubElement(run, _qn(NS_W, "t"))
            t.text = chunk
            if chunk.strip() != chunk:
                t.set(_qn(NS_XML, "space"), "preserve")

        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        buffer: list[str] = []
        emitted = False
        for char in normalized:
            if char in ("\n", "\t"):
                if buffer:
                    append_chunk("".join(buffer))
                    buffer.clear()
                etree.SubElement(run, _qn(NS_W, "br" if char == "\n" else "tab"))
                emitted = True
            else:
                buffer.append(char)
        if buffer:
            append_chunk("".join(buffer))
            emitted = True
        if not emitted:
            etree.SubElement(run, _qn(NS_W, "t"))

    def _build_comment_paragraph(
        self,
        runs: list[tuple[str, dict]],
        para_id: str,
        text_id: str,
        rsid_r: str,
        rsid_default: str,
        rsid_rpr: str,
        include_annotation_ref: bool,
    ) -> etree._Element:
        """Build one detached comment w:p element; the caller appends it."""
        para = etree.Element(_qn(NS_W, "p"))
        para.set(_qn(NS_W, "rsidR"), rsid_r)
        para.set(_qn(NS_W, "rsidRDefault"), rsid_default)
        para.set(_qn(NS_W14, "paraId"), para_id)
        para.set(_qn(NS_W14, "textId"), text_id)

        pPr = etree.SubElement(para, _qn(NS_W, "pPr"))
        pStyle = etree.SubElement(pPr, _qn(NS_W, "pStyle"))
        pStyle.set(_qn(NS_W, "val"), "CommentText")

        # The annotationRef run marks the comment reference; Word puts it on
        # the first paragraph only.
        if include_annotation_ref:
            run1 = etree.SubElement(para, _qn(NS_W, "r"))
            rPr = etree.SubElement(run1, _qn(NS_W, "rPr"))
            rStyle = etree.SubElement(rPr, _qn(NS_W, "rStyle"))
            rStyle.set(_qn(NS_W, "val"), "CommentReference")
            etree.SubElement(run1, _qn(NS_W, "annotationRef"))

        if not runs:
            self._append_text_content(para, "", rsid_rpr)
        for run_text, fmt in runs:
            self._append_text_content(para, run_text, rsid_rpr, fmt)
        return para

    def _add_comment_xml(
        self,
        comment_id: str,
        para_id: str,
        text_id: str,
        content: list[list[tuple[str, dict]]],
        author: str,
        initials: Optional[str],
        timestamp: Optional[datetime] = None,
        used_hex_ids: Optional[set[str]] = None,
    ) -> datetime:
        """Add a comment element to comments.xml and return its timestamp."""
        rsid_r = uuid.uuid4().hex[:8].upper()
        rsid_default = uuid.uuid4().hex[:8].upper()
        rsid_rpr = uuid.uuid4().hex[:8].upper()

        # Build the comment detached so a failure partway through cannot
        # leave a half-built comment in the saved document.
        comment = etree.Element(_qn(NS_W, "comment"))
        comment.set(_qn(NS_W, "id"), comment_id)
        comment.set(_qn(NS_W, "author"), author)
        if initials:
            comment.set(_qn(NS_W, "initials"), initials)
        # Use local time with offset so Word displays the expected timestamp.
        if timestamp is None:
            timestamp = datetime.now().astimezone()
        elif timestamp.tzinfo is None:
            # Interpret naive datetimes as local time (matching the default).
            timestamp = timestamp.astimezone()
        comment.set(
            _qn(NS_W, "date"),
            timestamp.isoformat(timespec="seconds"),
        )

        if used_hex_ids is None:
            used_hex_ids = self._used_long_hex_ids()
        for index, runs in enumerate(content):
            is_last = index == len(content) - 1
            # Word keys satellite metadata to the LAST paragraph.
            comment.append(
                self._build_comment_paragraph(
                    runs,
                    para_id=para_id if is_last else self._new_long_hex_id(used_hex_ids),
                    text_id=text_id if is_last else self._new_long_hex_id(used_hex_ids),
                    rsid_r=rsid_r,
                    rsid_default=rsid_default,
                    rsid_rpr=rsid_rpr,
                    include_annotation_ref=index == 0,
                )
            )

        # Attach and save
        self._comments_xml.append(comment)
        self._save_comments()
        return timestamp

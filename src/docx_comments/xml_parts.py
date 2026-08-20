"""Handlers for XML parts: comments.xml, commentsExtended.xml, commentsIds.xml,
commentsExtensible.xml, and people.xml."""

from __future__ import annotations

import warnings
from typing import TYPE_CHECKING, Any, Optional

from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree

from docx_comments.exceptions import PersonNotFoundError
from docx_comments.models import PersonInfo

if TYPE_CHECKING:
    from docx.document import Document

try:  # python-docx >= 1.2.0 registers its own XmlPart class for comments.xml
    from docx.parts.comments import CommentsPart as _NativeCommentsPart
except ImportError:  # pragma: no cover - python-docx < 1.2.0
    _NativeCommentsPart = None  # type: ignore[assignment, misc]


# OOXML Namespaces
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
NS_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
NS_W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Relationship types
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_COMMENTS_EXT = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
REL_COMMENTS_IDS = (
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
)
REL_PEOPLE = "http://schemas.microsoft.com/office/2011/relationships/people"
REL_COMMENTS_EXTENSIBLE = (
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
)

# Content types
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_COMMENTS_EXT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
CT_COMMENTS_IDS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
)
CT_PEOPLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"
CT_COMMENTS_EXTENSIBLE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"
)


def _qn(ns: str, name: str) -> str:
    """Create qualified name with namespace."""
    return f"{{{ns}}}{name}"


# Attribute names used to cache a parsed XML tree on a generic (blob-backed)
# Part instance so every handler shares one live tree per part. The blob
# snapshot detects consumers replacing part._blob directly, which invalidates
# the cached tree.
_CACHED_ELEMENT_ATTR = "_docx_comments_element"
_CACHED_BLOB_ATTR = "_docx_comments_blob_snapshot"

# lxml parses blobs coming from inside the package; disable entity resolution
# and network access so a hostile document cannot use DOCTYPE tricks (XXE)
# through this library's own parsing.
_SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)


def parse_xml_bytes(blob: bytes) -> etree._Element:
    """Parse XML bytes with entity resolution disabled."""
    return etree.fromstring(blob, _SAFE_PARSER)


def validate_xml_text(value: Optional[str], what: str) -> None:
    """Raise a clear ValueError when a string cannot be stored in XML."""
    if value is None:
        return
    probe = etree.Element("probe")
    try:
        probe.text = value
    except ValueError as exc:
        raise ValueError(
            f"{what} contains characters not allowed in XML: {exc}"
        ) from exc


def part_is_blob_backed(part: Any) -> bool:
    """True when the part serialises from its ``_blob`` (generic Part).

    XmlPart subclasses serialise from their live ``_element``/``element``
    instead, so writes to ``_blob`` would be silently ignored for them.
    """
    return not (hasattr(part, "element") or hasattr(part, "_element"))


def part_element(part: Any) -> Optional[etree._Element]:
    """Return a live, mutable XML root for a part.

    XmlPart subclasses expose their element directly and mutations persist on
    save. Generic blob Parts get a parsed tree cached on the part object
    itself so every handler instance shares one tree; call
    :func:`sync_part_blob` after mutating it.
    """
    if part is None:
        return None

    if hasattr(part, "element"):
        try:
            elem = part.element
            if elem is not None:
                return elem
        except (AttributeError, TypeError, ValueError, etree.XMLSyntaxError):
            # Best-effort fallback for python-docx element access.
            pass

    if hasattr(part, "_element"):
        if getattr(part, "_element", None) is None:
            try:
                part._element = parse_xml_bytes(part.blob)
            except (AttributeError, TypeError, etree.XMLSyntaxError):
                return None
        return part._element

    try:
        blob = part.blob
    except (AttributeError, TypeError):
        return None
    cached = getattr(part, _CACHED_ELEMENT_ATTR, None)
    # Re-parse when the blob was replaced behind our back (identity check:
    # sync_part_blob records the blob it wrote).
    if cached is None or getattr(part, _CACHED_BLOB_ATTR, None) is not blob:
        try:
            cached = parse_xml_bytes(blob)
        except (TypeError, ValueError, etree.XMLSyntaxError):
            return None
        setattr(part, _CACHED_ELEMENT_ATTR, cached)
        setattr(part, _CACHED_BLOB_ATTR, blob)
    return cached


def sync_part_blob(part: Any) -> None:
    """Persist the cached tree of a blob-backed part back into its blob.

    No-op for XmlPart-backed parts (their element mutations persist on save)
    and for parts whose tree was never parsed.
    """
    if part is None or not part_is_blob_backed(part):
        return
    cached = getattr(part, _CACHED_ELEMENT_ATTR, None)
    if cached is not None:
        blob = etree.tostring(
            cached,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        part._blob = blob
        setattr(part, _CACHED_BLOB_ATTR, blob)


def ensure_mc_ignorable(root: etree._Element, prefixes: tuple = ("w14", "w15")) -> bool:
    """Add missing prefixes to the root's mc:Ignorable attribute.

    Only prefixes actually declared on the root are added (ISO 29500-3
    requires ignorable prefixes to be in scope), and only when the mc
    namespace itself is declared (lxml cannot add declarations to an
    existing root). Returns True when the attribute was modified.
    """
    declared = root.nsmap or {}
    if NS_MC not in declared.values():
        return False
    current = (root.get(_qn(NS_MC, "Ignorable")) or "").split()
    additions = [p for p in prefixes if p in declared and p not in current]
    if not additions:
        return False
    root.set(_qn(NS_MC, "Ignorable"), " ".join(current + additions))
    return True


class _BasePartHandler:
    """Shared plumbing for the comment-related part handlers.

    Reads prefer an XmlPart's live element; generic blob Parts get one parsed
    tree cached on the part object (see :func:`part_element`) so concurrent
    handler instances cannot clobber each other's writes.
    """

    _partname: str
    _reltype: str
    _content_type: str
    _root_tag: str
    _nsmap: dict
    _ignorable: str

    def __init__(self, document: Document) -> None:
        self._document = document

    def _get_part(self) -> Any:
        """Get the part from the main document part's relationships."""
        for rel in self._document.part.rels.values():
            if self._reltype in rel.reltype:
                return rel.target_part
        return None

    def ensure_exists(self) -> None:
        """Ensure the part exists, creating it if needed."""
        if self._get_part() is None:
            self._create_part()

    def _new_root(self) -> etree._Element:
        root = etree.Element(self._root_tag, nsmap=self._nsmap)
        if self._ignorable:
            root.set(_qn(NS_MC, "Ignorable"), self._ignorable)
        return root

    def _create_part(self) -> None:
        xml_content = etree.tostring(
            self._new_root(),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        part = Part(
            PackURI(self._partname),
            self._content_type,
            xml_content,
            self._document.part.package,
        )
        self._document.part.relate_to(part, self._reltype)

    @property
    def xml(self) -> etree._Element:
        """Get the XML root element.

        When the part is missing, a detached empty root is returned so read
        paths see "no entries". Mutating methods call ensure_exists() first,
        so writes never land on a detached element.

        Raises:
            ValueError: If the part exists but its XML cannot be parsed
                (raising is safer than silently dropping reads and writes).
        """
        part = self._get_part()
        if part is None:
            return self._new_root()
        elem = part_element(part)
        if elem is None:
            raise ValueError(
                f"cannot read {self._partname}: the part exists but its XML "
                "cannot be parsed"
            )
        return elem

    def _save(self) -> None:
        """Persist changes for blob-backed parts (no-op for XmlPart)."""
        sync_part_blob(self._get_part())


class CommentsPart(_BasePartHandler):
    """Handler for word/comments.xml part.

    Note: python-docx >= 1.2.0 loads comments.xml as an XmlPart subclass which
    serialises from its live element, while older versions (and some created
    parts) are generic blob Parts. part_element()/sync_part_blob() handle both.
    """

    _partname = "/word/comments.xml"
    _reltype = REL_COMMENTS
    _content_type = CT_COMMENTS
    _root_tag = _qn(NS_W, "comments")
    _nsmap = {"w": NS_W, "w14": NS_W14, "w15": NS_W15, "mc": NS_MC}
    _ignorable = "w14 w15"

    def _create_part(self) -> None:
        """Create a new comments.xml part.

        Prefer python-docx's registered CommentsPart class so its native
        comments API (doc.comments / doc.add_comment) keeps working on the
        same in-memory document.
        """
        if _NativeCommentsPart is not None:
            part = _NativeCommentsPart.default(self._document.part.package)
            # The native template declares w14/mc but omits mc:Ignorable;
            # this library writes w14:paraId/textId, so declare it ignorable.
            ensure_mc_ignorable(part.element)
            self._document.part.relate_to(part, REL_COMMENTS)
            return
        super()._create_part()  # pragma: no cover - python-docx < 1.2.0

    def remove_comment(self, comment_id: str) -> Optional[list[str]]:
        """
        Remove a comment from comments.xml.

        Args:
            comment_id: Comment ID to remove.

        Returns:
            List of paraIds found on the removed comment, or None if not found.
        """
        removed_para_ids: list[str] = []
        removed_count = 0

        for elem in list(self.xml):
            if etree.QName(elem).localname != "comment":
                continue
            if elem.get(_qn(NS_W, "id")) != comment_id:
                continue
            for para in elem.findall(_qn(NS_W, "p")):
                para_id = para.get(_qn(NS_W14, "paraId"))
                if para_id:
                    removed_para_ids.append(para_id)
            elem.getparent().remove(elem)
            removed_count += 1

        if removed_count:
            if removed_count > 1:
                warnings.warn(
                    f"multiple comments share id {comment_id}; all were removed",
                    UserWarning,
                    stacklevel=2,
                )
            self._save()
            return removed_para_ids

        return None


def ensure_comment_parts(document: Document) -> None:
    """
    Ensure all required comment parts exist in the document.

    Creates:
    - comments.xml if missing
    - commentsExtended.xml if missing
    - commentsIds.xml if missing
    - commentsExtensible.xml if missing
    """
    CommentsPart(document).ensure_exists()
    CommentsExtendedPart(document).ensure_exists()
    CommentsIdsPart(document).ensure_exists()
    CommentsExtensiblePart(document).ensure_exists()


class CommentsExtendedPart(_BasePartHandler):
    """Handler for word/commentsExtended.xml part."""

    _partname = "/word/commentsExtended.xml"
    _reltype = REL_COMMENTS_EXT
    _content_type = CT_COMMENTS_EXT
    _root_tag = _qn(NS_W15, "commentsEx")
    _nsmap = {"mc": NS_MC, "w15": NS_W15}
    _ignorable = "w15"

    def get_threading_info(self) -> dict[str, dict]:
        """
        Get threading information for all comments.

        Returns:
            Dict mapping para_id to {"parent_para_id": str|None, "done": bool}
        """
        result = {}
        for elem in self.xml:
            if etree.QName(elem).localname == "commentEx":
                para_id = elem.get(_qn(NS_W15, "paraId"))
                parent = elem.get(_qn(NS_W15, "paraIdParent"))
                done_raw = elem.get(_qn(NS_W15, "done"), "0")
                # w15:done is ST_OnOff: "1"/"true"/"on" are all resolved.
                done = done_raw.strip().lower() in ("1", "true", "on")
                if para_id:
                    result[para_id] = {
                        "parent_para_id": parent,
                        "done": done,
                    }
        return result

    def add_comment_ex(
        self,
        para_id: str,
        parent_para_id: Optional[str] = None,
        done: bool = False,
    ) -> None:
        """
        Add a commentEx entry.

        Args:
            para_id: Paragraph ID of the comment.
            parent_para_id: Paragraph ID of parent (for replies).
            done: Whether comment is resolved.
        """
        self.ensure_exists()
        elem = etree.Element(_qn(NS_W15, "commentEx"))
        elem.set(_qn(NS_W15, "paraId"), para_id)
        elem.set(_qn(NS_W15, "done"), "1" if done else "0")
        if parent_para_id:
            elem.set(_qn(NS_W15, "paraIdParent"), parent_para_id)
        inserted = False
        if parent_para_id:
            for existing in self.xml:
                if (
                    etree.QName(existing).localname == "commentEx"
                    and existing.get(_qn(NS_W15, "paraId")) == parent_para_id
                ):
                    existing.addnext(elem)
                    inserted = True
                    break
        if not inserted:
            self.xml.append(elem)
        self._save()

    def set_done(self, para_id: str, done: bool) -> None:
        """
        Set the done status for a comment.

        Args:
            para_id: Paragraph ID of the comment.
            done: Whether comment is resolved.
        """
        updated = False
        # Update every matching entry: documents from other tools may carry
        # duplicate commentEx elements for one paraId (get_threading_info is
        # last-match-wins, so a first-match-only write would not stick).
        for elem in self.xml:
            if etree.QName(elem).localname == "commentEx":
                if elem.get(_qn(NS_W15, "paraId")) == para_id:
                    elem.set(_qn(NS_W15, "done"), "1" if done else "0")
                    updated = True
        if updated:
            self._save()
            return
        raise ValueError(f"Comment with para_id {para_id} not found in commentsExtended")

    def set_parent(self, para_id: str, parent_para_id: Optional[str]) -> bool:
        """
        Update the parent paraId for a comment.

        Args:
            para_id: Paragraph ID of the comment.
            parent_para_id: New parent paraId, or None to clear.

        Returns:
            True if an entry was updated, False otherwise.
        """
        updated = False
        # Update every matching entry: documents from other tools may carry
        # duplicate commentEx elements for one paraId.
        for elem in self.xml:
            if etree.QName(elem).localname != "commentEx":
                continue
            if elem.get(_qn(NS_W15, "paraId")) != para_id:
                continue
            if parent_para_id:
                elem.set(_qn(NS_W15, "paraIdParent"), parent_para_id)
            else:
                elem.attrib.pop(_qn(NS_W15, "paraIdParent"), None)
            updated = True
        if updated:
            self._save()
        return updated

    def remove_comment_ex(self, para_id: str) -> bool:
        """
        Remove a commentEx entry by paraId.

        Args:
            para_id: Paragraph ID of the comment.

        Returns:
            True if an entry was removed, False otherwise.
        """
        removed = False
        for elem in list(self.xml):
            if etree.QName(elem).localname != "commentEx":
                continue
            if elem.get(_qn(NS_W15, "paraId")) != para_id:
                continue
            elem.getparent().remove(elem)
            removed = True
        if removed:
            self._save()
        return removed


class CommentsExtensiblePart(_BasePartHandler):
    """Handler for word/commentsExtensible.xml part."""

    _partname = "/word/commentsExtensible.xml"
    _reltype = REL_COMMENTS_EXTENSIBLE
    _content_type = CT_COMMENTS_EXTENSIBLE
    _root_tag = _qn(NS_W16CEX, "commentsExtensible")
    _nsmap = {"mc": NS_MC, "w16cex": NS_W16CEX}
    _ignorable = "w16cex"

    def _get_part(self) -> Any:
        """Get the commentsExtensible part.

        Falls back to a package-wide partname scan for documents whose part
        exists but is not related from the main document part; Word only
        loads the part through that relationship, so it is added on the spot.
        """
        doc_part = self._document.part
        for rel in doc_part.rels.values():
            if REL_COMMENTS_EXTENSIBLE in rel.reltype:
                return rel.target_part
        package = getattr(doc_part, "package", None)
        if package is not None:
            for part in getattr(package, "parts", []):
                if str(part.partname) == "/word/commentsExtensible.xml":
                    doc_part.relate_to(part, REL_COMMENTS_EXTENSIBLE)
                    return part
        return None

    def get_extensible_info(self) -> dict[str, dict]:
        """
        Get metadata entries from commentsExtensible.xml.

        Returns:
            Dict mapping durable_id to {"date_utc": str|None}.
        """
        result = {}
        for elem in self.xml:
            if etree.QName(elem).localname == "commentExtensible":
                durable_id = elem.get(_qn(NS_W16CEX, "durableId"))
                date_utc = elem.get(_qn(NS_W16CEX, "dateUtc"))
                if durable_id:
                    result[durable_id] = {"date_utc": date_utc}
        return result

    def add_comment_extensible(self, durable_id: str, date_utc: Optional[str] = None) -> None:
        """
        Add or update a commentExtensible entry.

        Args:
            durable_id: Durable ID for the comment.
            date_utc: Optional UTC timestamp (ISO8601, Z-terminated).
        """
        self.ensure_exists()
        for elem in self.xml:
            if (
                etree.QName(elem).localname == "commentExtensible"
                and elem.get(_qn(NS_W16CEX, "durableId")) == durable_id
            ):
                if date_utc and not elem.get(_qn(NS_W16CEX, "dateUtc")):
                    elem.set(_qn(NS_W16CEX, "dateUtc"), date_utc)
                    self._save()
                return

        elem = etree.SubElement(self.xml, _qn(NS_W16CEX, "commentExtensible"))
        elem.set(_qn(NS_W16CEX, "durableId"), durable_id)
        if date_utc:
            elem.set(_qn(NS_W16CEX, "dateUtc"), date_utc)
        self._save()

    def remove_comment_extensible(self, durable_id: str) -> bool:
        """
        Remove a commentExtensible entry by durableId.

        Args:
            durable_id: Durable ID for the comment.

        Returns:
            True if an entry was removed, False otherwise.
        """
        removed = False
        for elem in list(self.xml):
            if etree.QName(elem).localname != "commentExtensible":
                continue
            if elem.get(_qn(NS_W16CEX, "durableId")) != durable_id:
                continue
            elem.getparent().remove(elem)
            removed = True
        if removed:
            self._save()
        return removed


class CommentsIdsPart(_BasePartHandler):
    """Handler for word/commentsIds.xml part."""

    _partname = "/word/commentsIds.xml"
    _reltype = REL_COMMENTS_IDS
    _content_type = CT_COMMENTS_IDS
    _root_tag = _qn(NS_W16CID, "commentsIds")
    _nsmap = {"mc": NS_MC, "w16cid": NS_W16CID}
    _ignorable = "w16cid"

    def get_durable_ids(self) -> dict[str, str]:
        """
        Get durable IDs for all comments.

        Returns:
            Dict mapping para_id to durable_id.
        """
        result = {}
        for elem in self.xml:
            if etree.QName(elem).localname == "commentId":
                para_id = elem.get(_qn(NS_W16CID, "paraId"))
                durable_id = elem.get(_qn(NS_W16CID, "durableId"))
                if para_id and durable_id:
                    result[para_id] = durable_id
        return result

    def add_comment_id(self, para_id: str, durable_id: str) -> None:
        """
        Add a commentId entry.

        Args:
            para_id: Paragraph ID of the comment.
            durable_id: Durable ID for persistence.
        """
        self.ensure_exists()
        elem = etree.SubElement(self.xml, _qn(NS_W16CID, "commentId"))
        elem.set(_qn(NS_W16CID, "paraId"), para_id)
        elem.set(_qn(NS_W16CID, "durableId"), durable_id)
        self._save()

    def remove_comment_id(self, para_id: str) -> Optional[str]:
        """
        Remove a commentId entry by paraId.

        Args:
            para_id: Paragraph ID of the comment.

        Returns:
            The durableId removed, or None if not found.
        """
        removed_durable_id = None
        removed = False
        for elem in list(self.xml):
            if etree.QName(elem).localname != "commentId":
                continue
            if elem.get(_qn(NS_W16CID, "paraId")) != para_id:
                continue
            removed_durable_id = elem.get(_qn(NS_W16CID, "durableId"))
            elem.getparent().remove(elem)
            removed = True
        if removed:
            self._save()
        return removed_durable_id


class PeoplePart(_BasePartHandler):
    """Handler for word/people.xml part."""

    _partname = "/word/people.xml"
    _reltype = REL_PEOPLE
    _content_type = CT_PEOPLE
    _root_tag = _qn(NS_W15, "people")
    _nsmap = {"mc": NS_MC, "w": NS_W, "w14": NS_W14, "w15": NS_W15, "wp14": NS_WP14}
    _ignorable = "w14 w15 wp14"

    @staticmethod
    def _attr_by_localname(elem: etree._Element, localname: str) -> Optional[str]:
        for attr, value in elem.attrib.items():
            try:
                if etree.QName(attr).localname == localname:
                    return str(value)
            except (ValueError, TypeError):
                if attr == localname:
                    return str(value)
        return None

    @staticmethod
    def _find_child_by_localname(
        elem: etree._Element, localname: str
    ) -> Optional[etree._Element]:
        for child in elem:
            if etree.QName(child).localname == localname:
                return child
        return None

    def _person_info_from_elem(self, elem: etree._Element) -> PersonInfo:
        author = self._attr_by_localname(elem, "author") or ""
        presence_elem = self._find_child_by_localname(elem, "presenceInfo")
        provider_id = user_id = None
        if presence_elem is not None:
            provider_id = self._attr_by_localname(presence_elem, "providerId")
            user_id = self._attr_by_localname(presence_elem, "userId")
        return PersonInfo(author=author, provider_id=provider_id, user_id=user_id)

    def get_people(self) -> list[PersonInfo]:
        """List people entries in people.xml."""
        if self._get_part() is None:
            return []
        people: list[PersonInfo] = []
        for elem in self.xml:
            if etree.QName(elem).localname == "person":
                people.append(self._person_info_from_elem(elem))
        return people

    def _find_person_elem(self, author: str) -> Optional[etree._Element]:
        if self._get_part() is None:
            return None
        for elem in self.xml:
            if etree.QName(elem).localname != "person":
                continue
            if self._attr_by_localname(elem, "author") == author:
                return elem
        return None

    def get_person(self, author: str) -> PersonInfo:
        """Return a person entry by author name."""
        if not author:
            raise ValueError("author must be non-empty")
        elem = self._find_person_elem(author)
        if elem is None:
            raise PersonNotFoundError(f"person '{author}' not found")
        return self._person_info_from_elem(elem)

    @staticmethod
    def _normalize_presence(presence: dict[str, str]) -> tuple[str, str]:
        provider_id = presence.get("provider_id") or presence.get("providerId")
        user_id = presence.get("user_id") or presence.get("userId")
        if not provider_id or not user_id:
            raise ValueError("presence must include provider_id and user_id")
        return provider_id, user_id

    def ensure_person(
        self, author: str, presence: Optional[dict[str, str]] = None
    ) -> PersonInfo:
        """Ensure a person entry exists, optionally adding presence metadata."""
        if not author:
            raise ValueError("author must be non-empty")
        validate_xml_text(author, "person author")

        # Validate everything before touching the package or the cached
        # tree, so a bad spec cannot leave a part or half-built entry behind.
        normalized: Optional[tuple[str, str]] = None
        if presence:
            normalized = self._normalize_presence(presence)
            validate_xml_text(normalized[0], "presence provider_id")
            validate_xml_text(normalized[1], "presence user_id")

        person_elem = self._find_person_elem(author)
        if person_elem is None:
            # Build detached and attach only when complete.
            new_elem = etree.Element(_qn(NS_W15, "person"))
            new_elem.set(_qn(NS_W15, "author"), author)
            if normalized:
                presence_elem = etree.SubElement(new_elem, _qn(NS_W15, "presenceInfo"))
                presence_elem.set(_qn(NS_W15, "providerId"), normalized[0])
                presence_elem.set(_qn(NS_W15, "userId"), normalized[1])
            self.ensure_exists()
            self.xml.append(new_elem)
            person_elem = new_elem
        elif normalized:
            presence_elem = self._find_child_by_localname(person_elem, "presenceInfo")
            if presence_elem is None:
                presence_elem = etree.SubElement(person_elem, _qn(NS_W15, "presenceInfo"))
            presence_elem.set(_qn(NS_W15, "providerId"), normalized[0])
            presence_elem.set(_qn(NS_W15, "userId"), normalized[1])

        self._save()
        return self._person_info_from_elem(person_elem)

    def merge_from(
        self, source: "PeoplePart", include_presence: bool = False
    ) -> list[PersonInfo]:
        """
        Merge people entries from another document.

        Existing authors are preserved; new authors are added.
        """
        if source._get_part() is None:
            return []

        existing_authors = {person.author for person in self.get_people()}
        added: list[PersonInfo] = []

        for person in source.get_people():
            if not person.author or person.author in existing_authors:
                continue

            presence = None
            if include_presence and person.provider_id and person.user_id:
                presence = {
                    "provider_id": person.provider_id,
                    "user_id": person.user_id,
                }

            added.append(self.ensure_person(person.author, presence))
            existing_authors.add(person.author)

        return added

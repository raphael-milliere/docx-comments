# python-docx-comments

[![PyPI version](https://badge.fury.io/py/python-docx-comments.svg)](https://pypi.org/project/python-docx-comments/)
[![Python versions](https://img.shields.io/pypi/pyversions/python-docx-comments.svg)](https://pypi.org/project/python-docx-comments/)
[![CI](https://github.com/raphael-milliere/docx-comments/actions/workflows/ci.yml/badge.svg)](https://github.com/raphael-milliere/docx-comments/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Python module for complete Word document comment manipulation - adding, replying, editing, and resolving comments with full Word Online compatibility. Installed as `python-docx-comments`, imported as `docx_comments`.

> **Note:** this is the actively maintained fork of
> [sunt05/docx-comments](https://github.com/sunt05/docx-comments), published
> under a new PyPI name. The legacy `docx-comments` package on PyPI (0.3.0)
> predates important id-generation fixes and should not be used.

## Problem

python-docx >= 1.2.0 can create basic comments (content plus anchors), but has
no support for the rest of the comment machinery Word and Word Online rely on:
- No `commentsExtended.xml` (threading, resolution/done status)
- No `commentsIds.xml` (durable IDs)
- No `commentsExtensible.xml` (modern comment metadata)
- No reply, edit, resolve, delete, or move operations
- No way to anchor a comment to an exact substring or character range

Microsoft Graph API does NOT support Word comments (only Excel).

## Solution

This module provides complete OOXML comment manipulation based on ECMA-376 / ISO/IEC 29500:
- Add anchored comments to specific text ranges
- Anchor to exact text: substring/regex matching or character offsets
- Reply to existing comments (threaded)
- Edit comment text in place (threading, resolution, and anchors preserved)
- Rich comment content: bold/italic/underline runs, multiple paragraphs
- Mark comments as resolved
- Unresolve comments and toggle done status
- Delete comments or entire threads
- Move comment anchors to new locations
- Read back comments, threads, and the text a comment is anchored to
- Full Word Online compatibility
- Optional people.xml identity linkage (Word account presence)

## Installation

```bash
pip install python-docx-comments
```

## Usage

```python
from docx import Document
from docx_comments import CommentManager

doc = Document("document.docx")
mgr = CommentManager(doc)  # creates no parts until the first write

# author can be a plain string, or PersonInfo for identity linkage

# Add an anchored comment. By default the whole paragraph is anchored;
# start_run/end_run select a run range (indices are validated — out-of-range
# values raise IndexError rather than anchoring the wrong text).
comment_id = mgr.add_comment(
    paragraph=doc.paragraphs[0],
    text="Please review this section",
    author="Reviewer Name",
    initials="RN",
)

# Reply to existing comment
reply_id = mgr.reply_to_comment(
    parent_id=comment_id,
    text="Addressed in this revision",
    author="Author Name",
    initials="AN"
)

# Mark the comment's thread as resolved (thread-scoped, like Word)
mgr.resolve_comment(comment_id)

# Mark it as unresolved again
mgr.unresolve_comment(comment_id)

# Move the whole thread to a new paragraph. (move_comment() moves a single
# standalone comment and raises for threaded comments, so reply anchors can
# never be left behind.)
mgr.move_thread(
    comment_id=comment_id,
    paragraph=doc.paragraphs[1],
)

# Delete a comment thread (root + replies)
mgr.delete_thread(comment_id)

# List all comment threads
for thread in mgr.get_comment_threads():
    print(f"Root: {thread.root.text} by {thread.root.author}")
    for reply in thread.replies:
        print(f"  Reply: {reply.text} by {reply.author}")

doc.save("document_reviewed.docx")
```

### Anchoring to exact text

Instead of run indices, a comment can be anchored to the exact characters it
is about. Runs are split as needed; the document's visible text is never
changed.

```python
import re

para = doc.paragraphs[0]  # "Please review this section before the deadline."

# Anchor to a substring (first occurrence by default)
cid = mgr.add_comment_on_text(para, "review this", "Tighten the wording", "Reviewer Name")

# Nth occurrence, or a compiled regular expression
cid = mgr.add_comment_on_text(para, "e", "Second 'e'", "Reviewer Name", occurrence=2)
cid = mgr.add_comment_on_text(para, re.compile(r"\bsection\b"), "Rename?", "Reviewer Name")

# Or give character offsets directly (end is exclusive, measured over the
# paragraph's visible text)
cid = mgr.add_comment(para, "First word", "Reviewer Name", start_char=0, end_char=6)
```

### Editing and reading back

```python
# Edit a comment in place — id, durable id, threading, resolution state,
# and anchors are all preserved; only the content changes.
mgr.edit_comment(comment_id, "Please review this section (updated)")

# Optionally change author/initials/date at the same time
from datetime import datetime, timezone

mgr.edit_comment(
    comment_id,
    "Final wording agreed",
    author="Editor Name",
    initials="EN",
    timestamp=datetime(2026, 8, 20, 12, 0, tzinfo=timezone.utc),
)

# Read back a single comment, its thread, and its anchor
info = mgr.get_comment(comment_id)            # CommentInfo
thread = mgr.get_thread(comment_id)           # CommentThread (root + replies)
snippet = mgr.get_anchored_text(comment_id)   # anchored document text, or None
para = mgr.get_comment_paragraph(comment_id)  # Paragraph with the anchor, or None
```

### Rich content and timestamps

```python
# text is one paragraph of plain text by default; pass a sequence of
# paragraphs for more. Each paragraph is a str or a sequence of runs,
# where a run is a str or (text, {"bold"/"italic"/"underline": True}).
mgr.add_comment(
    paragraph=doc.paragraphs[0],
    text=[
        [("Blocking:", {"bold": True}), " needs a citation."],
        "See the style guide.",
    ],
    author="Reviewer Name",
)

# add_comment / reply_to_comment / edit_comment accept an explicit timestamp
# (naive datetimes are interpreted as local time; the default is "now")
mgr.add_comment(
    paragraph=doc.paragraphs[0],
    text="Imported from the previous review round",
    author="Reviewer Name",
    timestamp=datetime(2025, 3, 1, 9, 30, tzinfo=timezone.utc),
)
```

### Errors

Lookup failures raise typed exceptions that subclass the builtins previously
raised, so existing `except ValueError` / `except KeyError` code keeps working:

```python
from docx_comments import CommentNotFoundError, PersonNotFoundError

try:
    mgr.get_comment("999999")
except CommentNotFoundError:  # subclasses ValueError and LookupError
    ...

try:
    mgr.get_person("Nobody")
except PersonNotFoundError:  # subclasses KeyError
    ...
```

Comment ids are strings, but every method that takes a comment id also
accepts the `int` ids used by python-docx's native comments API:

```python
mgr.resolve_comment(int(comment_id))  # same comment as mgr.resolve_comment(comment_id)
```

## Identity Linkage (people.xml)

Word maps `w:comment/@w:author` to account identity using `word/people.xml`. By default, this library does
not create or modify `people.xml` unless you opt in.

```python
# Create a minimal people.xml entry without presence metadata
person = mgr.ensure_person("Reviewer Name")

# Or fetch an existing person entry (raises if missing)
try:
    person = mgr.get_person("Reviewer Name")
except PersonNotFoundError:  # subclasses KeyError
    person = mgr.ensure_person("Reviewer Name")

# Resolve a default author from the system or a DOCX source
person, initials = mgr.get_default_author_person()

# Merge people.xml entries from another document (adds missing authors only)
template_doc = Document("template.docx")
mgr.merge_people_from(template_doc, include_presence=False)

# Or request it when adding a comment
mgr.add_comment(
    paragraph=doc.paragraphs[0],
    text="Linked to people.xml",
    author=person,
    person=True,
)

# You can also pass a PersonInfo object from an existing people.xml
person = mgr.get_people()[0]
mgr.add_comment(
    paragraph=doc.paragraphs[0],
    text="Author from PersonInfo",
    author=person,
)

# Optional presence metadata (only if you explicitly supply it)
mgr.ensure_person(
    "Reviewer Name",
    presence={"provider_id": "provider", "user_id": "user"},
)
```

Note: Word comments are keyed by the author string (`w:comment/@w:author`). If two people share
the same name string, Word does not provide a separate comment author ID to disambiguate them.
Using `people.xml` presence metadata can improve account linkage, but cannot fully resolve
same-name conflicts.

You can also point the resolver at a known DOCX (kept private) using an environment variable:

```bash
export DOCX_COMMENTS_AUTHOR_DOCX="/path/to/author-source.docx"
```

Then call:

```python
person, initials = mgr.get_default_author_person(include_presence=True)
```

If the DOCX contains more than one `w15:person` entry, a warning is raised and the resolver
falls back to system user info.

## API Summary

All public `CommentManager` methods. Every method that takes a comment id
accepts `str` or `int`.

| Method | Description |
| --- | --- |
| `add_comment(paragraph, text, author, ...)` | Add an anchored comment: whole paragraph, run range, or `start_char`/`end_char` span |
| `add_comment_on_text(paragraph, match, text, author, ...)` | Anchor a comment to the nth occurrence of a substring or regex |
| `reply_to_comment(parent_id, text, author, ...)` | Add a threaded reply (inherits the thread's resolved state) |
| `edit_comment(comment_id, text, ...)` | Replace a comment's text (and optionally author/initials/date) in place |
| `resolve_comment(comment_id)` | Mark the comment's thread as resolved |
| `unresolve_comment(comment_id)` | Mark the comment's thread as unresolved |
| `set_comment_resolved(comment_id, resolved)` | Set the thread's resolved state explicitly |
| `delete_comment(comment_id)` | Delete one comment (replies survive, detached) |
| `delete_thread(comment_id)` | Delete an entire thread (root + replies) |
| `move_comment(comment_id, paragraph, ...)` | Re-anchor a standalone comment (raises for threaded comments) |
| `move_thread(comment_id, paragraph, ...)` | Re-anchor an entire thread, keeping reply anchors co-located |
| `list_comments()` | Iterate `CommentInfo` for every comment in the document |
| `get_comment(comment_id)` | Return one comment's `CommentInfo` (raises `CommentNotFoundError`) |
| `get_thread(comment_id)` | Return the `CommentThread` containing a comment (root or reply) |
| `get_comment_threads()` | All comments grouped into threads by root |
| `get_anchored_text(comment_id)` | Document text the comment is anchored to (or `None`) |
| `get_comment_paragraph(comment_id)` | Paragraph containing the comment's anchor (or `None`) |
| `get_authors()` | Map of comment author names to initials |
| `get_document_author()` | Document owner from core properties, with initials if known |
| `get_people()` | List entries from `word/people.xml` |
| `get_person(author)` | Fetch one people.xml entry (raises `PersonNotFoundError`) |
| `ensure_person(author, presence=None)` | Create a people.xml entry if missing |
| `merge_people_from(source, include_presence=False)` | Import missing people.xml entries from another document |
| `get_default_author_person(...)` | Resolve a default author from the system or a DOCX source |
| `migrate_comment_metadata()` | Backfill threading/durable-id/extensible metadata for comments created by other tools |

## OOXML Parts Handled

This module manages six XML parts:

1. **comments.xml** - Comment content and metadata
2. **document.xml** - Anchors (`commentRangeStart/End`, `commentReference`),
   including anchors in headers, footers, tables, and footnotes/endnotes
3. **commentsExtended.xml** - Threading (`paraId`, `paraIdParent`, `done`)
4. **commentsIds.xml** - Durable IDs for persistence
5. **commentsExtensible.xml** - Modern comment metadata (`w16cex:dateUtc`)
6. **people.xml** - Optional identity linkage (`w15:person`)

Parts are created lazily: a `CommentManager` used only for reading
(`list_comments()`, `get_comment_threads()`, …) leaves the document untouched.

## Requirements

- Python 3.9+
- python-docx >= 1.0.0
- lxml

## References

### OOXML Specification

- [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML File Formats (free download)
- [ISO/IEC 29500](https://www.loc.gov/preservation/digital/formats/fdd/fdd000395.shtml) - OOXML Format Family overview
- [MS-DOCX Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/) - Microsoft's DOCX extensions documentation

### Comment Elements

- [commentRangeStart](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_commentRangeStart_topic_ID0EFJMV.html) - Comment anchor range start element
- [commentRangeEnd](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_commentRangeEnd_topic_ID0ESCLV.html) - Comment anchor range end element
- [commentReference](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_commentReference_topic_ID0E4PNV.html) - Comment content reference mark
- [Comments Overview](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_Comments_topic_ID0EEHJV.html) - OOXML comments specification

### Threading & Extended Parts

- [CommentEx Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.word.commentex) - Office 2013 comment threading (paraId, paraIdParent, done)
- [commentsIds](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/22977b5a-5bb5-4f27-b7a1-c6d216c2bb94) - Durable IDs specification (Office 2016+)

### Related Libraries

- [python-docx](https://python-docx.readthedocs.io/) - Python library for Word documents (foundation for this module)
- [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) - Microsoft's .NET SDK for OOXML

## Acknowledgements

This project was conceptualised by [Ting Sun](https://github.com/sunt05) and implemented with the assistance of [Claude Code](https://claude.ai/code) (Anthropic's AI coding assistant) under his guidance. The collaboration involved iterative development of the OOXML comment handling logic, with Claude Code contributing to code implementation and Ting Sun providing architectural direction and domain expertise. This fork is maintained by [Raphaël Millière](https://github.com/raphael-milliere) and published on PyPI as `python-docx-comments`.

## License

MIT

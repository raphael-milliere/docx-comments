# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

`docx_comments` is a Python module for complete Word document comment manipulation (distributed on PyPI as `python-docx-comments`; this repo is the maintained fork of sunt05/docx-comments). It fills gaps in python-docx by providing:
- Anchored comments with proper OOXML structure
- Threaded replies
- Comment resolution (done status)
- Full Word Online compatibility

## Development Commands

```bash
# Setup
uv venv && uv pip install -e ".[dev]"

# Run tests
pytest                                  # Full suite
pytest tests/test_manager_basic.py -v   # Single file with verbose
pytest -k "test_add_comment"            # Single test by name

# Type checking
mypy src/docx_comments

# Linting
ruff check src/ tests/
ruff format src/
```

## Architecture

### OOXML Comment System

Word comments require coordination across six XML parts:

1. **comments.xml** - Comment content (text, author, timestamp)
2. **document.xml** - Anchors linking comments to text ranges (also headers,
   footers, tables, footnotes/endnotes)
3. **commentsExtended.xml** - Threading (parent-child relationships, done status)
4. **commentsIds.xml** - Durable IDs for persistence across edits
5. **commentsExtensible.xml** - Modern comment metadata (w16cex:dateUtc)
6. **people.xml** - Optional author identity linkage (w15:person)

Parts are created lazily on the first mutating operation; read-only use of
`CommentManager` leaves the document untouched.

### Module Structure

- `manager.py` - `CommentManager` class: main public API
  - `add_comment()`, `add_comment_on_text()`, `reply_to_comment()`, `edit_comment()`
  - `resolve_comment()`, `unresolve_comment()`, `set_comment_resolved()`
  - `delete_comment()`, `delete_thread()`, `move_comment()`, `move_thread()`
  - `list_comments()`, `get_comment_threads()`, `get_comment()`, `get_thread()`
  - `get_anchored_text()`, `get_comment_paragraph()` - Anchor introspection
  - `get_authors()`, `get_document_author()` - Author introspection

- `exceptions.py` - Typed exceptions: `CommentNotFoundError` (subclasses
  ValueError + LookupError), `PersonNotFoundError` (subclasses KeyError)

- `xml_parts.py` - Handlers for XML parts
  - `CommentsPart` - Main comments.xml (handles XmlPart vs generic Part serialisation)
  - `CommentsExtendedPart` - Threading info (w15:paraIdParent, w15:done)
  - `CommentsIdsPart` - Durable IDs (w16cid:durableId)
  - `ensure_comment_parts()` - Creates missing parts with proper relationships

- `anchors.py` - `CommentAnchor` class: manages document.xml anchors
  - Inserts `commentRangeStart`, `commentRangeEnd`, `commentReference`
  - Handles empty paragraphs and reply co-location

- `models.py` - Data classes (`CommentInfo`, `CommentThread`, `PersonInfo`)
  and the `CommentContent` type aliases for rich comment content

- `system_author.py` - Internal helpers resolving a default author from a
  DOCX source, the `DOCX_COMMENTS_AUTHOR_DOCX` env var, or system Office
  identity (macOS/Windows)

### Key Implementation Details

**ID Generation** (`manager.py`):
- `comment_id`: random positive 32-bit integer (ST_DecimalNumber is treated
  as Int32 by Word/Open XML SDK/python-docx — never exceed 0x7FFFFFFE)
- `para_id` / `text_id` / `durable_id`: 8 uppercase hex chars
  (ST_LongHexNumber, capped at 0x7FFFFFFE)
- All ids are drawn from a module-private RNG (immune to consumer
  `random.seed()`) and re-drawn until unique against the ids already present
  in the document

**Namespace Prefixes**:
- `w:` - Main WordprocessingML (2006)
- `w14:` - Word 2010 extensions (paraId on paragraphs)
- `w15:` - Word 2012 extensions (threading, done status)
- `w16cid:` - Word 2016 extensions (durable IDs)

**Part Creation** (`xml_parts.py`): comments.xml is created through
python-docx's native `CommentsPart.default()` when available (python-docx
>= 1.2.0), so python-docx's own comments API keeps working; the satellite
parts use python-docx internals (`docx.opc.part.Part`, `docx.opc.packuri.PackURI`)
with correct content types and relationships.

**XmlPart vs blob serialisation** (`xml_parts.py`): python-docx loads some
parts as `XmlPart` subclasses (serialised from their live `_element`) and
others as generic blob `Part`s (serialised from `_blob`). The shared
`_BasePartHandler` prefers a part's live element and, for blob parts, caches
one parsed tree on the part object itself (so concurrent handler instances
share it) and re-serialises the blob on `_save()`. `anchors.py` uses the same
helpers (`part_element` / `sync_part_blob`) for footnotes/endnotes parts.

## Testing Notes

Tests use `tmp_path` fixture for save/reload verification. `tests/conftest.py`
keeps the suite hermetic: an autouse fixture clears `DOCX_COMMENTS_AUTHOR_DOCX`
and stubs the system Office author lookup so no test depends on the
developer's machine. Coverage by file:
- `tests/test_manager_basic.py` — manager init, add/resolve, thread grouping
- `tests/test_threads.py` — replies (incl. reply-to-reply, tables, headers)
- `tests/test_editing.py` — delete/move/unresolve lifecycle
- `tests/test_api_polish.py` — typed exceptions, int comment-id acceptance,
  plain-str authors
- `tests/test_read_api.py` — `get_comment`/`get_thread`/`get_comment_paragraph`/
  `get_anchored_text`
- `tests/test_char_anchoring.py` — `start_char`/`end_char` spans, run
  splitting, `add_comment_on_text`
- `tests/test_rich_content.py` — formatted runs, multi-paragraph content
- `tests/test_timestamps.py` — caller-controlled timestamps, date parsing
- `tests/test_system_author.py` — system/default author resolution helpers
- `tests/test_migration.py` — metadata backfill
- `tests/test_xml.py` — Word Online compatibility (XML structure via zipfile)
- `tests/test_models.py`, `tests/test_people.py` — models and people.xml
- `tests/test_robustness.py` — regression tests from the 2026-08 adversarial
  review (id ranges/uniqueness, text fidelity, strict index validation,
  footnote anchors, interop with python-docx native comments, orphan
  handling, security hardening)

## References

### OOXML Specification

- [ECMA-376 Standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML File Formats (free download)
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
- [Open XML SDK Issue #484](https://github.com/OfficeDev/Open-XML-SDK/issues/484) - commentsIds part discussion

### Implementation Guides

- [MS Learn: Insert Comment](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-comment-into-a-word-processing-document) - C# implementation guide
- [WordprocessingML Structure](https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document) - Document structure overview
- [Office Open XML Anatomy](http://officeopenxml.com/anatomyofOOXML.php) - Package structure explained

### Related Libraries

- [python-docx](https://python-docx.readthedocs.io/) - Python library for Word documents
- [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) - Microsoft's .NET SDK for OOXML

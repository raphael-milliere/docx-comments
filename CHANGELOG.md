# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.4.0] - 2026-08-19

### Fixed

- Generated `w:id` values now stay within the 32-bit `ST_DecimalNumber` range
  the OOXML ecosystem enforces (previously ~87% of ids exceeded it, corrupting
  comments in LibreOffice round-trips and failing Open XML SDK validation)
- All generated ids (`w:id`, `w14:paraId`, `w14:textId`, `w16cid:durableId`)
  are checked for uniqueness against values already in the document, and are
  drawn from a private RNG immune to consumer `random.seed()` calls
- `comments.xml` is now created through python-docx's native comments part
  (when available), so `doc.comments`/`doc.add_comment` keep working after
  using `CommentManager`, and two managers on one document no longer lose
  each other's comments
- Resolution is now thread-scoped like Word: `resolve_comment()` /
  `unresolve_comment()` mark every comment in the thread, and both backfill
  missing metadata so comments created by other tools can be resolved
- Comment text fidelity: leading/trailing whitespace is preserved with
  `xml:space="preserve"`, newlines/tabs are encoded as `w:br`/`w:tab`, and
  `list_comments()` reads back breaks, tabs, and paragraph boundaries
- Comments anchored in footnotes/endnotes are now reachable (reply, delete,
  move) and their anchor edits persist on save
- Run indices in `add_comment()`/`move_comment()`/`move_thread()` are strictly
  validated (Python-style negatives accepted, out-of-range raises) instead of
  being silently clamped to the wrong text range
- Paragraphs whose content lives in `w:hyperlink`/`w:ins`/`w:del` containers
  are anchored around that content instead of getting a collapsed zero-width
  anchor before it
- `remove_anchors()` removes Word-authored reference runs entirely instead of
  leaving ghost `rPr`-only runs that corrupt run indices
- Comment ids read from documents can no longer break anchor lookups
  (matching is done in Python rather than interpolated path predicates)
- `find_paragraph_with_comment()` works for comments in table cells and other
  nested stories
- `move_comment()` on a comment that has replies now raises and points to
  `move_thread()` instead of silently splitting the thread's anchors
- Failed `delete_comment()`/`delete_thread()`/`reply_to_comment()` no longer
  mutate the document before raising; `add_comment()` validates text (illegal
  XML characters) before writing anything
- `migrate_comment_metadata()` reuses an unambiguous unclaimed `paraId`
  instead of flattening a reply into a root, and cleans up orphan metadata
- Anchoring a paragraph from a different `Document` now raises instead of
  silently corrupting both files
- `CommentManager` construction no longer creates the four comment parts;
  they are created lazily on the first mutating operation
- Orphaned replies (parent deleted by another tool) are normalized on read:
  `is_reply` is False and they become thread roots
- Duplicate `w:id` values are handled: deletes warn instead of silently
  removing unrelated comments, and thread deletion no longer aborts midway
- `people.xml` presence metadata is validated consistently across all spec
  forms, and a failed `ensure_person()` no longer creates an empty part
- XML parsed by this library uses a hardened parser (no external entity
  resolution, no network access)
- Type hints are now valid (`docx.document.Document`), making the shipped
  `py.typed` meaningful for consumers; `PersonSpec` is exported

## [0.3.0] - 2026-01-22

### Added

- `unresolve_comment()` and `set_comment_resolved()` for toggling done status
- `delete_comment()` and `delete_thread()` for removing comments and threads
- `move_comment()` and `move_thread()` for re-anchoring comments

## [0.2.0] - 2026-01-21

### Added

- Optional `people.xml` identity linkage for comment authors
- `PersonInfo` data model for people.xml authors and presence metadata
- `CommentManager` people APIs: `get_people()`, `get_person()`, `ensure_person()`,
  `merge_people_from()`, `get_default_author_person()`
- System author resolution from Office profiles or a DOCX source via
  `DOCX_COMMENTS_AUTHOR_DOCX`

### Changed

- **Breaking**: `author` parameters now require `PersonInfo` instead of raw strings
- `add_comment()`/`reply_to_comment()` can optionally ensure people.xml entries

## [0.1.1] - 2026-01-21

### Changed

- Switch to git tag-based versioning via hatch-vcs

## [0.1.0] - 2025-01-09

### Added

- Initial release
- `CommentManager` class for managing Word document comments
- `add_comment()` - Add anchored comments to specific text ranges
- `reply_to_comment()` - Create threaded replies to existing comments
- `resolve_comment()` - Mark comments as resolved (done status)
- `list_comments()` - List all comments in the document
- `get_comment_threads()` - Get comments grouped by thread
- `get_authors()` - Get all comment authors
- `get_document_author()` - Get document core properties author
- Full Word Online compatibility with proper OOXML structure
- Support for all four comment-related XML parts:
  - `comments.xml` - Comment content
  - `document.xml` - Anchors (commentRangeStart/End, commentReference)
  - `commentsExtended.xml` - Threading (paraId, paraIdParent, done)
  - `commentsIds.xml` - Durable IDs

[Unreleased]: https://github.com/sunt05/docx-comments/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/sunt05/docx-comments/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/sunt05/docx-comments/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/sunt05/docx-comments/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/sunt05/docx-comments/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/sunt05/docx-comments/releases/tag/v0.1.0

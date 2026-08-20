# Audit Improvements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers-extended-cc:subagent-driven-development (recommended) or superpowers-extended-cc:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement all 37 confirmed findings from the 2026-08-20 multi-agent audit of docx-comments: correctness bugs, missing capabilities, test-coverage gaps, and CI/docs improvements.

**Architecture:** Sequential tasks on one branch (`audit-improvements`), because manager.py/anchors.py/xml_parts.py are shared across nearly every task. Bug fixes first (they change behavior later tests depend on), then features, then coverage batches, then CI/docs, then final verification with adversarial review.

**Tech Stack:** Python 3.9+, python-docx (>=1.0.0 declared floor; native comments integration >=1.2.0), lxml, pytest, mypy, ruff. `uv` for env management (`.venv` exists; run tools via `. .venv/bin/activate`).

**Spec:** The audit findings with verified evidence and fix sketches: `/tmp/audit_matched.json` (37 findings, all adversarially confirmed). Key architecture context: `CLAUDE.md` in repo root.

## Global Constraints

- **Never break backward compatibility of the public API.** New exception types subclass the old ones (`CommentNotFoundError(ValueError, LookupError)`, `PersonNotFoundError(KeyError)`); parameter types only widen (`str` author additionally accepted, `int` ids additionally accepted); no signature parameter is removed or reordered.
- **Validate before mutating.** Every public mutator must raise all validation errors before touching any XML part (the repo's established invariant — see `_resolve_person_spec` docstring and `tests/test_robustness.py` TestLifecycleSafety).
- **All ids** drawn via the existing `_new_comment_id` / `_new_long_hex_id(used)` machinery — never `random` directly, never above `0x7FFFFFFE`.
- **TDD per task**: write the failing test, see it fail, implement, see it pass. Commit per task with the message given in the task.
- After each task: `pytest -q` must show 0 failures; `mypy src/docx_comments` and `ruff check src/ tests/` must pass before committing.
- Tests use the patterns already in the suite: build docs in-memory with `Document()`, save to `tmp_path`, reopen, assert on `zipfile`/lxml where XML shape matters.
- Line length 100 (ruff config).

**User decisions (already made):** User said "Proceed with all improvements" (2026-08-20) — full scope, this session, autonomous execution.

**Executor decisions (recorded, do not re-litigate):**
1. **Release/publish is OUT of scope** (finding: v0.4.0 stranded on the fork, PyPI serves broken 0.3.0). Publishing requires user-owned account actions (PyPI trusted-publisher config for the fork, or an upstream PR to sunt05). Local prep (CHANGELOG 0.5.0 section) is in Task 17; the canonical-repo decision is surfaced to the user at the end. pyproject URLs stay pointing at upstream until that decision.
2. **Declined** (with audit-verifier backing): styles.xml CommentText/CommentReference style definitions (python-docx also skips them; Word normalizes on save); `list_comments` Iterator→list change (would break `next(mgr.list_comments())` callers); id-pool caching (verifier showed it unsafe against external live-document mutation — the safe subset, a migrate pre-check, is Task 13); mkdocs site (README API table instead); hyperlinks inside comment text and cross-paragraph anchoring (deferred as phase 2 — relationship machinery).
3. `.coverage` repo-hygiene item dropped — verified untracked and already gitignored.

---

### Task 0: Branch + CONTRIBUTING fixes

**Goal:** Create the working branch and fix the stale CONTRIBUTING commands.

**Files:**
- Modify: `CONTRIBUTING.md:35,47,50`

**Acceptance Criteria:**
- [ ] Branch `audit-improvements` created from `main` and checked out
- [ ] CONTRIBUTING.md references `tests/test_manager_basic.py` (not the nonexistent `tests/test_basic.py`)
- [ ] CONTRIBUTING.md lint/format commands cover `src/ tests/` (matching CI)

**Verify:** `git branch --show-current` → `audit-improvements`; `grep -c "test_basic.py" CONTRIBUTING.md` → 0

**Steps:**

- [ ] **Step 1:** `git checkout -b audit-improvements`
- [ ] **Step 2:** In `CONTRIBUTING.md` change line 35 `pytest tests/test_basic.py -v` → `pytest tests/test_manager_basic.py -v`; line 47 `ruff check src/` → `ruff check src/ tests/`; line 50 `ruff format --check src/` → `ruff format --check src/ tests/`; line 53 `ruff format src/` → `ruff format src/ tests/`.
- [ ] **Step 3:** Commit: `git add CONTRIBUTING.md && git commit -m "docs: fix stale test path and lint scope in CONTRIBUTING"`

```json:metadata
{"files": ["CONTRIBUTING.md"], "verifyCommand": "grep -c test_basic.py CONTRIBUTING.md; git branch --show-current", "acceptanceCriteria": ["branch audit-improvements exists", "no test_basic.py reference remains", "lint commands cover src/ and tests/"], "modelTier": "mechanical"}
```

---

### Task 1: Exception types, int-id coercion, honest CommentInfo hints

**Goal:** Distinguishable error types (backward compatible), `int` comment-id acceptance, and type hints that match runtime reality.

**Files:**
- Create: `src/docx_comments/exceptions.py`
- Modify: `src/docx_comments/manager.py`, `src/docx_comments/xml_parts.py` (get_person), `src/docx_comments/models.py`, `src/docx_comments/__init__.py`
- Test: `tests/test_api_polish.py` (new)

**Acceptance Criteria:**
- [ ] `CommentNotFoundError` subclasses `ValueError` and `LookupError`; `PersonNotFoundError` subclasses `KeyError`; both exported from `docx_comments`
- [ ] Every "Comment ... not found" raise site uses `CommentNotFoundError`; `PeoplePart.get_person` raises `PersonNotFoundError`
- [ ] All comment-id-taking public methods accept `int` (coerced with `str()`); `bool` and other types raise `TypeError`
- [ ] `CommentInfo.comment_id` annotated `Optional[str]` with docstring explaining the id-less-comment case; `para_id` docstring documents the `""` sentinel
- [ ] `reply_to_comment` docstring Raises section lists TypeError and all ValueError variants

**Verify:** `pytest tests/test_api_polish.py -v` → all pass; `mypy src/docx_comments` → clean

**Steps:**

- [ ] **Step 1: Write failing tests** in `tests/test_api_polish.py`:

```python
"""API polish: exception types, int-id coercion, honest type hints."""

import pytest
from docx import Document

from docx_comments import (
    CommentManager,
    CommentNotFoundError,
    PersonInfo,
    PersonNotFoundError,
)


def _doc_with_comment():
    doc = Document()
    para = doc.add_paragraph("some text")
    mgr = CommentManager(doc)
    cid = mgr.add_comment(para, "note", PersonInfo(author="A"), initials="A")
    return doc, mgr, cid


class TestExceptionTypes:
    def test_not_found_is_comment_not_found_error(self):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(CommentNotFoundError):
            mgr.resolve_comment("999999")

    def test_not_found_still_caught_as_value_error(self):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(ValueError):
            mgr.delete_comment("999999")

    def test_not_found_caught_as_lookup_error(self):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(LookupError):
            mgr.delete_thread("999999")

    @pytest.mark.parametrize(
        "op",
        ["delete_comment", "delete_thread", "resolve_comment", "unresolve_comment"],
    )
    def test_all_lookup_ops_raise_comment_not_found(self, op):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(CommentNotFoundError):
            getattr(mgr, op)("424242")

    def test_move_ops_raise_comment_not_found(self):
        doc, mgr, _ = _doc_with_comment()
        target = doc.add_paragraph("target")
        with pytest.raises(CommentNotFoundError):
            mgr.move_comment("424242", target)
        with pytest.raises(CommentNotFoundError):
            mgr.move_thread("424242", target)

    def test_reply_to_missing_parent_raises_comment_not_found(self):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(CommentNotFoundError):
            mgr.reply_to_comment("424242", "r", PersonInfo(author="B"))

    def test_person_not_found_error(self):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(PersonNotFoundError):
            mgr.get_person("Nobody")
        with pytest.raises(KeyError):
            mgr.get_person("Nobody")


class TestIntIdCoercion:
    def test_int_id_accepted_everywhere(self):
        doc, mgr, cid = _doc_with_comment()
        int_id = int(cid)
        rid = mgr.reply_to_comment(int_id, "reply", PersonInfo(author="B"))
        mgr.resolve_comment(int(rid))
        mgr.unresolve_comment(int_id)
        mgr.delete_thread(int_id)
        assert list(mgr.list_comments()) == []

    def test_bool_id_raises_type_error(self):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(TypeError):
            mgr.resolve_comment(True)

    def test_other_types_raise_type_error(self):
        _, mgr, _ = _doc_with_comment()
        with pytest.raises(TypeError):
            mgr.delete_comment(3.14)


class TestHonestHints:
    def test_comment_id_optional_annotation(self):
        from typing import get_type_hints

        from docx_comments.models import CommentInfo

        hints = get_type_hints(CommentInfo)
        assert hints["comment_id"] == type(None).__mro__ and False or str(
            hints["comment_id"]
        ) in ("typing.Optional[str]", "typing.Union[str, NoneType]", "str | None")
```

(Note: the last assertion simplifies to checking `Optional[str]`; write it as:
`assert hints["comment_id"] == Optional[str]` with `from typing import Optional`.)

- [ ] **Step 2:** Run `pytest tests/test_api_polish.py -v` → FAIL (ImportError: CommentNotFoundError).

- [ ] **Step 3: Create** `src/docx_comments/exceptions.py`:

```python
"""Exception types for docx-comments.

Subclassing the previously raised builtins keeps backward compatibility:
existing ``except ValueError`` / ``except KeyError`` consumers keep working,
while new code can catch the precise types.
"""


class CommentNotFoundError(ValueError, LookupError):
    """Raised when a comment id does not match any comment in the document."""


class PersonNotFoundError(KeyError):
    """Raised when an author has no entry in word/people.xml."""
```

- [ ] **Step 4:** In `manager.py`:
  - Add `from docx_comments.exceptions import CommentNotFoundError` to imports.
  - Add module-level helper below `_validate_xml_text`:

```python
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
```

  - Replace `raise ValueError(f"Comment {comment_id} not found")` with `raise CommentNotFoundError(...)` at: `_thread_comments_for` (line ~272), `set_comment_resolved` (~1164), `delete_comment` (~1189), `delete_thread` (~1219), `move_comment` (~1274), `move_thread` (~1327). Replace `raise ValueError(f"Parent comment {parent_id} not found")` (~1027 and ~1042) with `CommentNotFoundError`.
  - At the top of `reply_to_comment`, `set_comment_resolved`, `delete_comment`, `delete_thread`, `move_comment`, `move_thread` add `comment_id = _coerce_comment_id(comment_id)` (for reply: `parent_id = _coerce_comment_id(parent_id)`), and widen the parameter annotations to `Union[int, str]`. (`resolve_comment`/`unresolve_comment` delegate to `set_comment_resolved`, so widening their annotations is enough.)
  - Update `reply_to_comment` docstring Raises section to:

```
        Raises:
            CommentNotFoundError: If the parent comment is not found.
            TypeError: If author is not a PersonInfo (or str once Task 12
                lands) or parent_id is neither str nor int.
            ValueError: If the text/author/initials contain characters not
                allowed in XML, the person spec is invalid, or the parent
                comment has no anchors in the document.
```

- [ ] **Step 5:** In `xml_parts.py`: add `from docx_comments.exceptions import PersonNotFoundError` and change `get_person`'s `raise KeyError(f"person '{author}' not found")` to `raise PersonNotFoundError(f"person '{author}' not found")`.

- [ ] **Step 6:** In `models.py`: change `comment_id: str` to `comment_id: Optional[str]` and its docstring to `"""Unique comment ID (w:id attribute). None for schema-invalid comments that lack w:id (the library still lists them)."""`. Extend `para_id` docstring: `"""Paragraph ID linking to extended/ids parts (w14:paraId). Empty string ("") when the comment has no identifiable paragraph id."""`.

- [ ] **Step 7:** In `__init__.py`: add `from docx_comments.exceptions import CommentNotFoundError, PersonNotFoundError` and extend `__all__` with `"CommentNotFoundError", "PersonNotFoundError"`.

- [ ] **Step 8:** `pytest tests/test_api_polish.py -v` → PASS; `pytest -q` → all pass; `mypy src/docx_comments` clean; `ruff check src/ tests/` clean.

- [ ] **Step 9:** Commit: `git add -A && git commit -m "feat: typed exceptions, int comment-id coercion, honest CommentInfo hints"`

```json:metadata
{"files": ["src/docx_comments/exceptions.py", "src/docx_comments/manager.py", "src/docx_comments/xml_parts.py", "src/docx_comments/models.py", "src/docx_comments/__init__.py", "tests/test_api_polish.py"], "verifyCommand": "pytest tests/test_api_polish.py -v && mypy src/docx_comments", "acceptanceCriteria": ["CommentNotFoundError/PersonNotFoundError exist, subclass legacy types, exported", "int ids coerced at all public entry points", "CommentInfo.comment_id is Optional[str]"], "modelTier": "standard"}
```

---

### Task 2: ST_OnOff done parsing + mc:Ignorable on native comments.xml

**Goal:** Read spec-valid `w15:done` lexical forms correctly, and declare `mc:Ignorable` on comments.xml created through python-docx's native template.

**Files:**
- Modify: `src/docx_comments/xml_parts.py` (get_threading_info, CommentsPart._create_part, new helper), `src/docx_comments/manager.py` (migrate_comment_metadata)
- Test: `tests/test_robustness.py`, `tests/test_xml.py`

**Acceptance Criteria:**
- [ ] `done="true"` / `done="on"` read as resolved; `done="false"` / `done="0"` / garbage as unresolved
- [ ] Saved comments.xml root created via the native python-docx path carries `mc:Ignorable` including `w14`
- [ ] `migrate_comment_metadata` adds `mc:Ignorable` to a foreign comments.xml root (that declares the prefixes) when it backfills w14 attributes

**Verify:** `pytest tests/test_robustness.py tests/test_xml.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests.** In `tests/test_robustness.py` add (reuse the module's existing `qn`/namespace helpers; `NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"`):

```python
class TestDoneLexicalForms:
    @pytest.mark.parametrize(
        "raw,expected",
        [("1", True), ("true", True), ("on", True), ("0", False),
         ("false", False), ("off", False), ("garbage", False), (" TRUE ", True)],
    )
    def test_st_onoff_values(self, raw, expected):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        ext = CommentsExtendedPart(doc)
        for elem in ext.xml:
            if etree.QName(elem).localname == "commentEx":
                elem.set(qn(NS_W15, "done"), raw)
        ext._save()
        comment = next(iter(mgr.list_comments()))
        assert comment.is_resolved is expected
```

  In `tests/test_xml.py` add:

```python
def test_comments_xml_declares_mc_ignorable(tmp_path):
    doc = Document()
    para = doc.add_paragraph("text")
    mgr = CommentManager(doc)
    mgr.add_comment(para, "c", PersonInfo(author="A"))
    path = tmp_path / "out.docx"
    doc.save(str(path))
    with zipfile.ZipFile(path) as zf:
        root = etree.fromstring(zf.read("word/comments.xml"))
    ns_mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    ignorable = root.get(f"{{{ns_mc}}}Ignorable")
    assert ignorable is not None and "w14" in ignorable.split()
```

- [ ] **Step 2:** Run them → FAIL.

- [ ] **Step 3:** In `xml_parts.py`:
  - `get_threading_info`: replace `done = elem.get(_qn(NS_W15, "done"), "0") == "1"` with:

```python
                done_raw = elem.get(_qn(NS_W15, "done"), "0")
                # w15:done is ST_OnOff: "1"/"true"/"on" are all resolved.
                done = done_raw.strip().lower() in ("1", "true", "on")
```

  - Add module-level helper after `sync_part_blob`:

```python
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
```

  - In `CommentsPart._create_part`, native branch, after `part = _NativeCommentsPart.default(...)` and before `relate_to`, add:

```python
            # The native template declares w14/mc but omits mc:Ignorable;
            # this library writes w14:paraId/textId, so declare it ignorable.
            ensure_mc_ignorable(part.element)
```

- [ ] **Step 4:** In `manager.py` `migrate_comment_metadata`, import `ensure_mc_ignorable` from xml_parts and, just before `if updated_comments: self._save_comments()`, add:

```python
        if updated_comments and ensure_mc_ignorable(self._comments_xml):
            pass  # attribute change is persisted by the save below
```

  (Simplify: `if updated_comments: ensure_mc_ignorable(self._comments_xml); self._save_comments()` keeping the existing structure.)

- [ ] **Step 5:** Full suite + mypy + ruff → pass. Commit: `git commit -am "fix: parse ST_OnOff done values; declare mc:Ignorable on native comments.xml"`

```json:metadata
{"files": ["src/docx_comments/xml_parts.py", "src/docx_comments/manager.py", "tests/test_robustness.py", "tests/test_xml.py"], "verifyCommand": "pytest tests/test_robustness.py tests/test_xml.py -q", "acceptanceCriteria": ["done=true/on read as resolved", "saved comments.xml root has mc:Ignorable with w14"], "modelTier": "standard"}
```

---

### Task 3: Block-level anchor reply corruption fix + reference-only-anchor replies

**Goal:** Replying to (or moving a thread onto) a comment whose range markers sit at block level (w:body/w:tr/w:tc) must not emit a schema-invalid bare `w:r` there; comments anchored only by a `commentReference` (no range markers — legal per ECMA-376 §17.13.4) must be reply-able.

**Files:**
- Modify: `src/docx_comments/anchors.py` (`_find_anchor_elements`, `add_anchors_at_comment`, new helpers), `src/docx_comments/manager.py` (reply anchor validation ~line 1059)
- Test: `tests/test_robustness.py` (new class `TestBlockLevelAnchors`)

**Acceptance Criteria:**
- [ ] Reply to a comment with body-level range markers puts the new reference run inside a paragraph (adjacent to the parent's reference run); no `w:r` is a direct child of `w:body`
- [ ] Reply to a comment with tr-level range markers and no parent reference run places the reference run inside the last paragraph within the range; no `w:r` is a direct child of `w:tr`
- [ ] Reply to a reference-only-anchored comment succeeds, synthesizing range markers around the parent's reference run
- [ ] A comment with no anchors at all still raises `ValueError("Could not find anchors ...")` and leaves no orphan reply
- [ ] Existing reply-ordering test (`tests/test_xml.py:172`) still passes

**Verify:** `pytest tests/test_robustness.py::TestBlockLevelAnchors tests/test_xml.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests.** Add to `tests/test_robustness.py` (module already imports `Document`, `CommentManager`, `PersonInfo`, `etree`, `qn`, `NS_W`):

```python
class TestBlockLevelAnchors:
    def _hoist_to_body(self, doc, p_first, p_last):
        body = doc.element.body
        start = body.find(f".//{qn(NS_W, 'commentRangeStart')}")
        end = body.find(f".//{qn(NS_W, 'commentRangeEnd')}")
        p_first._element.addprevious(start)
        p_last._element.addnext(end)

    def test_reply_to_body_level_range_is_schema_valid(self, tmp_path):
        doc = Document()
        p1 = doc.add_paragraph("first")
        p2 = doc.add_paragraph("second")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(p1, "root", PersonInfo(author="A"))
        self._hoist_to_body(doc, p1, p2)
        rid = mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
        body = doc.element.body
        assert all(etree.QName(c).localname != "r" for c in body), (
            "bare w:r under w:body is schema-invalid"
        )
        path = tmp_path / "b.docx"
        doc.save(str(path))
        doc2 = Document(str(path))
        threads = CommentManager(doc2).get_comment_threads()
        assert len(threads) == 1 and threads[0].reply_count == 1
        assert threads[0].replies[0].comment_id == rid

    def test_reply_body_level_without_parent_ref_run(self):
        doc = Document()
        p1 = doc.add_paragraph("first")
        p2 = doc.add_paragraph("second")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(p1, "root", PersonInfo(author="A"))
        self._hoist_to_body(doc, p1, p2)
        # Strip the parent's reference run to force the descend-into-
        # paragraph fallback.
        ref = doc.element.body.find(f".//{qn(NS_W, 'commentReference')}")
        ref_run = ref.getparent()
        ref_run.getparent().remove(ref_run)
        mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
        body = doc.element.body
        assert all(etree.QName(c).localname != "r" for c in body)
        # Reference run landed inside the last paragraph of the range.
        assert p2._element.find(f".//{qn(NS_W, 'commentReference')}") is not None

    def test_reply_to_tr_level_range_is_schema_valid(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=2)
        cell_para = table.cell(0, 0).paragraphs[0]
        cell_para.add_run("cell text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(cell_para, "root", PersonInfo(author="A"))
        tr = table._tbl.tr_lst[0]
        start = tr.find(f".//{qn(NS_W, 'commentRangeStart')}")
        end = tr.find(f".//{qn(NS_W, 'commentRangeEnd')}")
        first_tc = tr.find(qn(NS_W, "tc"))
        first_tc.addprevious(start)
        tr.append(end)
        mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
        assert all(etree.QName(c).localname != "r" for c in tr), (
            "bare w:r under w:tr is schema-invalid"
        )

    def test_reply_to_reference_only_anchor(self, tmp_path):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "root", PersonInfo(author="A"))
        # Strip range markers, keep the reference (legal per ECMA-376).
        body = doc.element.body
        for tag in ("commentRangeStart", "commentRangeEnd"):
            for elem in list(body.iter(qn(NS_W, tag))):
                elem.getparent().remove(elem)
        rid = mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
        path = tmp_path / "r.docx"
        doc.save(str(path))
        doc2 = Document(str(path))
        threads = CommentManager(doc2).get_comment_threads()
        assert len(threads) == 1 and threads[0].reply_count == 1
        # Range markers were synthesized for the reply.
        starts = [
            e for e in doc2.element.body.iter(qn(NS_W, "commentRangeStart"))
            if e.get(qn(NS_W, "id")) == rid
        ]
        assert len(starts) == 1

    def test_reply_with_no_anchors_at_all_still_raises(self):
        from docx_comments.anchors import CommentAnchor

        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "root", PersonInfo(author="A"))
        CommentAnchor(doc).remove_anchors(cid)
        with pytest.raises(ValueError, match="Could not find anchors"):
            mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
        assert len(list(mgr.list_comments())) == 1, "no orphan reply left behind"
```

- [ ] **Step 2:** Run → FAIL (body/tr contain bare `w:r`; ref-only raises).

- [ ] **Step 3:** In `anchors.py`, change `_find_anchor_elements` to fall back to reference-only matches:

```python
        fallback: tuple = (None, None, None, None)
        for part, root in self._iter_anchor_parts():
            start = end = ref = None
            for elem in root.iter(_START_TAG, _END_TAG, _REF_TAG):
                if elem.get(_ID_ATTR) != comment_id:
                    continue
                if elem.tag == _START_TAG and start is None:
                    start = elem
                elif elem.tag == _END_TAG and end is None:
                    end = elem
                elif elem.tag == _REF_TAG and ref is None:
                    ref = elem
            if start is not None and end is not None:
                return part, start, end, ref
            if ref is not None and fallback[3] is None:
                fallback = (part, None, None, ref)

        return fallback
```

  Update its docstring: "Returns (part, None, None, reference) for comments anchored only by a commentReference run; (None, None, None, None) when nothing matches."

- [ ] **Step 4:** Still in `anchors.py`, add two static helpers to `CommentAnchor`:

```python
    @staticmethod
    def _has_paragraph_ancestor(elem: etree._Element) -> bool:
        parent = elem.getparent()
        while parent is not None:
            if etree.QName(parent).localname == "p":
                return True
            parent = parent.getparent()
        return False

    @staticmethod
    def _last_paragraph_before(elem: etree._Element) -> Optional[etree._Element]:
        """Last w:p in document order before `elem`, searching up the tree."""
        p_tag = _qn(NS_W, "p")
        current: Optional[etree._Element] = elem
        while current is not None:
            for sib in current.itersiblings(preceding=True):
                if sib.tag == p_tag:
                    return sib
                inner = sib.findall(f".//{p_tag}")
                if inner:
                    return inner[-1]
            current = current.getparent()
            if current is not None and current.tag == p_tag:
                return current
        return None
```

- [ ] **Step 5:** Rewrite `add_anchors_at_comment` (keep signature). Capture the parent reference instead of discarding it, handle the ref-only case, and place the new reference run at a schema-valid position:

```python
        part, parent_start, parent_end, parent_ref = self._find_anchor_elements(
            parent_comment_id
        )

        if parent_start is None or parent_end is None:
            if parent_ref is None:
                raise ValueError(
                    f"Could not find anchors for comment {parent_comment_id}"
                )
            # Reference-only anchor (range markers are optional per ECMA-376
            # §17.13.4): synthesize a range around the parent's reference run
            # for the new comment, mirroring what Word produces on re-save.
            self._add_anchors_around_reference(parent_ref, new_comment_id)
            sync_part_blob(part)
            return
```

  Keep the existing start/end insertion logic (the two skip-sibling loops for `new_start` and `new_end`) unchanged. Replace the final reference-run block (`ref_run = ...` through `insert_ref_after.addnext(ref_run)`) with:

```python
        # Place the reference run at a schema-valid run position. A bare run
        # directly under w:body/w:tbl/w:tr/w:tc (block-level ranges) makes
        # document.xml invalid and triggers Word's repair prompt.
        ref_run = etree.Element(_qn(NS_W, "r"))
        ref = etree.SubElement(ref_run, _REF_TAG)
        ref.set(_ID_ATTR, new_comment_id)

        anchor_after: Optional[etree._Element] = None
        if parent_ref is not None:
            parent_ref_run = parent_ref.getparent()
            if (
                parent_ref_run is not None
                and etree.QName(parent_ref_run).localname == "r"
            ):
                anchor_after = parent_ref_run
        if anchor_after is None and self._has_paragraph_ancestor(new_end):
            anchor_after = new_end

        if anchor_after is not None:
            insert_ref_after = anchor_after
            sibling = insert_ref_after.getnext()
            while sibling is not None and is_comment_ref_run(sibling):
                insert_ref_after = sibling
                sibling = sibling.getnext()
            insert_ref_after.addnext(ref_run)
        else:
            para = self._last_paragraph_before(new_end)
            if para is None:
                raise ValueError(
                    f"cannot place the comment reference run for comment "
                    f"{new_comment_id}: no paragraph found within the range "
                    f"of comment {parent_comment_id}"
                )
            para.append(ref_run)
```

  Add the new private method:

```python
    def _add_anchors_around_reference(
        self, parent_ref: etree._Element, new_comment_id: str
    ) -> None:
        """Anchor a reply around a parent that has only a reference run."""
        ref_run_parent = parent_ref.getparent()
        target = (
            ref_run_parent
            if ref_run_parent is not None
            and etree.QName(ref_run_parent).localname == "r"
            else parent_ref
        )
        new_start = etree.Element(_START_TAG)
        new_start.set(_ID_ATTR, new_comment_id)
        target.addprevious(new_start)
        new_end = etree.Element(_END_TAG)
        new_end.set(_ID_ATTR, new_comment_id)
        target.addnext(new_end)
        ref_run = etree.Element(_qn(NS_W, "r"))
        ref = etree.SubElement(ref_run, _REF_TAG)
        ref.set(_ID_ATTR, new_comment_id)
        new_end.addnext(ref_run)
```

- [ ] **Step 6:** In `manager.py` `reply_to_comment` (~1059), accept reference-only parents:

```python
        _, parent_start, parent_end, parent_ref = anchor._find_anchor_elements(
            anchor_parent_id
        )
        if (parent_start is None or parent_end is None) and parent_ref is None:
            raise ValueError(f"Could not find anchors for comment {anchor_parent_id}")
```

- [ ] **Step 7:** Full suite + mypy + ruff → pass (pay attention to `tests/test_xml.py::test_reply_anchor_ordering`-style tests — placement for in-paragraph ranges must be unchanged). Commit: `git commit -am "fix: schema-valid reply anchors for block-level ranges; support reference-only anchors"`

```json:metadata
{"files": ["src/docx_comments/anchors.py", "src/docx_comments/manager.py", "tests/test_robustness.py"], "verifyCommand": "pytest tests/test_robustness.py::TestBlockLevelAnchors tests/test_xml.py -q", "acceptanceCriteria": ["no bare w:r under body/tr after replies to block-level ranges", "reference-only parents reply-able with synthesized ranges", "anchor-less parents still raise cleanly"], "modelTier": "frontier"}
```

---

### Task 4: Whole-paragraph anchors include containers; styled reference runs

**Goal:** The index-free default anchor spans ALL visible paragraph content (runs AND hyperlink/ins/del/sdt containers, in document order), and every reference run this library inserts carries Word's `rStyle CommentReference` wrapper.

**Files:**
- Modify: `src/docx_comments/anchors.py` (`_resolve_anchor_span`, new `_make_reference_run`, `add_anchors_at_span`, `_add_anchors_to_empty_paragraph`, `add_anchors_at_comment`, `_add_anchors_around_reference`)
- Test: `tests/test_robustness.py` (new class `TestMixedContentAnchors`)

**Acceptance Criteria:**
- [ ] Default `add_comment` on `[run, hyperlink]`, `[hyperlink, run]`, and `[ins(run), run]` paragraphs places `commentRangeEnd` after the LAST content element (container included)
- [ ] Explicit `start_run`/`end_run` behavior unchanged (still addresses direct runs only)
- [ ] Every inserted reference run has `<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>`
- [ ] `remove_anchors` still fully removes the styled reference runs (existing delete tests pass)

**Verify:** `pytest tests/test_robustness.py -q` → pass

**Steps:**

- [ ] **Step 1: Failing tests** (`tests/test_robustness.py`; the file already has a hyperlink-building pattern near `test_hyperlink_only_paragraph_anchors_whole_content`, line ~279 — reuse its helper if one exists, else use this):

```python
def _append_hyperlink(para, text):
    hyperlink = etree.SubElement(para._element, qn(NS_W, "hyperlink"))
    run = etree.SubElement(hyperlink, qn(NS_W, "r"))
    t = etree.SubElement(run, qn(NS_W, "t"))
    t.text = text
    return hyperlink


class TestMixedContentAnchors:
    def _child_locals(self, para):
        return [etree.QName(c).localname for c in para._element]

    def test_run_then_hyperlink_default_anchor_covers_all(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("See ")
        _append_hyperlink(para, "the link")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        locals_ = self._child_locals(para)
        assert locals_.index("commentRangeEnd") > locals_.index("hyperlink"), (
            f"anchor truncated before the hyperlink: {locals_}"
        )

    def test_hyperlink_then_run_default_anchor_covers_all(self):
        doc = Document()
        para = doc.add_paragraph()
        _append_hyperlink(para, "the link")
        para.add_run(" trailing")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        locals_ = self._child_locals(para)
        assert locals_.index("commentRangeStart") < locals_.index("hyperlink")

    def test_tracked_change_wrapper_covered(self):
        doc = Document()
        para = doc.add_paragraph()
        ins = etree.SubElement(para._element, qn(NS_W, "ins"))
        run = etree.SubElement(ins, qn(NS_W, "r"))
        t = etree.SubElement(run, qn(NS_W, "t"))
        t.text = "inserted"
        para.add_run(" kept")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        locals_ = self._child_locals(para)
        assert locals_.index("commentRangeStart") < locals_.index("ins")

    def test_explicit_indices_still_address_direct_runs(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("one")
        _append_hyperlink(para, "link")
        para.add_run("two")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"), start_run=0, end_run=0)
        children = list(para._element)
        end_idx = next(
            i for i, c in enumerate(children)
            if etree.QName(c).localname == "commentRangeEnd"
        )
        hyper_idx = next(
            i for i, c in enumerate(children)
            if etree.QName(c).localname == "hyperlink"
        )
        assert end_idx < hyper_idx, "explicit run indices must not span containers"

    def test_reference_run_carries_comment_reference_style(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        ref = para._element.find(f".//{qn(NS_W, 'commentReference')}")
        ref_run = ref.getparent()
        style = ref_run.find(f"{qn(NS_W, 'rPr')}/{qn(NS_W, 'rStyle')}")
        assert style is not None and style.get(qn(NS_W, "val")) == "CommentReference"

    def test_styled_reference_run_still_removed_on_delete(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        mgr.delete_comment(cid)
        assert para._element.find(f".//{qn(NS_W, 'commentReference')}") is None
        assert para._element.find(qn(NS_W, "r")) is not None  # text run kept
```

- [ ] **Step 2:** Run → FAIL.

- [ ] **Step 3:** In `anchors.py` `_resolve_anchor_span`, restructure the body:

```python
        runs = para_elem.findall(_qn(NS_W, "r"))
        default_span = start_run == 0 and end_run is None

        if default_span:
            # Index-free default: span all visible content in document
            # order, including runs wrapped in hyperlink/tracked-change/
            # field containers (explicit indices keep addressing direct
            # runs only, unchanged).
            anchorable = [
                child
                for child in para_elem
                if child.tag == _qn(NS_W, "r")
                or etree.QName(child).localname in _RUN_CONTAINER_TAGS
            ]
            if anchorable:
                return anchorable[0], anchorable[-1]
            return None, None

        if not runs:
            raise IndexError(
                "paragraph has no direct runs; omit start_run/end_run to "
                "anchor the whole paragraph"
            )
        # ... keep the existing index-validation logic for the runs case,
        # then `return runs[start_run], runs[end_run]`.
```

  Delete the old zero-runs container fallback block (subsumed by the default branch).

- [ ] **Step 4:** Add a module-level factory in `anchors.py` and use it at every reference-run construction site (`add_anchors_at_span`, `_add_anchors_to_empty_paragraph`, `add_anchors_at_comment`, `_add_anchors_around_reference`):

```python
def _make_reference_run(comment_id: str) -> etree._Element:
    """Build Word's styled comment reference run."""
    ref_run = etree.Element(_qn(NS_W, "r"))
    rpr = etree.SubElement(ref_run, _qn(NS_W, "rPr"))
    rstyle = etree.SubElement(rpr, _qn(NS_W, "rStyle"))
    rstyle.set(_qn(NS_W, "val"), "CommentReference")
    ref = etree.SubElement(ref_run, _REF_TAG)
    ref.set(_ID_ATTR, comment_id)
    return ref_run
```

  Replace the three-line `ref_run = etree.Element(...); ref = etree.SubElement(...); ref.set(...)` groups with `ref_run = _make_reference_run(comment_id)` (or `new_comment_id`). `remove_anchors` and `ensure_span_survives_removal` already tolerate the rPr wrapper — do not change them.

- [ ] **Step 5:** Full suite + mypy + ruff. Commit: `git commit -am "fix: default anchors span run containers; style reference runs like Word"`

```json:metadata
{"files": ["src/docx_comments/anchors.py", "tests/test_robustness.py"], "verifyCommand": "pytest tests/test_robustness.py -q", "acceptanceCriteria": ["default span covers containers in document order", "explicit indices unchanged", "reference runs styled CommentReference and still removable"], "modelTier": "standard"}
```

---

### Task 5: Replies inherit the thread's resolution state

**Goal:** `reply_to_comment` into a resolved thread writes `done="1"` for the reply instead of creating a mixed thread state Word never produces.

**Files:**
- Modify: `src/docx_comments/manager.py` (reply_to_comment, ~line 1102)
- Test: `tests/test_threads.py`

**Acceptance Criteria:**
- [ ] After add → resolve → reply: every member of the thread reports `is_resolved=True` and `CommentThread.is_resolved` is True
- [ ] Reply into an unresolved thread stays `done=False`

**Verify:** `pytest tests/test_threads.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing test** in `tests/test_threads.py`:

```python
def test_reply_into_resolved_thread_inherits_done():
    doc = Document()
    para = doc.add_paragraph("text")
    mgr = CommentManager(doc)
    cid = mgr.add_comment(para, "root", PersonInfo(author="A"))
    mgr.resolve_comment(cid)
    mgr.reply_to_comment(cid, "late reply", PersonInfo(author="B"))
    comments = list(mgr.list_comments())
    assert all(c.is_resolved for c in comments), (
        "reply into a resolved thread must inherit done=1"
    )
    thread = mgr.get_comment_threads()[0]
    assert thread.is_resolved


def test_reply_into_open_thread_stays_unresolved():
    doc = Document()
    para = doc.add_paragraph("text")
    mgr = CommentManager(doc)
    cid = mgr.add_comment(para, "root", PersonInfo(author="A"))
    mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
    assert not any(c.is_resolved for c in mgr.list_comments())
```

- [ ] **Step 2:** Run → first test FAILS.

- [ ] **Step 3:** In `reply_to_comment`, the threading dict is already fetched (`threading = ext_part.get_threading_info()`). Change the final `add_comment_ex` call:

```python
        # The reply inherits the thread's current resolution state so a
        # resolved thread stays consistently resolved (Word keeps all
        # members' done flags in sync).
        inherited_done = threading.get(effective_parent_para_id, {}).get("done", False)
        ext_part.add_comment_ex(
            para_id=para_id,
            parent_para_id=effective_parent_para_id,
            done=inherited_done,
        )
```

- [ ] **Step 4:** Full suite + commit: `git commit -am "fix: replies inherit the thread's done state"`

```json:metadata
{"files": ["src/docx_comments/manager.py", "tests/test_threads.py"], "verifyCommand": "pytest tests/test_threads.py -q", "acceptanceCriteria": ["reply into resolved thread is done=1", "reply into open thread is done=0"], "modelTier": "mechanical"}
```

---

### Task 6: XML-legality validation for person identity; atomic people.xml writes

**Goal:** XML-illegal person/presence strings raise a clear ValueError BEFORE any part creation or tree mutation, and `ensure_person` can never leave a half-built `w15:person` in the shared cached tree.

**Files:**
- Modify: `src/docx_comments/xml_parts.py` (new `validate_xml_text`, rewrite `PeoplePart.ensure_person`), `src/docx_comments/manager.py` (`_validate_xml_text` becomes an import alias; `_resolve_person_spec` validates values)
- Test: `tests/test_people.py`

**Acceptance Criteria:**
- [ ] `add_comment(..., person={"provider_id": "AD\x00", "user_id": "u"})` raises ValueError mentioning the field, and the saved document contains NO comment parts and NO people.xml
- [ ] `ensure_person("bad\x00author")` raises ValueError and leaves people.xml absent/unchanged (no ghost `<w15:person>` on the next save)
- [ ] `manager._validate_xml_text` still importable (alias) so nothing else breaks

**Verify:** `pytest tests/test_people.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_people.py`:

```python
class TestPersonXmlLegality:
    def test_illegal_presence_raises_before_any_mutation(self, tmp_path):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        with pytest.raises(ValueError, match="not allowed in XML"):
            mgr.add_comment(
                para,
                "c",
                PersonInfo(author="A"),
                person={"provider_id": "AD\x00", "user_id": "u"},
            )
        path = tmp_path / "clean.docx"
        doc.save(str(path))
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
        assert "word/comments.xml" not in names
        assert "word/people.xml" not in names

    def test_illegal_author_in_ensure_person_no_ghost_entry(self, tmp_path):
        doc = Document()
        mgr = CommentManager(doc)
        with pytest.raises(ValueError, match="not allowed in XML"):
            mgr.ensure_person("bad\x00author")
        assert mgr.get_people() == []
        # A later valid write must not resurrect a half-built entry.
        mgr.ensure_person("Good Author")
        path = tmp_path / "p.docx"
        doc.save(str(path))
        doc2 = Document(str(path))
        people = CommentManager(doc2).get_people()
        assert [p.author for p in people] == ["Good Author"]
```

  (Add `import zipfile` at the top of the file if missing.)

- [ ] **Step 2:** Run → FAIL (parts get created; ghost entry persists).

- [ ] **Step 3:** In `xml_parts.py`, add module-level (after `parse_xml_bytes`):

```python
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
```

- [ ] **Step 4:** In `manager.py`, delete the local `_validate_xml_text` definition and add to the xml_parts import block: `validate_xml_text`, then alias below the imports: `_validate_xml_text = validate_xml_text` (keeps `docx_comments.manager._validate_xml_text` importable).

- [ ] **Step 5:** In `manager._resolve_person_spec`, replace the trailing presence check:

```python
        _validate_xml_text(person_author, "person author")
        if presence is not None:
            provider_id, user_id = PeoplePart._normalize_presence(presence)
            _validate_xml_text(provider_id, "presence provider_id")
            _validate_xml_text(user_id, "presence user_id")
```

- [ ] **Step 6:** Rewrite `PeoplePart.ensure_person` to validate up front and build detached:

```python
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
                presence_elem = etree.SubElement(
                    new_elem, _qn(NS_W15, "presenceInfo")
                )
                presence_elem.set(_qn(NS_W15, "providerId"), normalized[0])
                presence_elem.set(_qn(NS_W15, "userId"), normalized[1])
            self.ensure_exists()
            self.xml.append(new_elem)
            person_elem = new_elem
        elif normalized:
            presence_elem = self._find_child_by_localname(person_elem, "presenceInfo")
            if presence_elem is None:
                presence_elem = etree.SubElement(
                    person_elem, _qn(NS_W15, "presenceInfo")
                )
            presence_elem.set(_qn(NS_W15, "providerId"), normalized[0])
            presence_elem.set(_qn(NS_W15, "userId"), normalized[1])

        self._save()
        return self._person_info_from_elem(person_elem)
```

- [ ] **Step 7:** Full suite + mypy + ruff. Commit: `git commit -am "fix: validate person identity XML-legality before mutation; atomic people.xml writes"`

```json:metadata
{"files": ["src/docx_comments/xml_parts.py", "src/docx_comments/manager.py", "tests/test_people.py"], "verifyCommand": "pytest tests/test_people.py -q", "acceptanceCriteria": ["illegal presence raises before any part creation", "no ghost person entries survive a failed ensure_person"], "modelTier": "standard"}
```

---

### Task 7: Read-side API — get_comment, get_thread, get_comment_paragraph, get_anchored_text

**Goal:** Public answers to "what does this comment annotate, and where?" plus single-comment/thread lookup.

**Files:**
- Modify: `src/docx_comments/manager.py`, `src/docx_comments/anchors.py` (new `get_anchored_text`, `_in_fallback` helper; import NS_MC), `src/docx_comments/__init__.py`
- Test: `tests/test_read_api.py` (new)

**Acceptance Criteria:**
- [ ] `get_comment(id) -> CommentInfo` and `get_thread(id) -> CommentThread` work with str and int ids and raise `CommentNotFoundError` otherwise
- [ ] `get_comment_paragraph(id)` returns the anchor paragraph (or None for anchor-less comments); raises `CommentNotFoundError` for unknown ids
- [ ] `get_anchored_text(id)` returns the exact text between the range markers: run-range subsets, container content, `\n` for br/cr and paragraph boundaries, `\t` for tabs; returns None for reference-only/anchor-less comments
- [ ] All four exported via `CommentManager`; `CommentAnchor` NOT added to `__all__` (the manager wrappers are the public path)

**Verify:** `pytest tests/test_read_api.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_read_api.py`:

```python
"""Read-side API: get_comment, get_thread, anchors introspection."""

import pytest
from docx import Document
from lxml import etree

from docx_comments import CommentManager, CommentNotFoundError, PersonInfo

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(ns, name):
    return f"{{{ns}}}{name}"


def _mgr_with_comment(text="hello world"):
    doc = Document()
    para = doc.add_paragraph(text)
    mgr = CommentManager(doc)
    cid = mgr.add_comment(para, "note", PersonInfo(author="A"))
    return doc, para, mgr, cid


class TestGetCommentAndThread:
    def test_get_comment(self):
        _, _, mgr, cid = _mgr_with_comment()
        info = mgr.get_comment(cid)
        assert info.comment_id == cid and info.text == "note"

    def test_get_comment_int_id(self):
        _, _, mgr, cid = _mgr_with_comment()
        assert mgr.get_comment(int(cid)).comment_id == cid

    def test_get_comment_not_found(self):
        _, _, mgr, _ = _mgr_with_comment()
        with pytest.raises(CommentNotFoundError):
            mgr.get_comment("999999")

    def test_get_thread_by_reply_id(self):
        _, _, mgr, cid = _mgr_with_comment()
        rid = mgr.reply_to_comment(cid, "r", PersonInfo(author="B"))
        thread = mgr.get_thread(rid)
        assert thread.root.comment_id == cid
        assert [c.comment_id for c in thread.replies] == [rid]

    def test_get_thread_not_found(self):
        _, _, mgr, _ = _mgr_with_comment()
        with pytest.raises(CommentNotFoundError):
            mgr.get_thread("999999")


class TestAnchorIntrospection:
    def test_get_comment_paragraph(self):
        _, para, mgr, cid = _mgr_with_comment()
        found = mgr.get_comment_paragraph(cid)
        assert found is not None and found._element is para._element

    def test_get_comment_paragraph_unknown_id(self):
        _, _, mgr, _ = _mgr_with_comment()
        with pytest.raises(CommentNotFoundError):
            mgr.get_comment_paragraph("999999")

    def test_anchored_text_whole_paragraph(self):
        _, _, mgr, cid = _mgr_with_comment("hello world")
        assert mgr.get_anchored_text(cid) == "hello world"

    def test_anchored_text_run_range(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("one ")
        para.add_run("two ")
        para.add_run("three")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(
            para, "c", PersonInfo(author="A"), start_run=1, end_run=1
        )
        assert mgr.get_anchored_text(cid) == "two "

    def test_anchored_text_includes_container_content(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("See ")
        hyperlink = etree.SubElement(para._element, qn(NS_W, "hyperlink"))
        run = etree.SubElement(hyperlink, qn(NS_W, "r"))
        t = etree.SubElement(run, qn(NS_W, "t"))
        t.text = "the link"
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        assert mgr.get_anchored_text(cid) == "See the link"

    def test_anchored_text_multi_paragraph_block_range(self):
        doc = Document()
        p1 = doc.add_paragraph("first")
        p2 = doc.add_paragraph("second")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(p1, "c", PersonInfo(author="A"))
        body = doc.element.body
        start = body.find(f".//{qn(NS_W, 'commentRangeStart')}")
        end = body.find(f".//{qn(NS_W, 'commentRangeEnd')}")
        p1._element.addprevious(start)
        p2._element.addnext(end)
        assert mgr.get_anchored_text(cid) == "first\nsecond"

    def test_anchored_text_empty_paragraph(self):
        doc = Document()
        para = doc.add_paragraph("")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        assert mgr.get_anchored_text(cid) == ""

    def test_anchored_text_none_without_range(self):
        from docx_comments.anchors import CommentAnchor

        doc, _, mgr, cid = (*_mgr_with_comment(),)
        body = doc.element.body
        for tag in ("commentRangeStart", "commentRangeEnd"):
            for elem in list(body.iter(qn(NS_W, tag))):
                elem.getparent().remove(elem)
        assert mgr.get_anchored_text(cid) is None
```

  (Fix the tuple-unpack in the last test to `doc, _, mgr, cid = _mgr_with_comment()`.)

- [ ] **Step 2:** Run → FAIL (AttributeError).

- [ ] **Step 3:** In `anchors.py`: import NS_MC (`NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"` at module top, near NS_W) and add to `CommentAnchor`:

```python
    @staticmethod
    def _in_fallback(elem: etree._Element) -> bool:
        parent = elem.getparent()
        fallback = _qn(NS_MC, "Fallback")
        while parent is not None:
            if parent.tag == fallback:
                return True
            parent = parent.getparent()
        return False

    def get_anchored_text(self, comment_id: str) -> Optional[str]:
        """Text between a comment's range markers, in document order.

        Mirrors the extraction rules of comment-text reading: w:t text,
        w:br/w:cr as newline, w:tab as tab, mc:Fallback content skipped,
        and a newline when the range crosses a paragraph boundary.
        Returns None when the comment has no commentRangeStart/End pair.
        """
        _, start, end, _ = self._find_anchor_elements(comment_id)
        if start is None or end is None:
            return None
        root = start.getroottree().getroot()
        p_tag = _qn(NS_W, "p")
        r_tag = _qn(NS_W, "r")
        start_in_paragraph = self._has_paragraph_ancestor(start)
        pieces: list[str] = []
        active = False
        paras_seen = 0
        for elem in root.iter():
            if elem is start:
                active = True
                continue
            if elem is end:
                break
            if not active:
                continue
            if elem.tag == p_tag:
                if paras_seen > 0 or start_in_paragraph:
                    pieces.append("\n")
                paras_seen += 1
                continue
            parent = elem.getparent()
            if parent is None or parent.tag != r_tag:
                continue  # e.g. w:tab inside w:pPr/w:tabs
            if self._in_fallback(elem):
                continue
            local = etree.QName(elem).localname
            if local == "t":
                if elem.text:
                    pieces.append(elem.text)
            elif local in ("br", "cr"):
                pieces.append("\n")
            elif local == "tab":
                pieces.append("\t")
        return "".join(pieces)
```

- [ ] **Step 4:** In `manager.py` add four public methods after `get_comment_threads` (all coerce ids and import nothing new beyond what Task 1 added):

```python
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
```

  `Paragraph` is already imported under TYPE_CHECKING — since annotations are strings (`from __future__ import annotations`), no runtime import is needed.

- [ ] **Step 5:** Full suite + mypy + ruff. Commit: `git commit -am "feat: get_comment, get_thread, get_comment_paragraph, get_anchored_text"`

```json:metadata
{"files": ["src/docx_comments/manager.py", "src/docx_comments/anchors.py", "tests/test_read_api.py"], "verifyCommand": "pytest tests/test_read_api.py -q", "acceptanceCriteria": ["single lookup methods with CommentNotFoundError", "anchored text exact for run ranges, containers, multi-paragraph ranges", "None for anchor-less comments"], "modelTier": "standard"}
```

---

### Task 8: Caller-controlled timestamps + date parsing tests

**Goal:** `add_comment`/`reply_to_comment` accept `timestamp:` (tz-aware or naive → local), and the date-parsing/round-trip behavior is pinned by tests.

**Files:**
- Modify: `src/docx_comments/manager.py` (add_comment, reply_to_comment, _add_comment_xml)
- Test: `tests/test_timestamps.py` (new)

**Acceptance Criteria:**
- [ ] `add_comment(..., timestamp=datetime(2020, 1, 2, 3, 4, 5, tzinfo=timezone.utc))` round-trips through save/reload equal to the second, and `commentsExtensible` dateUtc == `2020-01-02T03:04:05Z`
- [ ] Naive timestamps are interpreted as local time
- [ ] `_parse_comment_date` unit-tested for Z-form, offset-form, naive (→UTC), and garbage (→None)
- [ ] `get_comment_threads()` survives a comment with corrupted `w:date`

**Verify:** `pytest tests/test_timestamps.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_timestamps.py`:

```python
"""Timestamp control and date parsing."""

import zipfile
from datetime import datetime, timezone

from docx import Document
from lxml import etree

from docx_comments import CommentManager, PersonInfo
from docx_comments.manager import _parse_comment_date

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"


class TestParseCommentDate:
    def test_z_form(self):
        parsed = _parse_comment_date("2020-01-02T03:04:05Z")
        assert parsed == datetime(2020, 1, 2, 3, 4, 5, tzinfo=timezone.utc)

    def test_offset_form(self):
        parsed = _parse_comment_date("2020-01-02T05:04:05+02:00")
        assert parsed == datetime(2020, 1, 2, 3, 4, 5, tzinfo=timezone.utc)

    def test_naive_assumed_utc(self):
        parsed = _parse_comment_date("2020-01-02T03:04:05")
        assert parsed.tzinfo == timezone.utc

    def test_garbage_returns_none(self):
        assert _parse_comment_date("not-a-date") is None

    def test_empty_returns_none(self):
        assert _parse_comment_date(None) is None
        assert _parse_comment_date("") is None


class TestTimestampParameter:
    def test_round_trip_and_date_utc(self, tmp_path):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        stamp = datetime(2020, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
        mgr.add_comment(para, "c", PersonInfo(author="A"), timestamp=stamp)
        path = tmp_path / "t.docx"
        doc.save(str(path))
        doc2 = Document(str(path))
        comment = next(iter(CommentManager(doc2).list_comments()))
        assert comment.timestamp == stamp
        with zipfile.ZipFile(path) as zf:
            ext = etree.fromstring(zf.read("word/commentsExtensible.xml"))
        entries = [
            e.get(f"{{{NS_W16CEX}}}dateUtc")
            for e in ext
            if etree.QName(e).localname == "commentExtensible"
        ]
        assert entries == ["2020-01-02T03:04:05Z"]

    def test_naive_timestamp_treated_as_local(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        naive = datetime(2020, 6, 1, 12, 0, 0)
        mgr.add_comment(para, "c", PersonInfo(author="A"), timestamp=naive)
        comment = next(iter(mgr.list_comments()))
        assert comment.timestamp == naive.astimezone(timezone.utc)

    def test_reply_timestamp_orders_replies(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "root", PersonInfo(author="A"))
        t1 = datetime(2020, 1, 1, tzinfo=timezone.utc)
        t2 = datetime(2020, 1, 2, tzinfo=timezone.utc)
        r_late = mgr.reply_to_comment(cid, "late", PersonInfo(author="B"), timestamp=t2)
        r_early = mgr.reply_to_comment(
            cid, "early", PersonInfo(author="B"), timestamp=t1
        )
        thread = mgr.get_comment_threads()[0]
        assert [r.comment_id for r in thread.replies] == [r_early, r_late]

    def test_threads_survive_corrupted_date(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        for elem in mgr._comments_xml.findall(f"{{{NS_W}}}comment"):
            elem.set(f"{{{NS_W}}}date", "garbage")
        threads = mgr.get_comment_threads()
        assert len(threads) == 1
        assert threads[0].root.timestamp is None
```

- [ ] **Step 2:** Run → FAIL (unexpected keyword `timestamp`).

- [ ] **Step 3:** In `manager.py`:
  - `add_comment` and `reply_to_comment`: add parameter `timestamp: Optional[datetime] = None` (after `person`), document it (`"""timestamp: Optional creation time. Naive datetimes are interpreted as local time; None uses the current time."""`), and pass `timestamp=timestamp` to `self._add_comment_xml(...)`.
  - `_add_comment_xml`: replace the default block with:

```python
        if timestamp is None:
            timestamp = datetime.now().astimezone()
        elif timestamp.tzinfo is None:
            # Interpret naive datetimes as local time (matching the default).
            timestamp = timestamp.astimezone()
```

- [ ] **Step 4:** Full suite + commit: `git commit -am "feat: caller-controlled timestamps on add/reply; pin date parsing behavior"`

```json:metadata
{"files": ["src/docx_comments/manager.py", "tests/test_timestamps.py"], "verifyCommand": "pytest tests/test_timestamps.py -q", "acceptanceCriteria": ["timestamp param round-trips and feeds dateUtc", "naive interpreted as local", "date parsing unit-tested incl. garbage"], "modelTier": "mechanical"}
```

---

### Task 9: edit_comment

**Goal:** In-place comment editing that preserves comment id, primary paraId, durable id, threading, resolution, and anchors; optional author/initials/timestamp updates.

**Files:**
- Modify: `src/docx_comments/manager.py` (new `edit_comment`, extract `_primary_para_id` helper, refactor `list_comments` to use it), `src/docx_comments/xml_parts.py` (new `CommentsExtensiblePart.set_date_utc`)
- Test: `tests/test_editing.py`

**Acceptance Criteria:**
- [ ] `edit_comment(cid, "new text")` changes only the text: comment_id/para_id/durable_id/is_resolved/parent links identical before and after; replies stay attached; anchors untouched
- [ ] `w14:textId` changes (text revision id); optional `timestamp=` updates both `w:date` and the `dateUtc` in commentsExtensible
- [ ] Optional `author=`/`initials=` update the attributes; all validation (XML-legality, non-empty author) happens before any mutation
- [ ] Unknown id raises `CommentNotFoundError`; multi-paragraph comments collapse to one paragraph keyed to the surviving primary paraId with orphan metadata cleaned

**Verify:** `pytest tests/test_editing.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_editing.py`:

```python
class TestEditComment:
    def _setup(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "original", PersonInfo(author="A"), initials="A")
        return doc, mgr, cid

    def test_edit_preserves_identity_and_thread(self, tmp_path):
        doc, mgr, cid = self._setup()
        rid = mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
        mgr.resolve_comment(cid)
        before = mgr.get_comment(cid)
        mgr.edit_comment(cid, "corrected")
        after = mgr.get_comment(cid)
        assert after.text == "corrected"
        assert after.comment_id == before.comment_id
        assert after.para_id == before.para_id
        assert after.durable_id == before.durable_id
        assert after.is_resolved is True
        path = tmp_path / "e.docx"
        doc.save(str(path))
        threads = CommentManager(Document(str(path))).get_comment_threads()
        assert threads[0].root.text == "corrected"
        assert [r.comment_id for r in threads[0].replies] == [rid]

    def test_edit_changes_text_id(self):
        NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
        doc, mgr, cid = self._setup()
        para = mgr._comments_xml.find(f".//{qn(NS_W, 'p')}")
        old_text_id = para.get(qn(NS_W14, "textId"))
        mgr.edit_comment(cid, "new")
        para = mgr._comments_xml.find(f".//{qn(NS_W, 'p')}")
        assert para.get(qn(NS_W14, "textId")) != old_text_id

    def test_edit_timestamp_updates_date_and_date_utc(self):
        from datetime import datetime, timezone

        doc, mgr, cid = self._setup()
        stamp = datetime(2021, 5, 6, 7, 8, 9, tzinfo=timezone.utc)
        mgr.edit_comment(cid, "new", timestamp=stamp)
        info = mgr.get_comment(cid)
        assert info.timestamp == stamp
        from docx_comments.xml_parts import CommentsExtensiblePart

        ext = CommentsExtensiblePart(doc).get_extensible_info()
        assert ext[info.durable_id]["date_utc"] == "2021-05-06T07:08:09Z"

    def test_edit_author_and_initials(self):
        doc, mgr, cid = self._setup()
        mgr.edit_comment(cid, "new", author="B Author", initials="BA")
        info = mgr.get_comment(cid)
        assert info.author == "B Author" and info.initials == "BA"

    def test_edit_unknown_id(self):
        from docx_comments import CommentNotFoundError

        _, mgr, _ = self._setup()
        with pytest.raises(CommentNotFoundError):
            mgr.edit_comment("999999", "x")

    def test_edit_illegal_text_no_mutation(self):
        _, mgr, cid = self._setup()
        with pytest.raises(ValueError, match="not allowed in XML"):
            mgr.edit_comment(cid, "bad\x00text")
        assert mgr.get_comment(cid).text == "original"

    def test_edit_anchors_untouched(self):
        doc, mgr, cid = self._setup()
        before = etree.tostring(doc.element.body)
        mgr.edit_comment(cid, "new")
        assert etree.tostring(doc.element.body) == before
```

  (Import `qn`, `NS_W`, `etree` consistently with the file's existing imports.)

- [ ] **Step 2:** Run → FAIL.

- [ ] **Step 3:** In `xml_parts.py` add to `CommentsExtensiblePart`:

```python
    def set_date_utc(self, durable_id: str, date_utc: str) -> bool:
        """Overwrite the dateUtc for an existing entry.

        Returns True when an entry was updated.
        """
        updated = False
        for elem in self.xml:
            if (
                etree.QName(elem).localname == "commentExtensible"
                and elem.get(_qn(NS_W16CEX, "durableId")) == durable_id
            ):
                elem.set(_qn(NS_W16CEX, "dateUtc"), date_utc)
                updated = True
        if updated:
            self._save()
        return updated
```

- [ ] **Step 4:** In `manager.py` extract the primary-paraId selection into a static helper and use it in `list_comments` (replacing the inline reversed loops at ~lines 573-584) and in `edit_comment`:

```python
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
```

- [ ] **Step 5:** Add `edit_comment` after `delete_thread`:

```python
    def edit_comment(
        self,
        comment_id: Union[int, str],
        text: str,
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
            text: New comment text (same encoding rules as add_comment).
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
        _validate_xml_text(text, "comment text")
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

        new_para = etree.Element(_qn(NS_W, "p"))
        new_para.set(_qn(NS_W, "rsidR"), rsid_r)
        new_para.set(_qn(NS_W, "rsidRDefault"), rsid_default)
        new_para.set(_qn(NS_W14, "paraId"), primary)
        new_para.set(_qn(NS_W14, "textId"), text_id)
        pPr = etree.SubElement(new_para, _qn(NS_W, "pPr"))
        pStyle = etree.SubElement(pPr, _qn(NS_W, "pStyle"))
        pStyle.set(_qn(NS_W, "val"), "CommentText")
        run1 = etree.SubElement(new_para, _qn(NS_W, "r"))
        rPr = etree.SubElement(run1, _qn(NS_W, "rPr"))
        rStyle = etree.SubElement(rPr, _qn(NS_W, "rStyle"))
        rStyle.set(_qn(NS_W, "val"), "CommentReference")
        etree.SubElement(run1, _qn(NS_W, "annotationRef"))
        self._append_text_content(new_para, text, rsid_rpr)

        for para in comment_elem.findall(_qn(NS_W, "p")):
            comment_elem.remove(para)
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

        # Multi-paragraph comments collapse to one paragraph; metadata for
        # the dropped paragraphs (if any existed) is now orphaned.
        dropped = {pid for pid in para_ids if pid != primary}
        self._cleanup_comment_metadata(dropped)

        self._save_comments()
```

- [ ] **Step 6:** Full suite + mypy + ruff. Commit: `git commit -am "feat: edit_comment preserving identity, threading, and anchors"`

```json:metadata
{"files": ["src/docx_comments/manager.py", "src/docx_comments/xml_parts.py", "tests/test_editing.py"], "verifyCommand": "pytest tests/test_editing.py -q", "acceptanceCriteria": ["identity/thread/anchors preserved across edit", "textId refreshed", "timestamp updates date + dateUtc", "validation before mutation"], "modelTier": "standard"}
```

---

### Task 10: Rich comment content — formatted runs and multiple paragraphs

**Goal:** `add_comment`/`reply_to_comment`/`edit_comment` accept structured content: a plain str (unchanged), or a sequence of paragraphs where each paragraph is a str or a sequence of runs `(text, {"bold": True, "italic": True, "underline": True})`.

**Files:**
- Modify: `src/docx_comments/models.py` (type aliases), `src/docx_comments/manager.py` (`_normalize_content`, `_append_text_content` fmt support, `_add_comment_xml` multi-paragraph, signatures)
- Test: `tests/test_rich_content.py` (new)

**Acceptance Criteria:**
- [ ] `text="plain"` behaves byte-identically to today (single paragraph, \n→w:br)
- [ ] `text=[[("bold", {"bold": True}), " rest"]]` emits `<w:rPr><w:b/></w:rPr>` on the first run only
- [ ] `text=["para one", "para two"]` emits two `w:p` elements, each with fresh paraId/textId; satellite metadata (commentsExtended/Ids/Extensible) is keyed to the LAST paragraph's paraId; list_comments returns `"para one\npara two"`
- [ ] Unknown format keys raise ValueError; XML-illegal run text raises before mutation; empty content sequence raises
- [ ] Works identically through reply_to_comment and edit_comment

**Verify:** `pytest tests/test_rich_content.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_rich_content.py`:

```python
"""Rich comment content: formatted runs and multiple paragraphs."""

import pytest
from docx import Document
from lxml import etree

from docx_comments import CommentManager, PersonInfo
from docx_comments.xml_parts import CommentsExtendedPart

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def qn(ns, name):
    return f"{{{ns}}}{name}"


def _setup():
    doc = Document()
    para = doc.add_paragraph("text")
    return doc, para, CommentManager(doc)


class TestFormattedRuns:
    def test_bold_italic_underline(self):
        doc, para, mgr = _setup()
        mgr.add_comment(
            para,
            [[("b", {"bold": True}), ("i", {"italic": True}),
              ("u", {"underline": True}), " plain"]],
            PersonInfo(author="A"),
        )
        comment = doc.element.getroottree().getroot()  # placeholder; use mgr
        comment = mgr._comments_xml.find(qn(NS_W, "comment"))
        runs = [
            r for r in comment.iter(qn(NS_W, "r"))
            if r.find(qn(NS_W, "annotationRef")) is None
        ]
        assert runs[0].find(f"{qn(NS_W, 'rPr')}/{qn(NS_W, 'b')}") is not None
        assert runs[1].find(f"{qn(NS_W, 'rPr')}/{qn(NS_W, 'i')}") is not None
        u = runs[2].find(f"{qn(NS_W, 'rPr')}/{qn(NS_W, 'u')}")
        assert u is not None and u.get(qn(NS_W, "val")) == "single"
        assert runs[3].find(qn(NS_W, "rPr")) is None

    def test_unknown_format_key_raises(self):
        _, para, mgr = _setup()
        with pytest.raises(ValueError, match="unsupported run formatting"):
            mgr.add_comment(para, [[("x", {"blink": True})]], PersonInfo(author="A"))

    def test_plain_str_unchanged(self):
        doc, para, mgr = _setup()
        cid = mgr.add_comment(para, "line1\nline2", PersonInfo(author="A"))
        assert mgr.get_comment(cid).text == "line1\nline2"
        comment = mgr._comments_xml.find(qn(NS_W, "comment"))
        assert len(comment.findall(qn(NS_W, "p"))) == 1


class TestMultiParagraph:
    def test_two_paragraphs_keyed_to_last(self, tmp_path):
        doc, para, mgr = _setup()
        cid = mgr.add_comment(para, ["para one", "para two"], PersonInfo(author="A"))
        comment = mgr._comments_xml.find(qn(NS_W, "comment"))
        paras = comment.findall(qn(NS_W, "p"))
        assert len(paras) == 2
        pids = [p.get(qn(NS_W14, "paraId")) for p in paras]
        assert all(pids) and pids[0] != pids[1]
        threading = CommentsExtendedPart(doc).get_threading_info()
        assert pids[1] in threading and pids[0] not in threading
        assert mgr.get_comment(cid).text == "para one\npara two"
        path = tmp_path / "m.docx"
        doc.save(str(path))
        reread = CommentManager(Document(str(path))).get_comment(cid)
        assert reread.text == "para one\npara two"
        assert reread.para_id == pids[1]

    def test_annotation_ref_on_first_paragraph_only(self):
        doc, para, mgr = _setup()
        mgr.add_comment(para, ["one", "two"], PersonInfo(author="A"))
        comment = mgr._comments_xml.find(qn(NS_W, "comment"))
        paras = comment.findall(qn(NS_W, "p"))
        assert paras[0].find(f".//{qn(NS_W, 'annotationRef')}") is not None
        assert paras[1].find(f".//{qn(NS_W, 'annotationRef')}") is None

    def test_reply_and_edit_accept_rich_content(self):
        doc, para, mgr = _setup()
        cid = mgr.add_comment(para, "root", PersonInfo(author="A"))
        rid = mgr.reply_to_comment(cid, [[("hot", {"bold": True})]], PersonInfo(author="B"))
        assert mgr.get_comment(rid).text == "hot"
        mgr.edit_comment(cid, ["new one", "new two"])
        assert mgr.get_comment(cid).text == "new one\nnew two"

    def test_empty_content_raises(self):
        _, para, mgr = _setup()
        with pytest.raises(ValueError, match="at least one paragraph"):
            mgr.add_comment(para, [], PersonInfo(author="A"))

    def test_illegal_char_in_run_no_mutation(self, tmp_path):
        doc, para, mgr = _setup()
        with pytest.raises(ValueError, match="not allowed in XML"):
            mgr.add_comment(para, [["ok", ("bad\x00", {})]], PersonInfo(author="A"))
        path = tmp_path / "c.docx"
        doc.save(str(path))
        import zipfile

        with zipfile.ZipFile(path) as zf:
            assert "word/comments.xml" not in zf.namelist()
```

  (Remove the stray placeholder line in the first test.)

- [ ] **Step 2:** Run → FAIL.

- [ ] **Step 3:** In `models.py` add (with `from typing import Dict, Sequence, Tuple, Union` — keep 3.9 compatibility):

```python
RunSpec = Union[str, Tuple[str, Dict[str, bool]]]
"""A comment run: plain text, or (text, {"bold"/"italic"/"underline": True})."""

ParagraphSpec = Union[str, Sequence[RunSpec]]
"""A comment paragraph: plain text or a sequence of runs."""

CommentContent = Union[str, Sequence[ParagraphSpec]]
"""Comment content: a plain string (one paragraph) or paragraph sequence."""
```

- [ ] **Step 4:** In `manager.py`:
  - Import the aliases: `from docx_comments.models import CommentContent, CommentInfo, CommentThread, PersonInfo` and re-export `CommentContent` from `__init__.py` (`__all__` += `"CommentContent"`).
  - Add normalization helper (near `_parse_author_spec`):

```python
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
                if isinstance(run_spec, str):
                    run_text, fmt = run_spec, {}
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
```

  - Extend `_append_text_content` with `fmt: Optional[dict] = None` (added after `rsid_rpr`); right after creating `run` and setting rsidRPr:

```python
        if fmt:
            rpr = etree.SubElement(run, _qn(NS_W, "rPr"))
            if fmt.get("bold"):
                etree.SubElement(rpr, _qn(NS_W, "b"))
            if fmt.get("italic"):
                etree.SubElement(rpr, _qn(NS_W, "i"))
            if fmt.get("underline"):
                u = etree.SubElement(rpr, _qn(NS_W, "u"))
                u.set(_qn(NS_W, "val"), "single")
```

  - Rework `_add_comment_xml`: signature becomes `(self, comment_id, para_id, text_id, content, author, initials, timestamp=None, used_hex_ids=None)` where `content: list[list[tuple[str, dict]]]` is pre-normalized. Replace the single-paragraph build (from `para = etree.SubElement(...)` through `self._append_text_content(para, text, rsid_rpr)`) with:

```python
        if used_hex_ids is None:
            used_hex_ids = self._used_long_hex_ids()
        for index, runs in enumerate(content):
            is_last = index == len(content) - 1
            para = etree.SubElement(comment, _qn(NS_W, "p"))
            para.set(_qn(NS_W, "rsidR"), rsid_r)
            para.set(_qn(NS_W, "rsidRDefault"), rsid_default)
            # Word keys satellite metadata to the LAST paragraph.
            para.set(
                _qn(NS_W14, "paraId"),
                para_id if is_last else self._new_long_hex_id(used_hex_ids),
            )
            para.set(
                _qn(NS_W14, "textId"),
                text_id if is_last else self._new_long_hex_id(used_hex_ids),
            )
            pPr = etree.SubElement(para, _qn(NS_W, "pPr"))
            pStyle = etree.SubElement(pPr, _qn(NS_W, "pStyle"))
            pStyle.set(_qn(NS_W, "val"), "CommentText")
            if index == 0:
                run1 = etree.SubElement(para, _qn(NS_W, "r"))
                rPr = etree.SubElement(run1, _qn(NS_W, "rPr"))
                rStyle = etree.SubElement(rPr, _qn(NS_W, "rStyle"))
                rStyle.set(_qn(NS_W, "val"), "CommentReference")
                etree.SubElement(run1, _qn(NS_W, "annotationRef"))
            if not runs:
                self._append_text_content(para, "", rsid_rpr)
            for run_text, fmt in runs:
                self._append_text_content(para, run_text, rsid_rpr, fmt)
```

  - `add_comment`/`reply_to_comment`: change `text: str` to `text: CommentContent`; replace `_validate_xml_text(text, "comment text")` with `content = self._normalize_content(text)`; pass `content=content, used_hex_ids=used_hex_ids` to `_add_comment_xml` (the `used_hex_ids` set is already computed in both callers).
  - `edit_comment`: change `text: str` to `text: CommentContent`; replace validation with `content = self._normalize_content(text)`; replace the single `self._append_text_content(new_para, text, rsid_rpr)` + para build with a loop building one `w:p` per content paragraph mirroring the block above (primary/text_id on the LAST paragraph, annotationRef run on the first, fresh ids for the others drawn from `used_hex_ids`), then append all new paras to `comment_elem`.
  - Docstrings for all three: document the content forms with a short example.

- [ ] **Step 5:** Full suite + mypy + ruff. Commit: `git commit -am "feat: formatted runs and multi-paragraph comment content"`

```json:metadata
{"files": ["src/docx_comments/models.py", "src/docx_comments/manager.py", "src/docx_comments/__init__.py", "tests/test_rich_content.py"], "verifyCommand": "pytest tests/test_rich_content.py -q", "acceptanceCriteria": ["str behavior unchanged", "bold/italic/underline runs emitted", "multi-paragraph keyed to last paraId", "validation before mutation"], "modelTier": "frontier"}
```

---

### Task 11: Character-offset and substring anchoring

**Goal:** Anchor comments to exact character spans (`start_char`/`end_char` on `add_comment`) and to substrings/regex matches (`add_comment_on_text`), splitting runs as needed — the standard OOXML technique. Splitting must never change the document's visible text.

**Files:**
- Modify: `src/docx_comments/anchors.py` (atom iteration, run splitting, char-span resolution; imports `copy`, NS_XML), `src/docx_comments/manager.py` (add_comment params, add_comment_on_text; `import re`)
- Test: `tests/test_char_anchoring.py` (new)

**Acceptance Criteria:**
- [ ] `add_comment(para, ..., start_char=s, end_char=e)` anchors exactly `paragraph_text[s:e]` (verified via `get_anchored_text`), splitting runs when offsets fall mid-run; split halves keep their formatting (`w:rPr` deep-copied) and whitespace (`xml:space="preserve"` where needed)
- [ ] Paragraph visible text is IDENTICAL before and after anchoring (round-trip through save/reload)
- [ ] Offsets at existing run boundaries do not split; runs inside `w:hyperlink`/`w:ins` are split in place (markers land inside the container — valid EG_PContent positions)
- [ ] `add_comment_on_text(para, match, ...)` anchors the nth occurrence of a substring or `re.Pattern`; missing matches raise ValueError naming the count found
- [ ] Mixing run indices and char offsets raises ValueError; out-of-range offsets raise IndexError; validation happens before any mutation (no split on failure)

**Verify:** `pytest tests/test_char_anchoring.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_char_anchoring.py`:

```python
"""Character-offset and substring anchoring."""

import re
import zipfile

import pytest
from docx import Document
from lxml import etree

from docx_comments import CommentManager, PersonInfo

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(ns, name):
    return f"{{{ns}}}{name}"


def _setup(text="The quick brown fox"):
    doc = Document()
    para = doc.add_paragraph(text)
    return doc, para, CommentManager(doc)


class TestCharSpan:
    def test_mid_run_anchor(self):
        doc, para, mgr = _setup()
        cid = mgr.add_comment(
            para, "c", PersonInfo(author="A"), start_char=4, end_char=9
        )
        assert mgr.get_anchored_text(cid) == "quick"
        assert para.text == "The quick brown fox"

    def test_text_survives_save_reload(self, tmp_path):
        doc, para, mgr = _setup()
        mgr.add_comment(para, "c", PersonInfo(author="A"), start_char=4, end_char=9)
        path = tmp_path / "s.docx"
        doc.save(str(path))
        doc2 = Document(str(path))
        assert doc2.paragraphs[0].text == "The quick brown fox"

    def test_boundary_offsets_do_not_split(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("one ")
        para.add_run("two")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(
            para, "c", PersonInfo(author="A"), start_char=4, end_char=7
        )
        assert mgr.get_anchored_text(cid) == "two"
        assert len(para._element.findall(qn(NS_W, "r"))) == 3  # 2 text + 1 ref

    def test_split_preserves_formatting(self):
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run("boldtext")
        run.bold = True
        mgr = CommentManager(doc)
        cid = mgr.add_comment(
            para, "c", PersonInfo(author="A"), start_char=4, end_char=8
        )
        assert mgr.get_anchored_text(cid) == "text"
        for r in para._element.findall(qn(NS_W, "r")):
            t = r.find(qn(NS_W, "t"))
            if t is not None and t.text:
                assert r.find(f"{qn(NS_W, 'rPr')}/{qn(NS_W, 'b')}") is not None

    def test_whitespace_edges_preserved(self):
        doc, para, mgr = _setup("keep  spaces  here")
        mgr.add_comment(para, "c", PersonInfo(author="A"), start_char=6, end_char=12)
        assert para.text == "keep  spaces  here"

    def test_split_inside_hyperlink(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("See ")
        hyperlink = etree.SubElement(para._element, qn(NS_W, "hyperlink"))
        run = etree.SubElement(hyperlink, qn(NS_W, "r"))
        t = etree.SubElement(run, qn(NS_W, "t"))
        t.text = "the linked words"
        mgr = CommentManager(doc)
        cid = mgr.add_comment(
            para, "c", PersonInfo(author="A"), start_char=8, end_char=14
        )
        assert mgr.get_anchored_text(cid) == "linked"

    def test_br_counts_one_char(self):
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run("ab")
        etree.SubElement(run._element, qn(NS_W, "br"))
        run2 = para.add_run("cd")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(
            para, "c", PersonInfo(author="A"), start_char=3, end_char=5
        )
        assert mgr.get_anchored_text(cid) == "cd"

    def test_validation_errors(self):
        doc, para, mgr = _setup()
        with pytest.raises(ValueError, match="not both"):
            mgr.add_comment(
                para, "c", PersonInfo(author="A"),
                start_run=0, end_run=0, start_char=0, end_char=3,
            )
        with pytest.raises(ValueError, match="together"):
            mgr.add_comment(para, "c", PersonInfo(author="A"), start_char=1)
        with pytest.raises(IndexError):
            mgr.add_comment(
                para, "c", PersonInfo(author="A"), start_char=0, end_char=999
            )
        with pytest.raises(ValueError):
            mgr.add_comment(
                para, "c", PersonInfo(author="A"), start_char=5, end_char=5
            )
        # No mutation on failure: no runs were split.
        assert len(para._element.findall(qn(NS_W, "r"))) == 1


class TestAddCommentOnText:
    def test_substring(self):
        doc, para, mgr = _setup()
        cid = mgr.add_comment_on_text(para, "brown", "c", PersonInfo(author="A"))
        assert mgr.get_anchored_text(cid) == "brown"

    def test_occurrence(self):
        doc, para, mgr = _setup("aba aba aba")
        cid = mgr.add_comment_on_text(
            para, "aba", "c", PersonInfo(author="A"), occurrence=2
        )
        info_para = mgr.get_comment_paragraph(cid)
        assert mgr.get_anchored_text(cid) == "aba"
        start = info_para._element.find(f".//{qn(NS_W, 'commentRangeStart')}")
        # second occurrence starts at char 4: preceding text is "aba "
        preceding = "".join(
            t.text or ""
            for t in start.itersiblings(qn(NS_W, "r"), preceding=True)
            for t in t.findall(qn(NS_W, "t"))
        )
        assert preceding.endswith("aba ")

    def test_regex(self):
        doc, para, mgr = _setup()
        cid = mgr.add_comment_on_text(
            para, re.compile(r"qu\w+"), "c", PersonInfo(author="A")
        )
        assert mgr.get_anchored_text(cid) == "quick"

    def test_missing_match_raises_with_count(self):
        doc, para, mgr = _setup()
        with pytest.raises(ValueError, match="0 time"):
            mgr.add_comment_on_text(para, "zebra", "c", PersonInfo(author="A"))
        with pytest.raises(ValueError, match="1 time"):
            mgr.add_comment_on_text(
                para, "quick", "c", PersonInfo(author="A"), occurrence=2
            )
```

  (In `test_occurrence` simplify the preceding-text assertion if awkward: assert `mgr.get_anchored_text(cid) == "aba"` plus that exactly one commentRangeStart exists, and check its absolute offset by re-deriving `CommentAnchor(doc).paragraph_text(para._element)` — implementer's choice, but the second-occurrence position MUST be asserted somehow, e.g. by comparing the text before the marker.)

- [ ] **Step 2:** Run → FAIL.

- [ ] **Step 3:** In `anchors.py` add imports: `import copy` (top), `NS_XML = "http://www.w3.org/XML/1998/namespace"` (near NS_W). Add to `CommentAnchor`:

```python
    def _iter_paragraph_atoms(self, para_elem):
        """Yield (child, run, length, is_text) for text atoms in order.

        Mirrors the comment-text extraction rules: w:t contributes its
        characters, w:br/w:cr/w:tab one character each; runs inside
        mc:Fallback or nested paragraphs (text boxes) are skipped.
        """
        p_tag = _qn(NS_W, "p")
        fallback = _qn(NS_MC, "Fallback")
        for run in para_elem.iter(_qn(NS_W, "r")):
            skip = False
            parent = run.getparent()
            while parent is not None and parent is not para_elem:
                if parent.tag == fallback or parent.tag == p_tag:
                    skip = True
                    break
                parent = parent.getparent()
            if skip:
                continue
            for child in run:
                local = etree.QName(child).localname
                if local == "t":
                    yield child, run, len(child.text or ""), True
                elif local in ("br", "cr", "tab"):
                    yield child, run, 1, False

    def paragraph_text(self, para_elem: etree._Element) -> str:
        """The paragraph's visible text under the same rules as anchoring."""
        pieces: list[str] = []
        for child, _, _, is_text in self._iter_paragraph_atoms(para_elem):
            if is_text:
                pieces.append(child.text or "")
            else:
                local = etree.QName(child).localname
                pieces.append("\t" if local == "tab" else "\n")
        return "".join(pieces)

    def _check_char_bounds(
        self, para_elem: etree._Element, start_char: int, end_char: int
    ) -> None:
        if not isinstance(start_char, int) or not isinstance(end_char, int):
            raise TypeError("start_char and end_char must be integers")
        if start_char < 0 or end_char < 0:
            raise IndexError("character offsets must be non-negative")
        if end_char <= start_char:
            raise ValueError(
                f"end_char {end_char} must be greater than start_char {start_char}"
            )
        total = sum(
            length for _, _, length, _ in self._iter_paragraph_atoms(para_elem)
        )
        if end_char > total:
            raise IndexError(
                f"end_char {end_char} out of range for paragraph with "
                f"{total} character(s)"
            )

    def validate_char_span(
        self, paragraph: Paragraph, start_char: int, end_char: int
    ) -> None:
        """Validate a character span without mutating anything."""
        self._validate_owned(paragraph)
        self._check_char_bounds(paragraph._element, start_char, end_char)

    def _split_run_at_child(self, run, child):
        """Split `run` immediately before `child`. Returns (left, right)."""
        left = etree.Element(_qn(NS_W, "r"))
        for key, value in run.attrib.items():
            left.set(key, value)
        rpr = run.find(_qn(NS_W, "rPr"))
        if rpr is not None:
            left.append(copy.deepcopy(rpr))
        for sibling in list(run):
            if sibling is child:
                break
            if sibling.tag == _qn(NS_W, "rPr"):
                continue
            left.append(sibling)  # moves the element out of `run`
        run.addprevious(left)
        return left, run

    def _split_run_at_text(self, run, t_child, offset):
        """Split a run inside its w:t at `offset` characters. Returns (left, right)."""
        text = t_child.text or ""
        left_run, right_run = self._split_run_at_child(run, t_child)
        left_t = etree.SubElement(left_run, _qn(NS_W, "t"))
        left_text = text[:offset]
        left_t.text = left_text
        if left_text.strip() != left_text:
            left_t.set(_qn(NS_XML, "space"), "preserve")
        right_text = text[offset:]
        t_child.text = right_text
        if right_text.strip() != right_text:
            t_child.set(_qn(NS_XML, "space"), "preserve")
        else:
            t_child.attrib.pop(_qn(NS_XML, "space"), None)
        return left_run, right_run

    def _ensure_run_boundary(self, para_elem, pos):
        """Make a run boundary exist at char offset `pos`.

        Returns (run_ending_at_pos, run_starting_at_pos); either side is
        None at the paragraph edges.
        """
        cum = 0
        prev_run = None
        for child, run, length, is_text in self._iter_paragraph_atoms(para_elem):
            if cum == pos:
                if prev_run is run:
                    return self._split_run_at_child(run, child)
                return prev_run, run
            if cum < pos < cum + length:
                return self._split_run_at_text(run, child, pos - cum)
            cum += length
            prev_run = run
        return prev_run, None

    def add_anchors_at_char_span(
        self,
        paragraph: Paragraph,
        start_char: int,
        end_char: int,
        comment_id: str,
    ) -> None:
        """Anchor a comment to an exact character span, splitting runs.

        Splitting preserves formatting (rPr is deep-copied) and whitespace
        (xml:space="preserve" on chunks with edge whitespace); the
        paragraph's visible text is unchanged.
        """
        self._validate_owned(paragraph)
        para_elem = paragraph._element
        self._check_char_bounds(para_elem, start_char, end_char)
        _, first_run = self._ensure_run_boundary(para_elem, start_char)
        last_run, _ = self._ensure_run_boundary(para_elem, end_char)
        self.add_anchors_at_span(paragraph, (first_run, last_run), comment_id)
```

- [ ] **Step 4:** In `manager.py`:
  - `import re` at the top.
  - `add_comment` gains `start_char: Optional[int] = None, end_char: Optional[int] = None` (after `end_run`). In the validation block replace `anchor.validate_anchor_target(...)` with:

```python
        use_char_span = start_char is not None or end_char is not None
        if use_char_span:
            if start_char is None or end_char is None:
                raise ValueError("start_char and end_char must be provided together")
            if start_run != 0 or end_run is not None:
                raise ValueError(
                    "pass either start_run/end_run or start_char/end_char, not both"
                )
            anchor.validate_char_span(paragraph, start_char, end_char)
        else:
            anchor.validate_anchor_target(paragraph, start_run, end_run)
```

    and in step "2. Add anchors":

```python
        if use_char_span:
            anchor.add_anchors_at_char_span(
                paragraph, start_char, end_char, comment_id
            )
        else:
            anchor.add_anchors(
                paragraph=paragraph,
                comment_id=comment_id,
                start_run=start_run,
                end_run=end_run,
            )
```

    Document both params: character offsets over the paragraph's visible text (w:t characters; br/cr/tab count as one), end-exclusive.
  - Add `add_comment_on_text` right after `add_comment`:

```python
    def add_comment_on_text(
        self,
        paragraph: Paragraph,
        match: Union[str, "re.Pattern[str]"],
        text: CommentContent,
        author: PersonInfo,
        initials: Optional[str] = None,
        occurrence: int = 1,
        person: Optional[PersonSpec] = None,
        timestamp: Optional[datetime] = None,
    ) -> str:
        """Anchor a comment to the nth occurrence of a substring or pattern.

        Args:
            paragraph: The paragraph to search (visible text; br/cr/tab
                count as one character, matching get_anchored_text).
            match: Substring or compiled regular expression to anchor.
            occurrence: 1-based occurrence to anchor (default: first).
            (other args as add_comment)

        Returns:
            The comment ID of the new comment.

        Raises:
            ValueError: If the match is empty, occurs fewer than
                `occurrence` times, or matches zero characters.
        """
        if occurrence < 1:
            raise ValueError("occurrence must be >= 1")
        anchor = CommentAnchor(self._document)
        para_text = anchor.paragraph_text(paragraph._element)
        spans: list[tuple[int, int]] = []
        if isinstance(match, re.Pattern):
            spans = [m.span() for m in match.finditer(para_text)]
            label = match.pattern
        else:
            if not match:
                raise ValueError("match must be non-empty")
            idx = para_text.find(match)
            while idx != -1:
                spans.append((idx, idx + len(match)))
                idx = para_text.find(match, idx + 1)
            label = match
        if len(spans) < occurrence:
            raise ValueError(
                f"{label!r} occurs {len(spans)} time(s) in the paragraph; "
                f"occurrence {occurrence} requested"
            )
        start_char, end_char = spans[occurrence - 1]
        if start_char == end_char:
            raise ValueError(f"{label!r} matches zero characters")
        return self.add_comment(
            paragraph,
            text,
            author,
            initials=initials,
            start_char=start_char,
            end_char=end_char,
            person=person,
            timestamp=timestamp,
        )
```

- [ ] **Step 5:** Full suite + mypy + ruff (note: `re.Pattern[str]` needs quoting or `from __future__ import annotations` — manager already has it). Commit: `git commit -am "feat: character-offset and substring anchoring with run splitting"`

```json:metadata
{"files": ["src/docx_comments/anchors.py", "src/docx_comments/manager.py", "tests/test_char_anchoring.py"], "verifyCommand": "pytest tests/test_char_anchoring.py -q", "acceptanceCriteria": ["exact char spans anchored via run splitting", "visible text unchanged and round-trips", "substring/regex/occurrence targeting with clear errors", "no mutation on validation failure"], "modelTier": "frontier"}
```

---

### Task 12: Accept plain-str authors

**Goal:** `author` on `add_comment`/`reply_to_comment`/`add_comment_on_text` accepts `str` (the 95% case) alongside `PersonInfo`.

**Files:**
- Modify: `src/docx_comments/manager.py` (`_parse_author_spec`, signatures/docstrings)
- Test: `tests/test_api_polish.py`

**Acceptance Criteria:**
- [ ] `add_comment(para, "c", "Reviewer Name", initials="RN")` works; author round-trips
- [ ] `PersonInfo` authors (with presence auto-linking) behave exactly as before
- [ ] Non-str/PersonInfo authors raise TypeError

**Verify:** `pytest tests/test_api_polish.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_api_polish.py`:

```python
class TestStrAuthor:
    def test_plain_str_author(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", "Reviewer Name", initials="RN")
        info = mgr.get_comment(cid)
        assert info.author == "Reviewer Name" and info.initials == "RN"

    def test_str_author_in_reply(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", "A")
        rid = mgr.reply_to_comment(cid, "r", "B")
        assert mgr.get_comment(rid).author == "B"

    def test_invalid_author_type(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        with pytest.raises(TypeError):
            mgr.add_comment(para, "c", 42)
```

- [ ] **Step 2:** Run → FAIL (TypeError for str).

- [ ] **Step 3:** In `_parse_author_spec`:

```python
    def _parse_author_spec(
        self, author: Union[str, PersonInfo]
    ) -> tuple[str, Optional[dict[str, str]]]:
        if isinstance(author, str):
            author = PersonInfo(author=author)
        if not isinstance(author, PersonInfo):
            raise TypeError("author must be a str or PersonInfo")
        # ... rest unchanged
```

  Widen the `author` annotations on `add_comment`, `reply_to_comment`, `add_comment_on_text` to `Union[str, PersonInfo]` and update their docstrings ("author: Author name as a plain string, or a PersonInfo (required for presence linkage).").

- [ ] **Step 4:** Full suite + commit: `git commit -am "feat: accept plain-str authors"`

```json:metadata
{"files": ["src/docx_comments/manager.py", "tests/test_api_polish.py"], "verifyCommand": "pytest tests/test_api_polish.py -q", "acceptanceCriteria": ["str author works in add/reply", "PersonInfo path unchanged", "TypeError for other types"], "modelTier": "mechanical"}
```

---

### Task 13: Skip redundant metadata migration in resolve/delete

**Goal:** `set_comment_resolved`/`delete_comment`/`delete_thread` skip the full-document `migrate_comment_metadata()` scan when a cheap satellite-parts check proves it would be a no-op — removing the main per-operation O(document) cost in batch workflows without any caching of document state.

**Files:**
- Modify: `src/docx_comments/manager.py` (new `_metadata_complete`, three call sites)
- Test: `tests/test_editing.py`

**Acceptance Criteria:**
- [ ] On a document whose comments were all created by this library, `resolve_comment`/`delete_comment`/`delete_thread` never call `migrate_comment_metadata`
- [ ] With any incompleteness (missing paraId/textId, missing threading/durable/extensible entry, orphan entry, dangling parent) the migration still runs and the operation still works
- [ ] `_metadata_complete` reads ONLY comments.xml + the three satellite parts (no body/header/footnote scan)

**Verify:** `pytest tests/test_editing.py -v` → pass

**Steps:**

- [ ] **Step 1: Failing tests** in `tests/test_editing.py`:

```python
class TestMigrationSkip:
    def _spy(self, monkeypatch):
        calls = []
        original = CommentManager.migrate_comment_metadata

        def wrapper(self):
            calls.append(1)
            original(self)

        monkeypatch.setattr(CommentManager, "migrate_comment_metadata", wrapper)
        return calls

    def test_resolve_skips_migration_when_complete(self, monkeypatch):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        calls = self._spy(monkeypatch)
        mgr.resolve_comment(cid)
        assert calls == []
        assert mgr.get_comment(cid).is_resolved

    def test_delete_skips_migration_when_complete(self, monkeypatch):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        calls = self._spy(monkeypatch)
        mgr.delete_comment(cid)
        assert calls == []
        assert list(mgr.list_comments()) == []

    def test_incomplete_metadata_still_migrates(self, monkeypatch):
        NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        # Strip the paraId to simulate a foreign/legacy comment.
        for p in mgr._comments_xml.iter(qn(NS_W, "p")):
            p.attrib.pop(qn(NS_W14, "paraId"), None)
        calls = self._spy(monkeypatch)
        mgr.resolve_comment(cid)
        assert calls == [1]
        assert mgr.get_comment(cid).is_resolved
```

- [ ] **Step 2:** Run → FAIL (calls == [1] in the skip tests).

- [ ] **Step 3:** Add to `CommentManager` (near `migrate_comment_metadata`):

```python
    def _metadata_complete(self) -> bool:
        """Cheap check whether migrate_comment_metadata would be a no-op.

        Reads only comments.xml and the satellite parts (no body scan), so
        callers can skip the full migration on well-formed documents. Any
        incompleteness — missing paraId/textId, missing threading/durable/
        extensible entries, orphan entries, dangling parents — returns
        False and the caller runs the real migration.
        """
        threading = CommentsExtendedPart(self._document).get_threading_info()
        durable_ids = CommentsIdsPart(self._document).get_durable_ids()
        extensible = CommentsExtensiblePart(self._document).get_extensible_info()

        valid_para_ids: set[str] = set()
        for comment_elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            para_ids: list[str] = []
            for para in comment_elem.findall(_qn(NS_W, "p")):
                pid = para.get(_qn(NS_W14, "paraId"))
                if not pid or not para.get(_qn(NS_W14, "textId")):
                    return False
                para_ids.append(pid)
                valid_para_ids.add(pid)
            if not para_ids:
                continue
            primary = self._primary_para_id(para_ids, threading, durable_ids)
            if primary not in threading:
                return False
            durable = durable_ids.get(primary)
            if not durable:
                return False
            if not (extensible.get(durable) or {}).get("date_utc"):
                return False

        for pid, info in threading.items():
            if pid not in valid_para_ids:
                return False  # orphan commentEx entry
            parent = info.get("parent_para_id")
            if parent and parent not in valid_para_ids:
                return False  # dangling reply link
        for pid in durable_ids:
            if pid not in valid_para_ids:
                return False  # orphan commentId entry
        return True
```

- [ ] **Step 4:** In `set_comment_resolved`, `delete_comment`, and `delete_thread` replace `self.migrate_comment_metadata()` with:

```python
        if not self._metadata_complete():
            self.migrate_comment_metadata()
```

- [ ] **Step 5:** Full suite + commit: `git commit -am "perf: skip full metadata migration when satellite parts are complete"`

```json:metadata
{"files": ["src/docx_comments/manager.py", "tests/test_editing.py"], "verifyCommand": "pytest tests/test_editing.py -q", "acceptanceCriteria": ["migration skipped on complete metadata", "still runs on any incompleteness", "no document-body scan in the check"], "modelTier": "standard"}
```

---

### Task 14: Coverage batch A — person specs, auto_migrate, lifecycle errors, empty paragraphs

**Goal:** Test the untested README-documented person forms, the auto_migrate constructor flag, the newer lifecycle ops' error paths, empty-paragraph anchoring, and move_thread with explicit indices. Fix the add_comment docstring's person=None claim.

**Files:**
- Modify: `src/docx_comments/manager.py` (docstring only), `tests/test_people.py`, `tests/test_migration.py`, `tests/test_robustness.py`, `tests/test_editing.py`

**Acceptance Criteria:**
- [ ] person=True / dict (snake_case AND camelCase keys, and the `{"presence": {...}}` form) / person=42 TypeError / reply person=True all tested with saved-XML assertions
- [ ] The auto-link behavior (presence-bearing PersonInfo author with person=None creates a people.xml entry; person=False suppresses) is pinned by a test AND the add_comment docstring says so
- [ ] auto_migrate=True backfills on init; default False does not
- [ ] delete_thread/move_comment/move_thread with unknown ids raise CommentNotFoundError and mutate nothing
- [ ] Empty-paragraph anchoring: element order (commentRangeStart, commentRangeEnd, styled ref run), after w:pPr when present, round-trips; move_comment onto an empty paragraph works
- [ ] move_thread with explicit start_run/end_run covered

**Verify:** `pytest tests/test_people.py tests/test_migration.py tests/test_robustness.py tests/test_editing.py -q` → pass; coverage of `manager.py` lines 837, 853-870, 937, 939, 1019, 1021, 147 and `anchors.py` 397-417 no longer missing

**Steps:**

- [ ] **Step 1:** In `tests/test_people.py` add:

```python
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def _people_entries(path):
    with zipfile.ZipFile(path) as zf:
        if "word/people.xml" not in zf.namelist():
            return []
        root = etree.fromstring(zf.read("word/people.xml"))
    return [e for e in root if etree.QName(e).localname == "person"]


class TestPersonSpecForms:
    def _mgr(self):
        doc = Document()
        para = doc.add_paragraph("text")
        return doc, para, CommentManager(doc)

    def test_person_true_creates_entry(self, tmp_path):
        doc, para, mgr = self._mgr()
        mgr.add_comment(para, "c", PersonInfo(author="A"), person=True)
        path = tmp_path / "p.docx"
        doc.save(str(path))
        entries = _people_entries(path)
        assert len(entries) == 1
        assert entries[0].get(f"{{{NS_W15}}}author") == "A"

    @pytest.mark.parametrize(
        "spec",
        [
            {"provider_id": "AD", "user_id": "S::u"},
            {"providerId": "AD", "userId": "S::u"},
            {"presence": {"provider_id": "AD", "user_id": "S::u"}},
        ],
    )
    def test_person_dict_forms(self, tmp_path, spec):
        doc, para, mgr = self._mgr()
        mgr.add_comment(para, "c", PersonInfo(author="A"), person=spec)
        path = tmp_path / "p.docx"
        doc.save(str(path))
        entries = _people_entries(path)
        assert len(entries) == 1
        presence = entries[0].find(f"{{{NS_W15}}}presenceInfo")
        assert presence is not None
        assert presence.get(f"{{{NS_W15}}}providerId") == "AD"
        assert presence.get(f"{{{NS_W15}}}userId") == "S::u"

    def test_person_invalid_type(self):
        _, para, mgr = self._mgr()
        with pytest.raises(TypeError, match="person must be"):
            mgr.add_comment(para, "c", PersonInfo(author="A"), person=42)

    def test_reply_person_true(self, tmp_path):
        doc, para, mgr = self._mgr()
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        mgr.reply_to_comment(cid, "r", PersonInfo(author="B"), person=True)
        path = tmp_path / "p.docx"
        doc.save(str(path))
        authors = {
            e.get(f"{{{NS_W15}}}author") for e in _people_entries(path)
        }
        assert authors == {"B"}

    def test_presence_author_auto_links(self, tmp_path):
        doc, para, mgr = self._mgr()
        mgr.add_comment(
            para, "c",
            PersonInfo(author="A", provider_id="AD", user_id="S::u"),
        )
        path = tmp_path / "p.docx"
        doc.save(str(path))
        assert len(_people_entries(path)) == 1

    def test_person_false_suppresses_auto_link(self, tmp_path):
        doc, para, mgr = self._mgr()
        mgr.add_comment(
            para, "c",
            PersonInfo(author="A", provider_id="AD", user_id="S::u"),
            person=False,
        )
        path = tmp_path / "p.docx"
        doc.save(str(path))
        assert _people_entries(path) == []
```

- [ ] **Step 2:** Fix `add_comment`'s docstring person paragraph: replace "None/False leave people.xml untouched." with "False leaves people.xml untouched. None (the default) also leaves it untouched UNLESS the author PersonInfo carries presence metadata, in which case the entry is auto-linked; pass person=False to suppress that."

- [ ] **Step 3:** In `tests/test_migration.py` add:

```python
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


def _strip_comment_metadata(path):
    """Remove w14 attrs + satellite entries from a saved doc (helper for
    auto_migrate tests). Reuse the module's existing stripping helper if one
    exists instead of this."""
    doc = Document(str(path))
    mgr = CommentManager(doc)
    for p in mgr._comments_xml.iter(f"{{{NS_W}}}p"):
        p.attrib.pop(f"{{{NS_W14}}}paraId", None)
        p.attrib.pop(f"{{{NS_W14}}}textId", None)
    mgr._save_comments()
    doc.save(str(path))


def test_auto_migrate_backfills_on_init(tmp_path):
    doc = Document()
    para = doc.add_paragraph("text")
    mgr = CommentManager(doc)
    mgr.add_comment(para, "c", PersonInfo(author="A"))
    path = tmp_path / "a.docx"
    doc.save(str(path))
    _strip_comment_metadata(path)

    doc2 = Document(str(path))
    mgr2 = CommentManager(doc2, auto_migrate=True)
    comment = next(iter(mgr2.list_comments()))
    assert comment.para_id and comment.durable_id


def test_no_auto_migrate_by_default(tmp_path):
    doc = Document()
    para = doc.add_paragraph("text")
    mgr = CommentManager(doc)
    mgr.add_comment(para, "c", PersonInfo(author="A"))
    path = tmp_path / "a.docx"
    doc.save(str(path))
    _strip_comment_metadata(path)

    doc2 = Document(str(path))
    mgr2 = CommentManager(doc2)
    comment = next(iter(mgr2.list_comments()))
    assert comment.para_id == ""
```

  (Check `tests/test_migration.py` first — if it already has a metadata-stripping helper, reuse it.)

- [ ] **Step 4:** In `tests/test_robustness.py` extend `TestLifecycleSafety` (or add a new class):

```python
    def test_delete_thread_unknown_id_no_mutation(self):
        doc = Document()
        para = doc.add_paragraph("text")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        before = etree.tostring(doc.element.body)
        with pytest.raises(CommentNotFoundError):
            mgr.delete_thread("999999")
        assert etree.tostring(doc.element.body) == before
        assert len(list(mgr.list_comments())) == 1

    def test_move_ops_unknown_id_no_mutation(self):
        doc = Document()
        para = doc.add_paragraph("text")
        target = doc.add_paragraph("target")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        before = etree.tostring(doc.element.body)
        with pytest.raises(CommentNotFoundError):
            mgr.move_comment("999999", target)
        with pytest.raises(CommentNotFoundError):
            mgr.move_thread("999999", target)
        assert etree.tostring(doc.element.body) == before

    def test_move_thread_with_explicit_indices(self):
        doc = Document()
        para = doc.add_paragraph("source")
        target = doc.add_paragraph()
        target.add_run("one ")
        target.add_run("two ")
        target.add_run("three")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "root", PersonInfo(author="A"))
        mgr.reply_to_comment(cid, "reply", PersonInfo(author="B"))
        mgr.move_thread(cid, target, start_run=1, end_run=1)
        assert mgr.get_anchored_text(cid) == "two "
        thread = mgr.get_thread(cid)
        assert thread.reply_count == 1


class TestEmptyParagraphAnchoring:
    def test_default_anchor_on_runless_paragraph(self, tmp_path):
        doc = Document()
        para = doc.add_paragraph("")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        locals_ = [etree.QName(c).localname for c in para._element]
        # pPr may precede (python-docx may not add one for a bare paragraph)
        anchored = [x for x in locals_ if x != "pPr"]
        assert anchored == ["commentRangeStart", "commentRangeEnd", "r"]
        path = tmp_path / "e.docx"
        doc.save(str(path))
        mgr2 = CommentManager(Document(str(path)))
        assert mgr2.get_comment(cid).text == "c"
        assert mgr2.get_comment_paragraph(cid) is not None

    def test_anchor_after_ppr_on_styled_empty_paragraph(self):
        doc = Document()
        para = doc.add_paragraph("", style="Heading 1")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="A"))
        locals_ = [etree.QName(c).localname for c in para._element]
        assert locals_[0] == "pPr"
        assert locals_[1] == "commentRangeStart"

    def test_move_comment_onto_empty_paragraph(self):
        doc = Document()
        para = doc.add_paragraph("source")
        empty = doc.add_paragraph("")
        mgr = CommentManager(doc)
        cid = mgr.add_comment(para, "c", PersonInfo(author="A"))
        mgr.move_comment(cid, empty)
        assert mgr.get_comment_paragraph(cid)._element is empty._element
```

  (Import `CommentNotFoundError` from `docx_comments` at the top of the file.)

- [ ] **Step 5:** Run everything; fix any fixture/import mismatches against the actual helpers in those files. Coverage spot check: `pytest --cov=docx_comments --cov-report=term-missing -q | grep -E "manager|anchors"` — the listed lines must no longer appear.

- [ ] **Step 6:** Commit: `git commit -am "test: person spec forms, auto_migrate, lifecycle error paths, empty-paragraph anchors"`

```json:metadata
{"files": ["src/docx_comments/manager.py", "tests/test_people.py", "tests/test_migration.py", "tests/test_robustness.py", "tests/test_editing.py"], "verifyCommand": "pytest tests/test_people.py tests/test_migration.py tests/test_robustness.py tests/test_editing.py -q", "acceptanceCriteria": ["person spec forms tested with saved-XML assertions", "auto-link behavior pinned and documented", "lifecycle not-found paths tested with no-mutation assertions", "empty-paragraph XML shape pinned"], "modelTier": "mechanical"}
```

---

### Task 15: Coverage batch B — hermetic conftest, system_author, author getters, extensible invariants, cycle guard

**Goal:** Make the suite hermetic against machine state, and cover system_author.py readers, get_authors/get_document_author, the commentsExtensible cross-part invariant, and the parent-chain cycle guard.

**Files:**
- Create: `tests/conftest.py`, `tests/test_system_author.py`
- Modify: `tests/test_xml.py`, `tests/test_manager_basic.py`, `tests/test_robustness.py`

**Acceptance Criteria:**
- [ ] Autouse fixture removes DOCX_COMMENTS_AUTHOR_DOCX and stubs `_system_office_user_info` to (None, None) for every test
- [ ] `_macos_office_user_info` tested with a crafted plist (valid values, wrong-typed values, missing file)
- [ ] Env-var author source, no-people.xml warning + fallback, and the final "no default author" ValueError tested
- [ ] `get_authors` initials preference and `get_document_author` (core author, last_modified_by fallback, empty) tested
- [ ] Saved zip contains `word/commentsExtensible.xml`; every durableId in commentsIds has a matching commentExtensible whose dateUtc equals the comment's w:date normalized to UTC
- [ ] paraIdParent cycle: get_comment_threads terminates and returns both comments; resolve_comment doesn't hang

**Verify:** `pytest -q` → pass; `pytest --cov=docx_comments -q` shows `system_author.py` ≥ 85%

**Steps:**

- [ ] **Step 1:** Create `tests/conftest.py`:

```python
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
```

- [ ] **Step 2:** Create `tests/test_system_author.py`:

```python
"""Direct tests for the system/default author resolution helpers."""

import plistlib

import pytest
from docx import Document

from docx_comments import CommentManager, PersonInfo
from docx_comments import system_author
from docx_comments.system_author import (
    _default_person_from_system,
    _macos_office_user_info,
)


def _write_plist(tmp_path, data):
    office_dir = tmp_path / "Library/Group Containers/UBF8T346G9.Office"
    office_dir.mkdir(parents=True)
    with (office_dir / "MeContact.plist").open("wb") as handle:
        plistlib.dump(data, handle)


class TestMacosPlist:
    def test_reads_name_and_initials(self, tmp_path, monkeypatch):
        _write_plist(tmp_path, {"Name": "Jane Doe", "Initials": "JD"})
        monkeypatch.setattr(
            system_author.Path, "home", classmethod(lambda cls: tmp_path)
        )
        assert _macos_office_user_info() == ("Jane Doe", "JD")

    def test_wrong_types_become_none(self, tmp_path, monkeypatch):
        _write_plist(tmp_path, {"Name": 42, "Initials": ["J"]})
        monkeypatch.setattr(
            system_author.Path, "home", classmethod(lambda cls: tmp_path)
        )
        assert _macos_office_user_info() == (None, None)

    def test_missing_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(
            system_author.Path, "home", classmethod(lambda cls: tmp_path)
        )
        assert _macos_office_user_info() == (None, None)


def _author_docx(tmp_path, author="Env Author"):
    doc = Document()
    mgr = CommentManager(doc)
    mgr.ensure_person(author)
    path = tmp_path / "author-source.docx"
    doc.save(str(path))
    return path


class TestEnvVarSource:
    def test_env_var_used(self, tmp_path, monkeypatch):
        path = _author_docx(tmp_path)
        monkeypatch.setenv("DOCX_COMMENTS_AUTHOR_DOCX", str(path))
        person, initials = _default_person_from_system()
        assert person is not None and person.author == "Env Author"

    def test_docx_without_people_warns_and_falls_back(self, tmp_path, monkeypatch):
        doc = Document()
        path = tmp_path / "plain.docx"
        doc.save(str(path))
        with pytest.warns(UserWarning, match="no people.xml"):
            person, _ = _default_person_from_system(docx_path=str(path))
        assert person is None


class TestManagerDefaultAuthor:
    def test_no_default_author_raises(self):
        doc = Document()
        doc.core_properties.author = ""
        doc.core_properties.last_modified_by = ""
        mgr = CommentManager(doc)
        with pytest.raises(ValueError, match="no default author"):
            mgr.get_default_author_person()

    def test_core_properties_fallback(self):
        doc = Document()
        doc.core_properties.author = "Core Author"
        mgr = CommentManager(doc)
        person, initials = mgr.get_default_author_person()
        assert person.author == "Core Author" and initials is None

    def test_last_modified_by_fallback(self):
        doc = Document()
        doc.core_properties.author = ""
        doc.core_properties.last_modified_by = "Modifier"
        mgr = CommentManager(doc)
        person, _ = mgr.get_default_author_person()
        assert person.author == "Modifier"
```

- [ ] **Step 3:** In `tests/test_manager_basic.py` add:

```python
class TestAuthorGetters:
    def test_get_authors_prefers_non_empty_initials(self):
        doc = Document()
        p1 = doc.add_paragraph("one")
        p2 = doc.add_paragraph("two")
        mgr = CommentManager(doc)
        mgr.add_comment(p1, "a", PersonInfo(author="Alice"))
        mgr.add_comment(p2, "b", PersonInfo(author="Alice"), initials="AL")
        assert mgr.get_authors() == {"Alice": "AL"}

    def test_get_document_author_with_initials_lookup(self):
        doc = Document()
        para = doc.add_paragraph("text")
        doc.core_properties.author = "Alice"
        mgr = CommentManager(doc)
        mgr.add_comment(para, "c", PersonInfo(author="Alice"), initials="AL")
        assert mgr.get_document_author() == ("Alice", "AL")

    def test_get_document_author_last_modified_fallback(self):
        doc = Document()
        doc.core_properties.author = ""
        doc.core_properties.last_modified_by = "Bob"
        mgr = CommentManager(doc)
        assert mgr.get_document_author() == ("Bob", None)

    def test_get_document_author_empty(self):
        doc = Document()
        doc.core_properties.author = ""
        doc.core_properties.last_modified_by = ""
        mgr = CommentManager(doc)
        assert mgr.get_document_author() == ("", None)
```

- [ ] **Step 4:** In `tests/test_xml.py`: extend the parts-created test to assert `"word/commentsExtensible.xml" in names`, and add:

```python
def test_extensible_linked_by_durable_id_with_utc_date(tmp_path):
    from datetime import timezone

    from docx_comments.manager import _parse_comment_date

    doc = Document()
    para = doc.add_paragraph("text")
    mgr = CommentManager(doc)
    mgr.add_comment(para, "c", PersonInfo(author="A"))
    path = tmp_path / "x.docx"
    doc.save(str(path))
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns_cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
    ns_cex = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
    with zipfile.ZipFile(path) as zf:
        ids_root = etree.fromstring(zf.read("word/commentsIds.xml"))
        cex_root = etree.fromstring(zf.read("word/commentsExtensible.xml"))
        comments_root = etree.fromstring(zf.read("word/comments.xml"))
    durable_ids = {
        e.get(f"{{{ns_cid}}}durableId")
        for e in ids_root
        if etree.QName(e).localname == "commentId"
    }
    cex_entries = {
        e.get(f"{{{ns_cex}}}durableId"): e.get(f"{{{ns_cex}}}dateUtc")
        for e in cex_root
        if etree.QName(e).localname == "commentExtensible"
    }
    assert durable_ids and durable_ids == set(cex_entries)
    w_date = comments_root.find(f"{{{ns_w}}}comment").get(f"{{{ns_w}}}date")
    expected = _parse_comment_date(w_date).astimezone(timezone.utc)
    assert list(cex_entries.values())[0] == expected.strftime("%Y-%m-%dT%H:%M:%SZ")
```

- [ ] **Step 5:** In `tests/test_robustness.py` add:

```python
class TestParentChainCycle:
    def test_cycle_terminates(self):
        doc = Document()
        p1 = doc.add_paragraph("one")
        p2 = doc.add_paragraph("two")
        mgr = CommentManager(doc)
        c1 = mgr.add_comment(p1, "a", PersonInfo(author="A"))
        c2 = mgr.add_comment(p2, "b", PersonInfo(author="B"))
        pid1 = mgr.get_comment(c1).para_id
        pid2 = mgr.get_comment(c2).para_id
        ext = CommentsExtendedPart(doc)
        ext.set_parent(pid1, pid2)
        ext.set_parent(pid2, pid1)
        threads = mgr.get_comment_threads()
        total = sum(len(t.all_comments) for t in threads)
        assert total == 2
        mgr.resolve_comment(c1)  # must terminate
```

- [ ] **Step 6:** Full suite + coverage check + commit: `git commit -am "test: hermetic conftest, system_author, author getters, extensible invariants, cycle guard"`

```json:metadata
{"files": ["tests/conftest.py", "tests/test_system_author.py", "tests/test_manager_basic.py", "tests/test_xml.py", "tests/test_robustness.py"], "verifyCommand": "pytest -q && pytest --cov=docx_comments --cov-report=term -q | tail -15", "acceptanceCriteria": ["autouse hermetic fixture", "system_author plist/env/warning/error paths covered", "author getters covered", "extensible cross-part invariant pinned", "cycle guard exercised"], "modelTier": "mechanical"}
```

---

### Task 16: CI — OS matrix, format check, drop dead codecov step

**Goal:** CI runs the suite on macOS and Windows (where system_author has real code paths), enforces `ruff format --check`, and drops the silently-failing codecov upload.

**Files:**
- Modify: `.github/workflows/ci.yml`

**Acceptance Criteria:**
- [ ] Dedicated `lint` job: ruff check, ruff format --check, mypy (once, ubuntu, 3.12)
- [ ] `test` job matrix: full 3.9–3.14 sweep on ubuntu; 3.11+3.14 on macos-latest; 3.9+3.14 on windows-latest (macOS arm64 runners lack reliable <3.11 builds)
- [ ] Cross-platform venv activation (bash shell, `bin/activate || Scripts/activate`)
- [ ] codecov-action step removed; `--cov-fail-under=80` kept
- [ ] `ci-passed` needs lint + test + test-dependency-floor

**Verify:** `python3 -c "import yaml,sys; yaml.safe_load(open('.github/workflows/ci.yml'))"` (or `uv run python -c ...`) parses; visual review of the matrix

**Steps:**

- [ ] **Step 1:** Rewrite `.github/workflows/ci.yml`:

```yaml
name: CI

on:
  push:
    branches: [main, master]
  pull_request:
    branches: [main, master]

jobs:
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Install uv
        uses: astral-sh/setup-uv@v4

      - name: Install dependencies
        run: |
          uv venv
          uv pip install -e ".[dev]"

      - name: Ruff lint
        run: |
          source .venv/bin/activate
          ruff check src/ tests/

      - name: Ruff format check
        run: |
          source .venv/bin/activate
          ruff format --check src/ tests/

      - name: Type check
        run: |
          source .venv/bin/activate
          mypy src/docx_comments

  test:
    runs-on: ${{ matrix.os }}
    strategy:
      fail-fast: false
      matrix:
        os: [ubuntu-latest]
        python-version: ["3.9", "3.10", "3.11", "3.12", "3.13", "3.14"]
        include:
          # system_author.py has real macOS/Windows code paths; run the
          # suite where they actually execute (oldest supported + newest).
          - os: macos-latest
            python-version: "3.11"
          - os: macos-latest
            python-version: "3.14"
          - os: windows-latest
            python-version: "3.9"
          - os: windows-latest
            python-version: "3.14"
    defaults:
      run:
        shell: bash

    steps:
      - uses: actions/checkout@v4

      - name: Set up Python ${{ matrix.python-version }}
        uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}

      - name: Install uv
        uses: astral-sh/setup-uv@v4

      - name: Install dependencies
        run: |
          uv venv
          uv pip install -e ".[dev]"

      - name: Run tests
        run: |
          source .venv/bin/activate 2>/dev/null || source .venv/Scripts/activate
          pytest --cov=docx_comments --cov-report=xml --cov-fail-under=80

  # Exercise the declared dependency floor (python-docx>=1.0.0) so a commit
  # cannot silently start requiring a newer python-docx API.
  test-dependency-floor:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python 3.9
        uses: actions/setup-python@v5
        with:
          python-version: "3.9"

      - name: Install uv
        uses: astral-sh/setup-uv@v4

      - name: Install dependencies at the floor
        run: |
          uv venv
          uv pip install -e ".[dev]" "python-docx==1.0.0" "lxml==4.9.4"

      - name: Run tests
        run: |
          source .venv/bin/activate
          pytest

  ci-passed:
    name: CI Passed
    needs: [lint, test, test-dependency-floor]
    if: always()
    runs-on: ubuntu-latest
    steps:
      - name: Check all jobs passed
        run: |
          if [ "${{ needs.lint.result }}" != "success" ]; then
            echo "Lint job failed"
            exit 1
          fi
          if [ "${{ needs.test.result }}" != "success" ]; then
            echo "Test job failed"
            exit 1
          fi
          if [ "${{ needs.test-dependency-floor.result }}" != "success" ]; then
            echo "Dependency-floor job failed"
            exit 1
          fi
          echo "All checks passed!"
```

- [ ] **Step 2:** Run `ruff format src/ tests/` locally and commit any reformatting NOW (the new CI gate must pass on this branch). Then verify `ruff format --check src/ tests/` is clean.

- [ ] **Step 3:** Commit: `git commit -am "ci: macOS/Windows matrix legs, ruff format gate, drop dead codecov upload"`

```json:metadata
{"files": [".github/workflows/ci.yml"], "verifyCommand": "python3 -c \"import yaml; yaml.safe_load(open('.github/workflows/ci.yml'))\" && ruff format --check src/ tests/", "acceptanceCriteria": ["lint job with format check", "mac/windows matrix legs", "codecov step removed"], "modelTier": "mechanical"}
```

---

### Task 17: Documentation — README, CHANGELOG, CLAUDE.md

**Goal:** Docs match the new API surface; CHANGELOG gets a complete Unreleased 0.5.0 section.

**Files:**
- Modify: `README.md`, `CHANGELOG.md`, `CLAUDE.md`

**Acceptance Criteria:**
- [ ] README Usage shows: plain-str author, `add_comment_on_text`, `start_char/end_char`, `edit_comment`, `get_anchored_text`/`get_comment_paragraph`, `get_comment`/`get_thread`, rich content, `timestamp=`, exception types, int-id note
- [ ] README gains an "API summary" table listing EVERY public CommentManager method with a one-line description (add_comment, add_comment_on_text, reply_to_comment, edit_comment, resolve_comment, unresolve_comment, set_comment_resolved, delete_comment, delete_thread, move_comment, move_thread, list_comments, get_comment, get_thread, get_comment_threads, get_anchored_text, get_comment_paragraph, get_authors, get_document_author, get_people, get_person, ensure_person, merge_people_from, get_default_author_person, migrate_comment_metadata)
- [ ] README line "Author must be a PersonInfo object, not a raw string." removed/replaced
- [ ] CHANGELOG has `## [0.5.0] - Unreleased` with Added / Fixed / Changed subsections covering every task in this plan (one bullet each)
- [ ] CLAUDE.md Module Structure lists exceptions.py and the new manager methods; Testing Notes lists the new test files and the hermetic conftest

**Verify:** `grep -c "add_comment_on_text" README.md` ≥ 2; `grep -c "0.5.0" CHANGELOG.md` ≥ 1; manual read-through

**Steps:**

- [ ] **Step 1:** Update README.md. Key content decisions: keep the existing structure; in Usage replace the `PersonInfo` requirement comment with `# author can be a plain string, or PersonInfo for identity linkage`; add a subsection "Anchoring to exact text" showing `add_comment_on_text(para, "the phrase", "comment", "Reviewer")` and `start_char`/`end_char`; a subsection "Editing and reading back" showing `edit_comment`, `get_anchored_text`, `get_comment_paragraph`, `get_thread`; a subsection "Rich content" with the paragraphs/runs example; a subsection "Errors" naming `CommentNotFoundError`/`PersonNotFoundError` and the int-id acceptance; then the API summary table before "OOXML Parts Handled".
- [ ] **Step 2:** CHANGELOG.md: insert at the top (below the header/intro, above `## [0.4.0]`):

```markdown
## [0.5.0] - Unreleased

### Added
- `edit_comment()` — in-place text/author/initials/date editing preserving comment id, durable id, threading, resolution, and anchors
- Character-offset anchoring (`add_comment(..., start_char=, end_char=)`) and substring/regex anchoring (`add_comment_on_text()`), splitting runs without changing document text
- Read-side API: `get_comment()`, `get_thread()`, `get_anchored_text()`, `get_comment_paragraph()`
- Rich comment content: formatted runs (bold/italic/underline) and multiple paragraphs
- Caller-controlled timestamps on `add_comment`/`reply_to_comment`/`edit_comment`
- Typed exceptions: `CommentNotFoundError` (subclasses ValueError+LookupError), `PersonNotFoundError` (subclasses KeyError)
- Plain-`str` authors accepted alongside `PersonInfo`; python-docx native `int` comment ids accepted everywhere
- CI: macOS and Windows matrix legs, `ruff format --check` gate

### Fixed
- Replies to comments with block-level range markers (body/table level) no longer emit a schema-invalid bare `w:r` that triggers Word's repair prompt
- Comments anchored only by a `commentReference` (no range markers — legal per ECMA-376) can now be replied to
- Default whole-paragraph anchors now include hyperlink/tracked-change/field containers instead of silently truncating
- Replies into a resolved thread inherit `done=1` instead of creating a mixed thread state
- `w15:done` now parsed as ST_OnOff (`"true"`/`"on"` recognized, not just `"1"`)
- XML-illegal person identity values are rejected before any part creation; `ensure_person` writes atomically (no half-built entries)
- comments.xml created via python-docx's native template now declares `mc:Ignorable`
- Anchor reference runs carry Word's `rStyle CommentReference`
- `CommentInfo.comment_id` is `Optional[str]`, matching runtime behavior on id-less comments

### Changed
- `resolve`/`delete` operations skip the full metadata migration scan when satellite parts are already complete (large-document batch performance)
```

- [ ] **Step 3:** CLAUDE.md: add `exceptions.py` to Module Structure; extend the manager.py bullet with the new methods; update Testing Notes with `tests/conftest.py` (hermetic author env), `test_api_polish.py`, `test_read_api.py`, `test_timestamps.py`, `test_rich_content.py`, `test_char_anchoring.py`, `test_system_author.py`.
- [ ] **Step 4:** Commit: `git commit -am "docs: README API coverage, CHANGELOG 0.5.0, CLAUDE.md updates"`

```json:metadata
{"files": ["README.md", "CHANGELOG.md", "CLAUDE.md"], "verifyCommand": "grep -c add_comment_on_text README.md && grep -c '0.5.0' CHANGELOG.md", "acceptanceCriteria": ["README documents all new APIs + API summary table", "CHANGELOG 0.5.0 section complete", "CLAUDE.md updated"], "modelTier": "standard"}
```

---

### Task 18: Final verification + adversarial review

**Goal:** Whole-branch verification with fresh eyes: full gates, then an adversarial multi-agent review of the complete diff; fix confirmed findings.

**Files:**
- Possibly any file touched above (fixes)

**Acceptance Criteria:**
- [ ] `pytest -q` → 0 failures; `pytest --cov=docx_comments --cov-fail-under=80` passes (report the actual figure — expect ≥ 90)
- [ ] `mypy src/docx_comments` clean; `ruff check src/ tests/` clean; `ruff format --check src/ tests/` clean
- [ ] Adversarial review of `git diff main...audit-improvements` (correctness / OOXML-validity / API-compat lenses, findings verified before acting) run; every CONFIRMED finding fixed or explicitly recorded as declined with reason
- [ ] Suite re-run green after fixes

**Verify:** `pytest --cov=docx_comments -q && mypy src/docx_comments && ruff check src/ tests/ && ruff format --check src/ tests/` → all green

**Steps:**

- [ ] **Step 1:** Run all four gates; capture output.
- [ ] **Step 2:** Dispatch an adversarial review of the full branch diff (coordinator: use the Workflow tool with 3 finder lenses — correctness regressions, OOXML schema validity of every new XML shape, backward API compatibility — each followed by a refute-first verifier; this mirrors the audit that produced this plan).
- [ ] **Step 3:** Fix confirmed findings (with a regression test each), re-run gates.
- [ ] **Step 4:** Commit: `git commit -am "fix: address adversarial review findings"` (or note zero findings).
- [ ] **Step 5:** Report to the user: summary of all changes, coverage figure, and the open release decision (upstream PR vs fork-canonical PyPI publishing — needs user's account-level choices).

```json:metadata
{"files": [], "verifyCommand": "pytest --cov=docx_comments -q && mypy src/docx_comments && ruff check src/ tests/ && ruff format --check src/ tests/", "acceptanceCriteria": ["all gates green", "adversarial review run and confirmed findings addressed"], "modelTier": "standard"}
```

---

## Self-Review (completed by plan author)

1. **Spec coverage:** All 37 audit findings map to tasks: block-level reply (T3), stranded release (recorded as user decision + T17 CHANGELOG), mixed-content anchors (T4), reply-to-resolved (T5), done parsing (T2), person XML-legality (T6), CommentInfo hints (T1), edit (T9), anchored text/location (T7), char/substring anchoring (T11), rich content (T10), timestamps (T8), exceptions/int-ids/docstrings (T1), str authors (T12), single lookup (T7), migration skip (T13), person-spec/auto_migrate/lifecycle/empty-para/move-indices tests (T14), conftest/system_author/getters/extensible/cycle tests (T15), CI matrix/format/codecov (T16), CONTRIBUTING (T0), README/API reference/CHANGELOG/CLAUDE.md (T17). Declined items are recorded in the header with reasons. `.coverage` item verified moot.
2. **Placeholder scan:** No TBDs; every code step has concrete code. Two intentional "reuse the file's existing helper if present" notes (T14 migration stripping, T4 hyperlink helper) direct the implementer to prefer existing fixtures — the full fallback code is provided in both cases.
3. **Type consistency:** `_coerce_comment_id` (T1) used by T7/T9; `CommentNotFoundError` (T1) used by T3+ tests; `_primary_para_id` (T9) used by T13; `_make_reference_run` (T4) covers T3's sites (T3 lands first with inline construction; T4 refactors); `content` normalization (T10) feeds `_add_comment_xml(content=...)` and `edit_comment`; `get_anchored_text` (T7) used by T11/T14 tests; `paragraph_text`/`validate_char_span`/`add_anchors_at_char_span` names consistent between anchors.py and manager.py.

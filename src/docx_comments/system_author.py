"""Internal helpers for resolving system/default author information."""

from __future__ import annotations

import os
import sys
import warnings
from pathlib import Path
from typing import Optional, Tuple
from zipfile import ZipFile

from lxml import etree

from docx_comments.models import PersonInfo
from docx_comments.xml_parts import parse_xml_bytes


def _system_office_user_info() -> Tuple[Optional[str], Optional[str]]:
    if sys.platform == "darwin":
        return _macos_office_user_info()
    if sys.platform.startswith("win"):
        return _windows_office_user_info()
    return None, None


def _macos_office_user_info() -> Tuple[Optional[str], Optional[str]]:
    path = Path.home() / "Library/Group Containers/UBF8T346G9.Office/MeContact.plist"
    if not path.exists():
        return None, None
    try:
        import plistlib

        with path.open("rb") as handle:
            data = plistlib.load(handle)
    except Exception:
        return None, None

    if not isinstance(data, dict):
        return None, None

    name = data.get("Name")
    initials = data.get("Initials")
    if not isinstance(name, str):
        name = None
    if not isinstance(initials, str):
        initials = None
    return name, initials


def _windows_office_user_info() -> Tuple[Optional[str], Optional[str]]:
    if sys.platform != "win32":
        return None, None
    try:
        import winreg
    except Exception:
        return None, None

    keys = [
        r"Software\Microsoft\Office\Common\UserInfo",
        r"Software\Microsoft\Office\16.0\Common\UserInfo",
    ]
    for key_path in keys:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                name, _ = winreg.QueryValueEx(key, "UserName")
                initials, _ = winreg.QueryValueEx(key, "UserInitials")
                if not isinstance(name, str):
                    name = None
                if not isinstance(initials, str):
                    initials = None
                return name, initials
        except FileNotFoundError:
            continue
        except OSError:
            continue
    return None, None


class _DocxAuthorAmbiguous(RuntimeError):
    pass


def _person_from_docx(
    docx_path: str, include_presence: bool = False
) -> Tuple[Optional[PersonInfo], Optional[str]]:
    if not docx_path:
        return None, None
    if not Path(docx_path).exists():
        return None, None

    try:
        with ZipFile(docx_path) as zf:
            names = set(zf.namelist())
            person = _docx_single_person(zf, names, include_presence)
            return person, None
    except _DocxAuthorAmbiguous:
        raise
    except Exception:
        return None, None


def _docx_single_person(zf: ZipFile, names: set[str], include_presence: bool) -> PersonInfo:
    if "word/people.xml" not in names:
        raise _DocxAuthorAmbiguous("DOCX author source has no people.xml")
    try:
        xml = parse_xml_bytes(zf.read("word/people.xml"))
    except Exception:
        raise _DocxAuthorAmbiguous("DOCX author source has invalid people.xml")

    people = [elem for elem in xml if etree.QName(elem).localname == "person"]
    if len(people) != 1:
        raise _DocxAuthorAmbiguous(f"DOCX author source has {len(people)} people entries")

    elem = people[0]
    author = _attr_by_localname(elem, "author") or ""
    if not author:
        raise _DocxAuthorAmbiguous("DOCX author source person has no author name")

    if not include_presence:
        return PersonInfo(author=author)

    presence_elem = _find_child_by_localname(elem, "presenceInfo")
    provider_id = user_id = None
    if presence_elem is not None:
        provider_id = _attr_by_localname(presence_elem, "providerId")
        user_id = _attr_by_localname(presence_elem, "userId")
    return PersonInfo(author=author, provider_id=provider_id, user_id=user_id)


def _attr_by_localname(elem: etree._Element, localname: str) -> Optional[str]:
    for attr, value in elem.attrib.items():
        try:
            if etree.QName(attr).localname == localname:
                return str(value)
        except (ValueError, TypeError):
            if attr == localname:
                return str(value)
    return None


def _find_child_by_localname(elem: etree._Element, localname: str) -> Optional[etree._Element]:
    for child in elem:
        if etree.QName(child).localname == localname:
            return child
    return None


def _default_person_from_system(
    docx_path: Optional[str] = None,
    include_presence: bool = False,
    strict_docx: bool = False,
) -> Tuple[Optional[PersonInfo], Optional[str]]:
    env_path = os.environ.get("DOCX_COMMENTS_AUTHOR_DOCX")
    source = docx_path or env_path
    if source:
        docx_ambiguous = False
        try:
            person, initials = _person_from_docx(source, include_presence)
        except _DocxAuthorAmbiguous as exc:
            warnings.warn(str(exc), UserWarning)
            person = None
            initials = None
            docx_ambiguous = True

        if person:
            return person, initials
        if strict_docx and source and not docx_ambiguous:
            return None, None

    name, initials = _system_office_user_info()
    if name:
        return PersonInfo(author=name), initials
    return None, None

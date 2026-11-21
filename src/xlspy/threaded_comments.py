"""Excel 365 threaded comments: the conversation model and its package parts.

A threaded comment is a conversation anchored to a cell. Excel stores it in
``xl/threadedComments/threadedCommentN.xml`` (one part per sheet), with the
people in ``xl/persons/person.xml``. A classic note carrying the same text is
written beside it: Excel uses the note's ``tc={id}`` author to link the two, so
older Excel versions can still show the conversation. xlspy writes both.

The public object is ``ThreadedComment``: its ``replies`` are the messages that
follow the first one. Excel flattens replies, so every reply points at the
first message; xlspy accepts the same one-level model.
"""

from __future__ import annotations

from collections.abc import Collection, Iterable, Mapping
from dataclasses import dataclass, field
import datetime as _datetime
import re
from typing import Any
import uuid
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from ._metadata import MAIN, PKG_REL, append_child, cell_ref, element_pattern, namespace_of, xml_text
from ._safe_xml import fromstring as safe_fromstring
from ._text import check_cell_text, xstring_escape, xstring_unescape, xml_space_attribute
from .comments import Comment, _fragment_node, rels_path, resolve_part
from .exceptions import InvalidFileError
from .refs import RangeIndex, parse_cell_range

#: The threaded-comment namespace and content types Excel writes.
THREADED_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
THREADED_COMMENTS_TYPE = "application/vnd.ms-excel.threadedcomments+xml"
PERSONS_TYPE = "application/vnd.ms-excel.person+xml"

#: Relationship types (the sheet points at its conversations; the workbook at its people).
THREADED_REL = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
PERSON_REL = "http://schemas.microsoft.com/office/2017/10/relationships/person"

#: Excel's one shared people part.
PERSONS_PART = "xl/persons/person.xml"


@dataclass(frozen=True, slots=True)
class Person:
    """One author in the workbook's people part. Internal to the thread codecs."""

    display_name: str
    id: str
    user_id: str = ""
    provider_id: str = "None"


@dataclass
class ThreadedComment:
    """One message in an Excel 365 threaded conversation.

    The first message owns ``replies`` and ``resolved``. Replies are one level
    deep; ``reply()`` appends one. ``created`` is optional: when it is ``None``
    the message carries no timestamp, which keeps a save deterministic. ``id``
    is generated on save when it is ``None``.

    Examples:
        Write a conversation and a reply into a new workbook:

        ```python
        from xlspy import ThreadedComment, XlsxWriter

        thread = ThreadedComment("Confirm the amount", "Jan")
        thread.reply("Confirmed", "Ola")
        with XlsxWriter("review.xlsx") as workbook:
            workbook.add_sheet("Data")
            workbook.write_sheet(
                [["Amount"], [120]],
                threaded_comments={"A2": thread},
            )
        ```

    See Also:
        `Comment` for classic notes, and `XlsxUpdater.set_threaded_comments`
        for conversations in an existing workbook.
    """

    text: str
    author: str
    created: _datetime.datetime | None = None
    resolved: bool = False
    id: str | None = field(default=None, compare=False)
    replies: list[ThreadedComment] = field(default_factory=list)
    _person: Person | None = field(default=None, repr=False, compare=False)
    _parent_id: str | None = field(default=None, repr=False, compare=False)

    def __post_init__(self) -> None:
        xml_text(self.text)
        check_cell_text(self.text)
        xml_text(self.author)
        if self.created is not None and not isinstance(self.created, _datetime.datetime):
            raise TypeError("created must be a datetime or None")
        if not isinstance(self.resolved, bool):
            raise TypeError("resolved must be a bool")
        if self.id is not None and not isinstance(self.id, str):
            raise TypeError("id must be a string or None")
        if not isinstance(self.replies, list):
            raise TypeError("replies must be a list")
        for reply in self.replies:
            if not isinstance(reply, ThreadedComment):
                raise TypeError("replies must contain ThreadedComment objects")
            if reply.replies:
                raise ValueError("a threaded comment reply cannot have replies")
            if reply.resolved:
                raise ValueError("only the first threaded comment of a conversation can be resolved")
            if reply.parent_id is not None:
                raise ValueError("a reply's parent is its conversation, not another message")
        if self._person is not None and not isinstance(self._person, Person):
            raise TypeError("_person must be a Person or None")

    @property
    def parent_id(self) -> str | None:
        """The id of the message this one answers; ``None`` on the first message."""
        return self._parent_id

    def reply(
        self,
        text: str,
        author: str,
        created: _datetime.datetime | None = None,
        id: str | None = None,
    ) -> ThreadedComment:
        """Append a reply to this conversation and return it."""
        message = ThreadedComment(text, author, created, False, id)
        self.replies.append(message)
        return message

    @property
    def messages(self) -> list[ThreadedComment]:
        """The first message followed by its replies."""
        return [self, *self.replies]


def normalize_threaded_comments(
    comments: Mapping[str, ThreadedComment | None] | None,
    *,
    merges: Iterable[str] = (),
) -> dict[str, ThreadedComment | None]:
    """Check a cell-to-conversation mapping and normalize its cell references."""
    if comments is None:
        return {}
    if not isinstance(comments, Mapping):
        raise TypeError("threaded_comments must map cell references to ThreadedComment objects")
    result: dict[str, ThreadedComment | None] = {}
    merged = None
    for ref, thread in comments.items():
        ref = cell_ref(ref)
        if ref in result:
            raise ValueError(f"duplicate threaded comment reference: {ref}")
        if not isinstance(thread, ThreadedComment) and thread is not None:
            raise TypeError("threaded comment values must be ThreadedComment objects or None")
        if thread is not None:
            c, r, _, _ = parse_cell_range(ref)
            if merged is None:
                merged = RangeIndex()
                for merge in merges:
                    merged.add(parse_cell_range(merge))
            for c1, r1, _, _, _ in merged.overlapping((c, r, c, r)):
                if (c, r) != (c1, r1):
                    raise ValueError("threaded comments on merged cells must use the top-left cell")
        result[ref] = thread
    return result


def _thread_id(
    *,
    scope: str,
    ref: str,
    author: str,
    text: str,
    created: _datetime.datetime | None,
    parent: str | None,
    index: int,
) -> str:
    """A stable GUID for a message that has no id yet.

    ``scope`` is the sheet's part name: two sheets can hold the same text, but
    their messages must still get different ids.
    """
    stamp = "" if created is None else created.isoformat()
    seed = f"xlspy:threaded-comment:{scope}:{ref}:{parent or ''}:{index}:{author}:{stamp}:{text}"
    return "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, seed)).upper() + "}"


def _person_id(author: str) -> str:
    return "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"xlspy:threaded-person:{author}")).upper() + "}"


def _normalize_thread_id(value: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError("threaded comment ids must be non-empty strings")
    value = value.strip()
    if not (value.startswith("{") and value.endswith("}")):
        value = "{" + value + "}"
    return value


def _datetime_text(value: _datetime.datetime) -> str:
    return value.isoformat(timespec="milliseconds")


def _parse_datetime(value: str | None) -> _datetime.datetime | None:
    if not value:
        return None
    try:
        return _datetime.datetime.fromisoformat(value)
    except ValueError:
        return None


def _person_for(author: str, persons: dict[str, Person], by_name: dict[str, Person]) -> Person:
    person = by_name.get(author)
    if person is not None:
        return person
    person = Person(author, _person_id(author), uuid.uuid5(uuid.NAMESPACE_URL, f"xlspy:threaded-user:{author}").hex)
    persons[person.id] = person
    by_name[author] = person
    return person


def collect_people(
    threads: Iterable[ThreadedComment],
    existing: Mapping[str, Person] | None = None,
) -> dict[str, Person]:
    """The people part entries needed by ``threads``, keeping existing ids."""
    people = dict(existing or {})
    by_name: dict[str, Person] = {}
    for person in people.values():
        by_name.setdefault(person.display_name, person)
    for thread in threads:
        for message in thread.messages:
            existing_person: Person | None = message._person
            if existing_person is not None and existing_person.display_name == message.author:
                people.setdefault(existing_person.id, existing_person)
                by_name.setdefault(existing_person.display_name, existing_person)
                continue
            message._person = _person_for(message.author, people, by_name)
    return people


def collect_persons(
    threads: Mapping[str, ThreadedComment],
    existing: Mapping[str, Person] | None = None,
) -> dict[str, Person]:
    """The people part entries needed by ``threads``, keeping existing ids."""
    return collect_people((thread for thread in threads.values() if thread is not None), existing)


def read_persons(data: bytes) -> dict[str, Person]:
    """Read a workbook's people part, keyed by person id."""
    root = safe_fromstring(data)
    ns = namespace_of(root.tag) or THREADED_NS
    people: dict[str, Person] = {}
    for node in root.findall(f"{{{ns}}}person"):
        person_id = node.get("id")
        if not person_id:
            raise InvalidFileError("a person entry has no id")
        people[person_id] = Person(
            node.get("displayName", ""),
            person_id,
            node.get("userId", ""),
            node.get("providerId", ""),
        )
    return people


def persons_xml(persons: Mapping[str, Person]) -> bytes:
    """Encode the workbook's people part; entries are sorted for stable output."""
    items = []
    for person in sorted(persons.values(), key=lambda item: (item.display_name.casefold(), item.id)):
        for value in (person.display_name, person.id, person.user_id, person.provider_id):
            xml_text(value)
        items.append(
            f"<person displayName={quoteattr(xstring_escape(person.display_name))} "
            f"id={quoteattr(xstring_escape(person.id))} "
            f"userId={quoteattr(xstring_escape(person.user_id))} "
            f"providerId={quoteattr(xstring_escape(person.provider_id))}/>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<personList xmlns="{THREADED_NS}" xmlns:x="{MAIN}">{"".join(items)}</personList>'
    ).encode()


def _message_xml(
    ref: str,
    message: ThreadedComment,
    *,
    message_id: str,
    parent_id: str | None,
    person: Person,
    resolved: bool,
) -> str:
    attributes = [
        f'ref="{ref}"',
        f"personId={quoteattr(xstring_escape(person.id))}",
        f"id={quoteattr(xstring_escape(message_id))}",
    ]
    if message.created is not None:
        attributes.append(f'dT="{_datetime_text(message.created)}"')
    if parent_id is not None:
        attributes.append(f"parentId={quoteattr(xstring_escape(parent_id))}")
    if resolved:
        attributes.append('done="1"')
    text = xstring_escape(message.text)
    return (
        f"<threadedComment {' '.join(attributes)}>"
        f"<text{xml_space_attribute(message.text)}>{escape(text)}</text>"
        f"</threadedComment>"
    )


def threaded_comments_xml(
    threads: Mapping[str, ThreadedComment | None],
    *,
    existing: bytes | None = None,
    changed: Collection[str] | None = None,
    persons: Mapping[str, Person] | None = None,
    scope: str = "",
) -> tuple[bytes, dict[str, Person]]:
    """Encode a sheet's conversations, replacing only the refs in ``changed``.

    ``threads`` is the complete mapping for the sheet after the edit. Existing
    messages whose ref is not in ``changed`` keep their original bytes, so
    mentions and rich text that xlspy does not model survive an edit elsewhere.
    """
    if existing is None:
        xml = f'<ThreadedComments xmlns="{THREADED_NS}" xmlns:x="{MAIN}"></ThreadedComments>'
    else:
        xml = existing.decode("utf-8")
    safe_fromstring(xml)
    replacement_refs = {cell_ref(ref) for ref in (changed if changed is not None else threads)}
    seen_refs = set()

    def keep(match: re.Match[str]) -> str:
        node = _fragment_node(xml, match.group(0))
        ref = cell_ref(node.attrib["ref"]) if "ref" in node.attrib else ""
        if ref in replacement_refs:
            seen_refs.add(ref)
            return ""
        return match.group(0)

    xml = element_pattern("threadedComment").sub(keep, xml)
    all_persons = collect_persons({ref: thread for ref, thread in threads.items() if thread is not None}, persons)
    used_ids: set[str] = set()
    for ref in threads:
        if ref not in replacement_refs:
            continue
        thread = threads[ref]
        if thread is None:
            continue
        root_id = _normalize_thread_id(thread.id) if thread.id else ""
        if not root_id:
            root_id = _thread_id(
                scope=scope,
                ref=ref,
                author=thread.author,
                text=thread.text,
                created=thread.created,
                parent=None,
                index=0,
            )
        if root_id in used_ids:
            raise ValueError(f"duplicate threaded comment id {root_id!r}")
        used_ids.add(root_id)
        thread.id = root_id
        root_person = thread._person or _person_for(thread.author, all_persons, {})
        fragments = [
            _message_xml(
                ref,
                thread,
                message_id=root_id,
                parent_id=None,
                person=root_person,
                resolved=thread.resolved,
            )
        ]
        for index, reply in enumerate(thread.replies, start=1):
            reply_id = _normalize_thread_id(reply.id) if reply.id else ""
            if not reply_id:
                reply_id = _thread_id(
                    scope=scope,
                    ref=ref,
                    author=reply.author,
                    text=reply.text,
                    created=reply.created,
                    parent=root_id,
                    index=index,
                )
            if reply_id in used_ids or reply_id == root_id:
                raise ValueError(f"duplicate threaded comment id {reply_id!r}")
            used_ids.add(reply_id)
            reply.id = reply_id
            reply_person = reply._person or _person_for(reply.author, all_persons, {})
            fragments.append(
                _message_xml(
                    ref,
                    reply,
                    message_id=reply_id,
                    parent_id=root_id,
                    person=reply_person,
                    resolved=False,
                )
            )
            # Cache the parent link for a later read/write round trip.
            reply._parent_id = root_id
        xml = append_child(xml, "ThreadedComments", "".join(fragments))
    return xml.encode("utf-8"), all_persons


def read_threaded_comments(
    data: bytes,
    persons: Mapping[str, Person] | None = None,
) -> dict[str, ThreadedComment]:
    """Read a sheet's conversations, keyed by cell reference."""
    root = safe_fromstring(data)
    ns = namespace_of(root.tag) or THREADED_NS
    people = persons or {}
    by_ref: dict[str, list[tuple[ThreadedComment, str | None]]] = {}
    for node in root.findall(f"{{{ns}}}threadedComment"):
        ref = cell_ref(node.attrib["ref"]) if "ref" in node.attrib else ""
        if not ref:
            raise InvalidFileError("a threaded comment has no cell reference")
        message_id = node.get("id")
        if not message_id:
            raise InvalidFileError(f"a threaded comment on {ref} has no id")
        person = people.get(node.get("personId", ""))
        author = person.display_name if person is not None else "Unknown"
        text_node = node.find(f"{{{ns}}}text")
        if text_node is None:
            text = ""
        else:
            # Excel writes plain <text>...</text>; a conversation created by
            # another producer may carry runs (<text><r><t>...</t></r></text>).
            runs = text_node.findall(f"{{{ns}}}t") + text_node.findall(f"{{{ns}}}r/{{{ns}}}t")
            text = "".join(run.text or "" for run in runs) if runs else text_node.text or ""
        message = ThreadedComment(
            xstring_unescape(text),
            author,
            _parse_datetime(node.get("dT")),
            node.get("done") == "1" and node.get("parentId") is None,
            message_id,
        )
        if person is not None:
            message._person = person
        by_ref.setdefault(ref, []).append((message, node.get("parentId")))
    result: dict[str, ThreadedComment] = {}
    for ref, messages in by_ref.items():
        roots = [message for message, parent in messages if parent is None]
        if len(roots) != 1:
            raise InvalidFileError(f"threaded comments on {ref} need exactly one first message")
        thread = roots[0]
        replies = [message for message, parent in messages if parent is not None]
        thread.replies = replies
        for reply in replies:
            reply._parent_id = thread.id
        result[ref] = thread
    return result


def classic_companion(thread: ThreadedComment, root_id: str | None = None) -> Comment:
    """The classic note that links modern Excel to ``thread``.

    Excel hides it from ``Cell.Comment`` because its author starts with
    ``tc={id}``; older Excel versions show its text.
    """
    identifier = root_id or thread.id
    if identifier is None:
        raise ValueError("the conversation needs an id before its classic note can be written")
    identifier = _normalize_thread_id(identifier)
    lines = [
        "[Threaded comment]",
        "",
        "Your version of Excel allows you to read this threaded comment; however, any changes "
        "you make to it will be lost when you open the file in a newer version of Excel. "
        "Learn more: https://go.microsoft.com/fwlink/?linkid=870924",
        "",
        "Comment:",
        "    " + thread.text,
    ]
    for reply in thread.replies:
        lines.extend(["", f"{reply.author}:", "    " + reply.text])
    return Comment("\n".join(lines), f"tc={identifier}")


def is_threaded_companion(comment: Comment) -> bool:
    """Whether a classic note is the compatibility copy of a threaded comment."""
    author = comment.author
    return author.startswith("tc={") and author.endswith("}")


def classic_comments_only(comments: Mapping[str, Comment]) -> dict[str, Comment]:
    """Drop the compatibility copies from a classic comments mapping."""
    return {ref: comment for ref, comment in comments.items() if not is_threaded_companion(comment)}


def find_threaded_part(get: Any, names: Collection[str], sheet_part: str) -> tuple[str, str] | None:
    """The sheet's threaded-comment part and relationship id, if it has one."""
    path = rels_path(sheet_part)
    if path not in names:
        return None
    root = safe_fromstring(get(path))
    found: tuple[str, str] | None = None
    for relationship in root:
        kind = relationship.get("Type", "").rsplit("/", 1)[-1]
        if kind not in {"threadedComment", "threadedComments"}:
            continue
        if relationship.get("TargetMode") == "External":
            raise ValueError("threaded comment relationships cannot be external")
        target = relationship.get("Target")
        if not target or found is not None:
            raise ValueError("unsupported threaded comment relationships")
        found = (resolve_part(sheet_part, target), relationship.attrib["Id"])
    return found


def find_person_part(get: Any, names: Collection[str], workbook_part: str) -> tuple[str, str] | None:
    """The workbook's people part and relationship id, if it has one."""
    path = rels_path(workbook_part)
    if path not in names:
        return None
    root = safe_fromstring(get(path))
    found: tuple[str, str] | None = None
    for relationship in root:
        if relationship.get("Type", "").rsplit("/", 1)[-1] != "person":
            continue
        if relationship.get("TargetMode") == "External":
            raise ValueError("person relationships cannot be external")
        target = relationship.get("Target")
        if not target or found is not None:
            raise ValueError("unsupported person relationships")
        found = (resolve_part(workbook_part, target), relationship.attrib["Id"])
    return found


def add_relationship(
    root: ET.Element,
    *,
    rel_type: str,
    target: str,
    used_ids: Collection[str],
) -> str:
    """Append a relationship with the next free numeric id and return that id."""
    index = 1
    while f"rId{index}" in used_ids:
        index += 1
    rid = f"rId{index}"
    ET.SubElement(root, f"{{{PKG_REL}}}Relationship", {"Id": rid, "Type": rel_type, "Target": target})
    return rid

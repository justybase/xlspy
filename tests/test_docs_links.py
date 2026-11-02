"""Every relative link in the Markdown documentation points to a file that exists and,
when it names an anchor, to a heading of that file (GitHub's anchor rules)."""
from __future__ import annotations

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
DOCS = sorted([ROOT / "README.md", ROOT / "CONTRIBUTING.md", ROOT / "examples" / "README.md",
               *(ROOT / "docs").rglob("*.md")])
LINK = re.compile(r"(?<!!)\[[^\]]*\]\(([^)\s]+)\)")


def _slug(title: str) -> str:
    text = re.sub(r"[`*]", "", title).strip().lower()
    text = re.sub(r"[^\w\- ]", "", text)
    return text.replace(" ", "-")


def _anchors(path: Path) -> set[str]:
    anchors: set[str] = set()
    seen: dict[str, int] = {}
    fence = False
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.lstrip().startswith("```"):
            fence = not fence
            continue
        match = None if fence else re.match(r"^#{1,6} (.+)$", line)
        if match:
            base = _slug(match.group(1))
            count = seen.get(base, 0)
            seen[base] = count + 1
            anchors.add(base if count == 0 else f"{base}-{count}")
    return anchors


def _links(path: Path) -> list[str]:
    text, fence, lines = path.read_text(encoding="utf-8"), False, []
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            fence = not fence
        elif not fence:
            lines.append(line)
    return LINK.findall("\n".join(lines))


@pytest.mark.parametrize("doc", [d for d in DOCS if d.exists()], ids=lambda d: str(d.relative_to(ROOT)))
def test_relative_links_resolve(doc: Path) -> None:
    broken = []
    for url in _links(doc):
        if re.match(r"^(https?:|mailto:)", url):
            continue
        target, _, anchor = url.partition("#")
        path = doc if not target else (doc.parent / target).resolve()
        if not path.exists():
            broken.append(f"{url}: no such file")
        elif anchor and path.suffix == ".md" and anchor not in _anchors(path):
            broken.append(f"{url}: no heading for #{anchor}")
    assert not broken, "\n".join(broken)

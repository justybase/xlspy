"""Every relative link in the Markdown documentation points to a file that exists and,
when it names an anchor, to a heading of that file (GitHub's anchor rules).

README.md is also the project description on PyPI, which resolves relative links against
pypi.org; it links to this repository on GitHub instead, and those links are checked
against the working tree the same way."""
from __future__ import annotations

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
DOCS = sorted([ROOT / "README.md", ROOT / "CONTRIBUTING.md", ROOT / "examples" / "README.md",
               *(ROOT / "docs").rglob("*.md")])
LINK = re.compile(r"(?<!!)\[[^\]]*\]\(([^)\s]+)\)")
ANY_LINK = re.compile(r"!?\[[^\]]*\]\(([^)\s]+)\)")
# Links into this repository's main branch, checked as files of the working tree.
REPO_URLS = ("https://github.com/justybase/xlspy/blob/main/",
             "https://raw.githubusercontent.com/justybase/xlspy/main/")


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
        repo_path = next((url[len(prefix):] for prefix in REPO_URLS if url.startswith(prefix)), None)
        if repo_path is None and re.match(r"^(https?:|mailto:)", url):
            continue
        target, _, anchor = (repo_path if repo_path is not None else url).partition("#")
        if repo_path is not None:
            path = (ROOT / target).resolve()
        else:
            path = doc if not target else (doc.parent / target).resolve()
        if not path.exists():
            broken.append(f"{url}: no such file")
        elif anchor and path.suffix == ".md" and anchor not in _anchors(path):
            broken.append(f"{url}: no heading for #{anchor}")
    assert not broken, "\n".join(broken)


def test_readme_has_no_relative_links_for_pypi() -> None:
    """PyPI shows README.md as the project page, where a relative link or image is a 404."""
    readme = ROOT / "README.md"
    relative = [url for url in ANY_LINK.findall(readme.read_text(encoding="utf-8"))
                if not re.match(r"^(https?:|mailto:|#)", url)]
    assert not relative, relative
    for url in ANY_LINK.findall(readme.read_text(encoding="utf-8")):
        for prefix in REPO_URLS:
            if url.startswith(prefix):
                assert (ROOT / url[len(prefix):].partition("#")[0]).exists(), url

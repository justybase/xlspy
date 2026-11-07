"""Every Python block marked ``<!-- runnable -->`` in docs/learn runs as a script.

The tutorial is part of the product: if an example stops working, the docs build
still succeeds but the page lies. This test runs each marked block in an isolated
temporary directory with ``src`` on ``PYTHONPATH``.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
LEARN = ROOT / "docs" / "learn"
BLOCK = re.compile(r"<!--\s*runnable\s*-->\s*```python\n(.*?)```", re.DOTALL)


def _blocks(path: Path) -> list[str]:
    return [match.group(1) for match in BLOCK.finditer(path.read_text(encoding="utf-8"))]


@pytest.mark.parametrize(
    "doc",
    sorted(LEARN.glob("*.md")),
    ids=lambda path: path.name,
)
def test_tutorial_blocks_run(doc: Path, tmp_path: Path) -> None:
    blocks = _blocks(doc)
    if not blocks:
        pytest.skip(f"{doc.name} has no runnable blocks")
    env = os.environ.copy()
    env["PYTHONPATH"] = str(ROOT / "src") + os.pathsep + env.get("PYTHONPATH", "")
    env.pop("XLSPY_ACCEL", None)
    for index, code in enumerate(blocks, start=1):
        script = tmp_path / f"{doc.stem}_{index}.py"
        script.write_text(code, encoding="utf-8")
        result = subprocess.run(
            [sys.executable, str(script)],
            cwd=tmp_path,
            env=env,
            capture_output=True,
            text=True,
            timeout=120,
        )
        assert result.returncode == 0, (
            f"{doc.name} block {index} failed\n"
            f"--- stdout ---\n{result.stdout}\n"
            f"--- stderr ---\n{result.stderr}"
        )

"""Start excel-sovereign-mcp without relying on PATH or the current directory.

Cursor launched from the Start menu often does not include Python's Scripts
directory, so a bare ``excel-sovereign-mcp`` command exits and the tool list
stays empty. This file puts the repo's ``src`` on ``sys.path`` from its own
location, then speaks stdio.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from excel_sovereign.server import main

if __name__ == "__main__":
    main()

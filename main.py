"""项目根目录的方便入口.

* ``uv run python main.py``        -> 走 CLI
* ``uv run python main.py --gui``  -> 起 GUI
"""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))


def main() -> int:
    if "--gui" in sys.argv:
        sys.argv.remove("--gui")
        from benchmark.gui import main as gui_main

        return gui_main()
    from benchmark.cli import main as cli_main

    return cli_main()


if __name__ == "__main__":
    raise SystemExit(main())

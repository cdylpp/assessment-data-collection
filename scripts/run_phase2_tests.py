from __future__ import annotations

import os
import subprocess
import sys


def main() -> int:
    env = os.environ.copy()
    env.setdefault("QT_QPA_PLATFORM", "offscreen")
    return subprocess.call(
        [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
        env=env,
    )


if __name__ == "__main__":
    raise SystemExit(main())


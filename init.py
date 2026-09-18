#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""프로젝트 루트에서 01_Program/00_main 프로그램을 실행하는 메뉴."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent
SCRIPT_DIR = ROOT_DIR / "01_Program" / "00_main"
EXCLUDED_NAMES = {Path(__file__).name, "__init__.py"}
LAST_PROGRAM_NAME = "차종_미확인_비율_분석.py"


def _programs() -> list[Path]:
    return sorted(
        (
            path
            for path in SCRIPT_DIR.glob("*.py")
            if path.name not in EXCLUDED_NAMES and not path.name.startswith("test_")
        ),
        key=lambda path: (path.name == LAST_PROGRAM_NAME, path.name.lower()),
    )


def _print_menu(programs: list[Path]) -> None:
    print("\n실행할 프로그램을 선택하세요.")
    print("-" * 40)
    for idx, path in enumerate(programs, 1):
        print(f"{idx}. {path.stem}")
    print("0. 종료")


def _run_program(path: Path) -> None:
    print(f"\n[{path.name}] 실행\n")
    try:
        subprocess.run([sys.executable, str(path)], cwd=SCRIPT_DIR)
    except KeyboardInterrupt:
        print("\n프로그램 실행을 중단했습니다.")
    print(f"\n[{path.name}] 종료")


def main() -> int:
    while True:
        programs = _programs()
        if not programs:
            print(f"실행 가능한 Python 프로그램이 없습니다: {SCRIPT_DIR}")
            return 1

        _print_menu(programs)
        choice = input("\n번호 입력: ").strip()

        if choice in {"0", "q", "Q", "quit", "exit"}:
            print("종료합니다.")
            return 0

        if not choice.isdigit():
            print("숫자를 입력해 주세요.")
            continue

        idx = int(choice)
        if not 1 <= idx <= len(programs):
            print(f"1부터 {len(programs)} 사이의 번호를 입력해 주세요.")
            continue

        _run_program(programs[idx - 1])


if __name__ == "__main__":
    raise SystemExit(main())

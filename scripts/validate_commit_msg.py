import re
import sys
from pathlib import Path

ALLOWED_TYPES = "feat|fix|refactor|test|docs|chore|perf"
PATTERN = re.compile(rf"^(?:{ALLOWED_TYPES})(?:\([^)]+\))?!?: .+")


def main() -> int:
    if len(sys.argv) < 2:
        print("commit message file path is required.")
        return 1

    message_path = Path(sys.argv[1])
    if not message_path.exists():
        print(f"commit message file not found: {message_path}")
        return 1

    first_line = message_path.read_text(encoding="utf-8", errors="replace").splitlines()
    subject = first_line[0].strip() if first_line else ""

    if not subject:
        print("empty commit message is not allowed.")
        return 1

    if subject.startswith("Merge ") or subject.startswith("Revert "):
        return 0

    if PATTERN.match(subject):
        return 0

    print("invalid commit message format.")
    print("expected: <type>(optional-scope): <subject>")
    print("allowed types: feat, fix, refactor, test, docs, chore, perf")
    print(f"current: {subject}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

import re
import subprocess

PATTERN = re.compile(r"^(develop|main|(feature|fix|hotfix)/\S+)$")


def current_branch() -> str:
    output = subprocess.check_output(
        ["git", "rev-parse", "--abbrev-ref", "HEAD"],
        text=True,
        encoding="utf-8",
    )
    return output.strip()


def main() -> int:
    branch = current_branch()

    if branch == "HEAD":
        return 0

    if PATTERN.match(branch):
        return 0

    print("invalid branch name.")
    print("allowed: develop, main, feature/*, fix/*, hotfix/*")
    print(f"current: {branch}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

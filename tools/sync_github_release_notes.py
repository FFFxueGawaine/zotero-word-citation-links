from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


GITHUB_API_BASE = "https://api.github.com"
GITHUB_ACCEPT = "application/vnd.github+json"
GITHUB_API_VERSION = "2022-11-28"
REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CHANGELOG = REPO_ROOT / "CHANGELOG.md"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Sync a GitHub release body from a UTF-8 changelog section so "
            "Chinese release notes do not depend on terminal encoding."
        )
    )
    parser.add_argument("--repo", required=True, help="GitHub repo in owner/name format.")
    parser.add_argument("--tag", required=True, help="Release tag, for example v0.4.0.")
    parser.add_argument(
        "--changelog",
        type=Path,
        default=DEFAULT_CHANGELOG,
        help="UTF-8 changelog path. Default: CHANGELOG.md in the repo root.",
    )
    parser.add_argument(
        "--token-file",
        type=Path,
        help="Optional UTF-8 text file containing a GitHub token.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the extracted release notes without updating GitHub.",
    )
    return parser.parse_args()


def read_token(token_file: Path | None) -> str:
    env_token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if env_token:
        return env_token.strip()

    if token_file:
        return token_file.read_text(encoding="utf-8").strip()

    raise SystemExit(
        "Missing GitHub token. Provide --token-file or set GITHUB_TOKEN / GH_TOKEN."
    )


def extract_release_body(changelog_path: Path, tag: str) -> str:
    text = changelog_path.read_text(encoding="utf-8")
    header_pattern = re.compile(
        rf"^##\s+{re.escape(tag)}(?:\s+-.*)?\s*$",
        flags=re.MULTILINE,
    )
    header_match = header_pattern.search(text)
    if not header_match:
        raise SystemExit(f"Could not find a changelog section for {tag} in {changelog_path}.")

    next_header_match = re.search(r"^##\s+", text[header_match.end() :], flags=re.MULTILINE)
    if next_header_match:
        section_end = header_match.end() + next_header_match.start()
    else:
        section_end = len(text)

    body = text[header_match.end() : section_end].strip()
    if not body:
        raise SystemExit(f"The changelog section for {tag} is empty in {changelog_path}.")

    return body


def github_request(
    method: str,
    url: str,
    token: str,
    payload: dict[str, object] | None = None,
) -> dict[str, object]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": GITHUB_ACCEPT,
        "X-GitHub-Api-Version": GITHUB_API_VERSION,
        "User-Agent": "zotero-word-citation-links-release-sync",
    }
    data = None
    if payload is not None:
        headers["Content-Type"] = "application/json; charset=utf-8"
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        error_text = exc.read().decode("utf-8", errors="replace")
        raise SystemExit(f"GitHub API request failed ({exc.code}): {error_text}") from exc


def get_release(repo: str, tag: str, token: str) -> dict[str, object]:
    encoded_tag = urllib.parse.quote(tag, safe="")
    url = f"{GITHUB_API_BASE}/repos/{repo}/releases/tags/{encoded_tag}"
    return github_request("GET", url, token)


def update_release_body(repo: str, release_id: int, token: str, body: str) -> dict[str, object]:
    url = f"{GITHUB_API_BASE}/repos/{repo}/releases/{release_id}"
    return github_request("PATCH", url, token, payload={"body": body})


def main() -> int:
    args = parse_args()
    body = extract_release_body(args.changelog, args.tag)

    if args.dry_run:
        print(body)
        return 0

    token = read_token(args.token_file)
    release = get_release(args.repo, args.tag, token)
    release_id = int(release["id"])
    updated_release = update_release_body(args.repo, release_id, token, body)

    print(f"Updated release notes for {args.tag}.")
    print(updated_release["html_url"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

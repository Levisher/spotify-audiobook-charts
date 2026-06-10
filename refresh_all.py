#!/usr/bin/env python3
"""Run refresh_one.py for every playlist in playlists.txt.

Usage:
    python3 refresh_all.py            # process all, skip ones already done today
    python3 refresh_all.py --force     # re-process everything
    python3 refresh_all.py --cache     # use the Firecrawl markdown cache (no credits)
"""
import argparse
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).parent
DOWNLOADS = Path(os.environ["OUTPUT_DIR"]) if os.environ.get("OUTPUT_DIR") \
    else Path.home() / "Downloads"
TODAY = date.today().isoformat()


def parse_playlists() -> list[tuple[str, str]]:
    out = []
    for line in (ROOT / "playlists.txt").read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("\t")
        if len(parts) >= 2:
            out.append((parts[0].strip(), parts[1].strip()))
        else:
            out.append((parts[0].strip(), ""))
    return out


def pdf_done_today(tab: str) -> bool:
    safe = tab.replace(" ", "_").replace("&", "and")
    return (DOWNLOADS / f"audiobooks_{safe}_{TODAY}.pdf").exists()


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--force", action="store_true",
                   help="re-process even if today's PDF already exists")
    p.add_argument("--cache", action="store_true",
                   help="use Firecrawl markdown cache (no credits)")
    args = p.parse_args()

    plists = parse_playlists()
    print(f"Found {len(plists)} playlists in playlists.txt")
    print(f"Today: {TODAY}\n")

    successes = []
    failures = []
    for pid, tab in plists:
        if not args.force and pdf_done_today(tab):
            print(f"[skip] {tab} — today's PDF already exists")
            continue
        print(f"\n{'=' * 60}\n>>> {tab} ({pid})\n{'=' * 60}")
        cmd = [sys.executable, str(ROOT / "refresh_one.py"), pid]
        if tab:
            cmd += ["--tab", tab]
        if args.cache:
            cmd += ["--cache"]
        r = subprocess.run(cmd)
        if r.returncode == 0:
            successes.append(tab)
        else:
            failures.append(tab)
            print(f"!! failed: {tab}", file=sys.stderr)

    print(f"\n\n{'=' * 60}")
    print(f"Summary: {len(successes)} succeeded, {len(failures)} failed")
    print(f"{'=' * 60}")
    for tab in successes:
        safe = tab.replace(" ", "_").replace("&", "and")
        print(f"  ✓ {tab} -> ~/Downloads/audiobooks_{safe}_{TODAY}.pdf")
    for tab in failures:
        print(f"  ✗ {tab}")
    # Exit non-zero if any playlist failed so CI marks the run as failed
    # (otherwise GitHub Actions reports green even when every scrape 402s).
    if failures:
        sys.exit(1)


if __name__ == "__main__":
    main()

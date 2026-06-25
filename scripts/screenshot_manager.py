"""
screenshot_manager.py — List, validate, and organize FPK screenshots.

Used by the dashboard backend to discover screenshots for a given period.
Also usable as a CLI to check what screenshots exist.

Usage:
  python scripts/screenshot_manager.py                     # list current period
  python scripts/screenshot_manager.py --year 2026 --month 5  # list specific period
"""

import sys
import os
import re
import yaml
from pathlib import Path
from datetime import date

if sys.platform == "win32":
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleOutputCP(65001)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent.parent
SCREENSHOTS_DIR = BASE_DIR / "competitors"
CONFIG_PATH = BASE_DIR / "screenshots_config.yaml"


def load_config():
    """Load screenshot config YAML."""
    if not CONFIG_PATH.exists():
        return {"screenshots": []}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {"screenshots": []}


def get_period_dir(year: int, month: int) -> Path:
    """Get the screenshot directory for a specific year/month."""
    return SCREENSHOTS_DIR / f"{year}_{month:02d}"


def list_screenshots(year: int, month: int) -> list[dict]:
    """
    List all screenshots for a given period.

    Returns a list of dicts:
      [{"file": "competitor_overview.png", "path": "...", "label": "...", "description": "...", "configured": True}]
    """
    period_dir = get_period_dir(year, month)
    config = load_config()
    configured = {s["file"]: s for s in config.get("screenshots", [])}

    if not period_dir.exists():
        return []

    results = []

    # Check configured screenshots first (preserves order)
    for entry in config.get("screenshots", []):
        file_stem = entry["file"]
        for ext in [".png", ".jpg", ".jpeg", ".webp"]:
            filepath = period_dir / f"{file_stem}{ext}"
            if filepath.exists():
                results.append({
                    "file": filepath.name,
                    "path": str(filepath.relative_to(BASE_DIR)),
                    "label": entry.get("label", file_stem),
                    "description": entry.get("description", ""),
                    "source": entry.get("source", ""),
                    "configured": True,
                    "size_kb": round(filepath.stat().st_size / 1024, 1),
                })
                break

    # Check for additional (unconfigured) screenshots
    if period_dir.exists():
        for filepath in sorted(period_dir.iterdir()):
            if filepath.suffix.lower() not in [".png", ".jpg", ".jpeg", ".webp"]:
                continue
            stem = filepath.stem
            if stem not in configured:
                results.append({
                    "file": filepath.name,
                    "path": str(filepath.relative_to(BASE_DIR)),
                    "label": stem.replace("_", " ").title(),
                    "description": "",
                    "source": "",
                    "configured": False,
                    "size_kb": round(filepath.stat().st_size / 1024, 1),
                })

    return results


def list_all_periods() -> list[tuple[int, int]]:
    """List all periods that have screenshot folders."""
    if not SCREENSHOTS_DIR.exists():
        return []
    periods = []
    pattern = re.compile(r"^(\d{4})_(\d{2})$")
    for d in SCREENSHOTS_DIR.iterdir():
        if d.is_dir():
            m = pattern.match(d.name)
            if m:
                periods.append((int(m.group(1)), int(m.group(2))))
    return sorted(periods, reverse=True)


def print_report(year: int = None, month: int = None):
    """Print a report of screenshots for a period (or all periods)."""
    config = load_config()
    expected = [s["file"] for s in config.get("screenshots", [])]

    if year and month:
        periods = [(year, month)]
    else:
        periods = list_all_periods()
        if not periods:
            print("No screenshot folders found.")
            print(f"Expected location: {SCREENSHOTS_DIR}")
            return

    for y, m in periods:
        screenshots = list_screenshots(y, m)
        found_names = [Path(s["file"]).stem for s in screenshots]
        missing = [name for name in expected if name not in found_names]

        print(f"\n{'='*50}")
        print(f"  Period: {y}-{m:02d}")
        print(f"  Folder: {get_period_dir(y, m)}")
        print(f"  Found: {len(screenshots)} / {len(expected)} configured")
        print(f"{'='*50}")

        if screenshots:
            for s in screenshots:
                tag = "" if s["configured"] else " [extra]"
                print(f"  [OK]   {s['file']:40s} ({s['size_kb']:>7.1f} KB){tag}")
                if s["description"]:
                    print(f"         -> {s['description']}")

        if missing:
            print(f"\n  Missing screenshots:")
            for name in missing:
                label = configured_label(config, name)
                print(f"  [MISS] {name}.png")
                if label:
                    print(f"         -> {label}")

    print(f"\nTotal periods: {len(periods)}")


def configured_label(config: dict, file_stem: str) -> str:
    for s in config.get("screenshots", []):
        if s["file"] == file_stem:
            return s.get("label", "")
    return ""


if __name__ == "__main__":
    today = date.today()
    default_year = today.year if today.month > 1 else today.year - 1
    default_month = today.month - 1 if today.month > 1 else 12

    import argparse
    parser = argparse.ArgumentParser(description="FPK Screenshot Manager")
    parser.add_argument("--year", type=int, default=None, help="Year (default: all)")
    parser.add_argument("--month", type=int, default=None, help="Month (default: all)")
    args = parser.parse_args()

    print_report(args.year, args.month)

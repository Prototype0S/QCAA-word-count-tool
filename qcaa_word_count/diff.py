"""Compute a per-section diff between two runs."""
from __future__ import annotations

import difflib
import json
from pathlib import Path

from .output import bold, cyan, green, grey, red, yellow


def load_saved_sections(path: str) -> tuple[list[dict], int]:
    """Return (sections, total_at_save) from a saved decisions file."""
    try:
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return [], 0
    return payload.get("sections", []), payload.get("total_at_save", 0)


def match_sections(old: list[dict], new: list[dict]) -> list[tuple]:
    """Pair up old and new sections by fuzzy name match.

    Returns a list of (old_section_or_None, new_section_or_None) tuples.
    Sections with no match get one None.
    """
    old_names = [s.get("heading", "") for s in old]
    new_names = [s.get("heading", "") for s in new]

    pairs: list[tuple] = []
    used_new = set()

    for old_idx, old_section in enumerate(old):
        old_name = old_names[old_idx]
        # Find best fuzzy match among unused new sections
        best_idx = None
        best_ratio = 0.0
        for new_idx, new_name in enumerate(new_names):
            if new_idx in used_new:
                continue
            ratio = difflib.SequenceMatcher(None, old_name, new_name).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_idx = new_idx
        if best_idx is not None and best_ratio >= 0.6:
            used_new.add(best_idx)
            pairs.append((old_section, new[best_idx]))
        else:
            pairs.append((old_section, None))

    for new_idx, new_section in enumerate(new):
        if new_idx not in used_new:
            pairs.append((None, new_section))

    return pairs


def print_diff(path: str, current_sections: list[dict],
               current_total: int) -> None:
    """Print a human-readable diff against a saved run."""
    old_sections, old_total = load_saved_sections(path)
    if not old_sections and not old_total:
        return

    print()
    print(bold("Comparison with last run:"))
    print(f"  Last run:  {grey(str(old_total))} words")
    print(f"  This run:  {bold(str(current_total))} words")
    delta = current_total - old_total
    if delta > 0:
        print(green(f"  Change:    +{delta}"))
    elif delta < 0:
        print(red(f"  Change:    {delta}"))
    else:
        print(grey("  Change:    no change"))
    print()

    pairs = match_sections(old_sections, current_sections)
    print(bold("  Per section:"))
    for old, new in pairs:
        if old and new:
            name = new.get("heading") or old.get("heading") or "(untitled)"
            old_w = old.get("words", 0)
            new_w = new.get("words", 0)
            d = new_w - old_w
            if d > 0:
                marker = green(f"+{d:>5}")
            elif d < 0:
                marker = red(f"{d:>6}")
            else:
                marker = grey("  =")
            print(f"    {name:<40}  {new_w:>5}  {marker}")
        elif new:
            name = new.get("heading", "(untitled)")
            print(f"    {name:<40}  {new.get('words', 0):>5}  {green('(new)')}")
        elif old:
            name = old.get("heading", "(untitled)")
            print(f"    {name:<40}  {grey('(removed)')}")
    print()
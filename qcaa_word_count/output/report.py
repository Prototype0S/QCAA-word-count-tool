"""Plain-text report."""
from __future__ import annotations

from ..model import Decision, Document
from .ansi import bold, bold_green, green, grey, red, yellow


def print_report(document: Document, result: dict) -> None:
    print()
    print("=" * 60)
    print(bold("QCAA WORD COUNT REPORT"))
    print(f"File: {document.filename}")
    print("=" * 60)
    print()
    print(f"  {bold_green('FINAL WORD COUNT:')} {bold_green(str(result['total']))}")
    print()

    # [NEW] Regions detected — makes a bad region assignment visible
    regions_seen = {}
    for b in document.blocks:
        regions_seen[b.region.value] = regions_seen.get(b.region.value, 0) + 1
    if any(r != "body" for r in regions_seen):
        print(bold("Regions detected:"))
        for name in sorted(regions_seen):
            n = regions_seen[name]
            if name == "body":
                marker = green
            elif name in ("title_page", "references", "contents", "abstract"):
                marker = red
            else:
                marker = yellow
            print(f"  {marker(name):<16} {n:>4} block(s)")
        print()

    if result["buckets"]:
        print(bold("Excluded / flagged (not counted):"))
        samples = result.get("bucket_samples", {})
        for key in sorted(result["buckets"]):
            col = yellow if (key.startswith("flagged:") or "unresolved" in key) else red
            count = result["buckets"][key]
            sample_str = ""
            if samples.get(key):
                parts = [f'"{t}"' for t in samples[key]]
                if count > len(samples[key]):
                    parts.append("...")
                sample_str = "  " + grey("(" + ", ".join(parts) + ")")
            print(f"  {col(key):<40} {count:>5} word(s){sample_str}")
        print()

    if result.get("tables"):
        print(bold("Tables:"))
        for t in result["tables"]:
            label = {
                "count_all": "counted whole",
                "exclude_all": "excluded whole",
                "headers_only": "headers only",
                "unresolved": "unresolved",
            }.get(t["answer"], t["answer"])
            line = (f"  Table {t['index']}: suggested={t['suggestion']:<12} "
                    f"chosen={label:<15} "
                    f"counted={t['counted']:<4} excluded={t['excluded']}")
            if t["answer"] == "count_all":
                print(green(line))
            elif t["answer"] in ("exclude_all", "headers_only"):
                print(red(line))
            else:
                print(yellow(line))
        print()

    cit = result.get("citations", {})
    if cit.get("auto") or cit.get("manual"):
        print(bold("In-text citations:"))
        if cit["auto"]:
            print(green(f"  Auto-detected (APA 7):   {cit['auto']:>4} "
                       f"({cit['auto_words']} words)"))
        if cit["manual"]:
            print(green(f"  Manually added:          {cit['manual']:>4} "
                       f"({cit['manual_words']} words)"))
        print()

    fn = result.get("footnotes", {})
    if fn.get("counted") or fn.get("excluded"):
        print(bold("Footnotes / endnotes:"))
        if fn["counted"]:
            print(green(f"  Counted (commentary):     {fn['counted']:>4} "
                       f"({fn['counted_words']} words)"))
        if fn["excluded"]:
            print(red(f"  Excluded (bibliographic): {fn['excluded']:>4} "
                     f"({fn['excluded_words']} words)"))
        print()

    if result.get("unresolved_footnotes"):
        print(yellow(f"  {result['unresolved_footnotes']} footnote(s) "
                     f"unresolved and NOT counted."))

    if result.get("math_objects"):
        print(red(f"Math objects (OMML):       {result['math_objects']:>4} "
                 f"(excluded)"))

    if document.drawings:
        msg = f"Drawings/shapes:          {document.drawings:>4}"
        if document.drawings_with_text:
            msg += f" ({document.drawings_with_text} contain text — not extracted)"
        print(grey(msg))
    if document.text_boxes:
        print(grey(f"Text boxes:               {len(document.text_boxes):>4}"))

    print()
    print(grey("Citation detection uses APA 7 patterns. Review the report and"))
    print(grey("use manual entry if any were missed."))
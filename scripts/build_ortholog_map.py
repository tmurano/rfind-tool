#!/usr/bin/env python3
"""
Download RGD ortholog data and build Mouse/Rat → Human symbol mapping JSON.

Source: RGD_ORTHOLOGS_Ensembl.txt
  Contains columns: RAT_GENE_SYMBOL, HUMAN_ORTHOLOG_SYMBOL, MOUSE_ORTHOLOG_SYMBOL

Output: ../ortholog_map.json
  - Keys: mouse or rat gene symbol (uppercased)
  - Values: human gene symbol (uppercased)
  - Only includes entries where uppercased symbols differ (mouse AQP4 = human AQP4 → skip)
  - Ambiguous mappings (one symbol → multiple human symbols) are excluded

Usage:
  python build_ortholog_map.py
"""

import csv
import json
import sys
import os
import io
import urllib.request
from collections import defaultdict

RGD_URL = "https://download.rgd.mcw.edu/data_release/orthologs/RGD_ORTHOLOGS_Ensembl.txt"
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "ortholog_map.json")


def main():
    print("Downloading RGD ortholog data...", file=sys.stderr)
    req = urllib.request.Request(RGD_URL, headers={"User-Agent": "RFind-Tool/1.0"})
    with urllib.request.urlopen(req, timeout=120) as resp:
        raw = resp.read().decode("utf-8")
    print(f"Downloaded {len(raw):,} bytes", file=sys.stderr)

    # Parse: skip comment lines starting with #
    lines = raw.split("\n")
    header_idx = None
    for i, line in enumerate(lines):
        if line.startswith("RAT_GENE_SYMBOL"):
            header_idx = i
            break

    if header_idx is None:
        print("ERROR: Could not find header row", file=sys.stderr)
        sys.exit(1)

    reader = csv.DictReader(lines[header_idx:], delimiter="\t")

    # Collect all non-human symbol → set of human symbols
    reverse = defaultdict(set)
    row_count = 0

    for row in reader:
        human = row.get("HUMAN_ORTHOLOG_SYMBOL", "").strip().upper()
        if not human:
            continue
        row_count += 1

        # Mouse → Human
        mouse = row.get("MOUSE_ORTHOLOG_SYMBOL", "").strip().upper()
        if mouse and mouse != human:
            reverse[mouse].add(human)

        # Rat → Human
        rat = row.get("RAT_GENE_SYMBOL", "").strip().upper()
        if rat and rat != human:
            reverse[rat].add(human)

    # Build mapping, excluding ambiguous (one symbol → multiple human symbols)
    mapping = {}
    collisions = []
    for key, humans in sorted(reverse.items()):
        if len(humans) == 1:
            mapping[key] = next(iter(humans))
        else:
            collisions.append((key, sorted(humans)))

    # Write JSON
    output_path = os.path.abspath(OUTPUT_PATH)
    with open(output_path, "w") as f:
        json.dump(mapping, f, separators=(",", ":"), ensure_ascii=True)

    file_size = os.path.getsize(output_path)
    print(f"\nRows processed: {row_count:,}", file=sys.stderr)
    print(f"Mapping entries: {len(mapping):,}", file=sys.stderr)
    print(f"  (symbols that differ when uppercased between species)", file=sys.stderr)
    print(f"Ambiguous excluded: {len(collisions):,}", file=sys.stderr)
    print(f"Output: {output_path} ({file_size:,} bytes / {file_size/1024:.1f} KB)", file=sys.stderr)

    if collisions:
        print(f"\nFirst 20 ambiguous mappings:", file=sys.stderr)
        for k, v in collisions[:20]:
            print(f"  {k} -> {v}", file=sys.stderr)

    # Stats
    mouse_only = sum(1 for k, v in mapping.items() if k != v)
    print(f"\nTotal non-identity mappings: {mouse_only:,}", file=sys.stderr)


if __name__ == "__main__":
    main()

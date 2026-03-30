#!/usr/bin/env python3
"""
Download HGNC gene data and build a flat alias->official symbol mapping JSON.

Output: ../hgnc_map.json
  - Keys: previous symbols, alias symbols, ENSEMBL IDs (all uppercased)
  - Values: current approved symbol (uppercased)
  - Ambiguous aliases (mapping to multiple symbols) are excluded

Usage:
  python build_hgnc_map.py
"""

import csv
import json
import sys
import os
import io
import urllib.request

HGNC_URL = (
    "https://www.genenames.org/cgi-bin/download/custom?"
    "col=gd_app_sym&col=gd_prev_sym&col=gd_aliases&col=gd_pub_ensembl_id"
    "&status=Approved&hgnc_dbtag=on&order_by=gd_app_sym_sort&format=text&submit=submit"
)

OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "hgnc_map.json")


def main():
    print("Downloading HGNC data...", file=sys.stderr)
    req = urllib.request.Request(HGNC_URL, headers={"User-Agent": "RFinD-Tool/1.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        raw = resp.read().decode("utf-8")
    print(f"Downloaded {len(raw):,} bytes", file=sys.stderr)

    reader = csv.DictReader(io.StringIO(raw), delimiter="\t")

    # Collect all alias -> set of official symbols
    from collections import defaultdict
    reverse = defaultdict(set)
    gene_count = 0

    for row in reader:
        official = row.get("Approved symbol", "").strip().upper()
        if not official:
            continue
        gene_count += 1

        # Previous symbols (comma-separated)
        for prev in row.get("Previous symbols", "").split(","):
            prev = prev.strip().upper()
            if prev and prev != official:
                reverse[prev].add(official)

        # Alias symbols (comma-separated)
        for alias in row.get("Alias symbols", "").split(","):
            alias = alias.strip().upper()
            if alias and alias != official:
                reverse[alias].add(official)

        # Ensembl gene ID
        ensembl = row.get("Ensembl gene ID", "").strip().upper()
        if ensembl and ensembl != official:
            reverse[ensembl].add(official)

    # Build mapping, excluding ambiguous aliases
    mapping = {}
    collisions = []
    for key, officials in sorted(reverse.items()):
        if len(officials) == 1:
            mapping[key] = next(iter(officials))
        else:
            collisions.append((key, sorted(officials)))

    # Write JSON
    output_path = os.path.abspath(OUTPUT_PATH)
    with open(output_path, "w") as f:
        json.dump(mapping, f, separators=(",", ":"), ensure_ascii=True)

    file_size = os.path.getsize(output_path)
    print(f"\nApproved genes: {gene_count:,}", file=sys.stderr)
    print(f"Mapping entries: {len(mapping):,}", file=sys.stderr)
    print(f"Ambiguous aliases excluded: {len(collisions):,}", file=sys.stderr)
    print(f"Output: {output_path} ({file_size:,} bytes / {file_size/1024/1024:.1f} MB)", file=sys.stderr)

    if collisions:
        print(f"\nFirst 20 ambiguous aliases:", file=sys.stderr)
        for k, v in collisions[:20]:
            print(f"  {k} -> {v}", file=sys.stderr)


if __name__ == "__main__":
    main()

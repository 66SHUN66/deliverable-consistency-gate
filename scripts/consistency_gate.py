#!/usr/bin/env python3
"""Consistency gate for a deliverable set: deck, document, workbook.

Reads .pptx / .docx / .xlsx without third-party libraries (zip + XML), pulls
numbers, tables, headings and placeholders, and reports where the same claim
carries two different values.

Usage:
    python scripts/consistency_gate.py --input deliverable/ --out out/
    python scripts/consistency_gate.py --input deck.pptx --input memo.docx --out out/
"""

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

PLACEHOLDERS = re.compile(r"\b(TODO|TBD|FIXME|XXX|Lorem ipsum|\[insert[^\]]*\]|<<[^>]*>>)\b", re.I)
FIG_REF = re.compile(r"\b(Figure|Fig\.|Table|Exhibit|Chart)\s+(\d+)", re.I)
NUM = re.compile(
    r"(?P<pre>(?P<label>[A-Za-z][A-Za-z0-9 \-/&']{0,38}?)\s*)?"
    r"(?P<cur>[$€£¥])?\s*(?P<num>\d[\d,]*(?:\.\d+)?)\s*"
    r"(?P<unit>%|bps|M|K|B|billion|million|thousand|bn|mn)?\b"
)
YEAR = re.compile(r"^(19|20)\d\d$")
UNIT_SCALE = {"K": 1e3, "thousand": 1e3, "M": 1e6, "million": 1e6, "mn": 1e6,
              "B": 1e9, "billion": 1e9, "bn": 1e9, "": 1.0}
STOP = {"reached", "totaled", "totalled", "hit", "grew", "rose", "grow", "to", "about", "approximately", "reported", "posted", "recorded", "came", "up", "down", "than", "the", "a", "an", "of", "in", "on", "for", "to", "and", "was", "were", "is",
        "are", "at", "by", "with", "as", "from", "that", "this", "its", "our", "we"}


def local(tag):
    return tag.split("}")[-1]


def read_xml(zf, name):
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


# ---------------------------------------------------------------- extractors


def pptx_units(path):
    units, meta = [], {"figures": 0, "tables": 0}
    with zipfile.ZipFile(path) as zf:
        slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                        key=lambda n: int(re.findall(r"\d+", n)[-1]))
        for name in slides:
            idx = int(re.findall(r"\d+", name)[-1])
            root = read_xml(zf, name)
            if root is None:
                continue
            for pi, p in enumerate(root.iter("{%s}p" % NS["a"]), 1):
                text = "".join(t.text or "" for t in p.iter("{%s}t" % NS["a"])).strip()
                if text:
                    units.append({"file": path, "where": "slide %d · para %d" % (idx, pi), "text": text})
            meta["figures"] += len(list(root.iter("{%s}pic" % NS["p"])))
            meta["tables"] += len(list(root.iter("{%s}tbl" % NS["a"])))
    return units, meta


def docx_units(path):
    units, meta = [], {"figures": 0, "tables": 0, "headings": []}
    with zipfile.ZipFile(path) as zf:
        root = read_xml(zf, "word/document.xml")
        if root is None:
            return units, meta
        for pi, p in enumerate(root.iter("{%s}p" % NS["w"]), 1):
            text = "".join(t.text or "" for t in p.iter("{%s}t" % NS["w"])).strip()
            if not text:
                continue
            style = ""
            ppr = p.find("{%s}pPr" % NS["w"])
            if ppr is not None:
                ps = ppr.find("{%s}pStyle" % NS["w"])
                if ps is not None:
                    style = ps.get("{%s}val" % NS["w"]) or ""
            units.append({"file": path, "where": "para %d" % pi, "text": text, "style": style})
            if style.lower().startswith("heading"):
                meta["headings"].append(text)
        meta["tables"] = len(list(root.iter("{%s}tbl" % NS["w"])))
        meta["figures"] = len([n for n in zf.namelist() if n.startswith("word/media/")])
    return units, meta


def _xlsx_sheets(zf):
    wb = read_xml(zf, "xl/workbook.xml")
    rels = read_xml(zf, "xl/_rels/workbook.xml.rels")
    rid_map = {}
    if rels is not None:
        for rel in rels:
            rid_map[rel.get("Id")] = rel.get("Target")
    out = []
    if wb is None:
        return out
    for sh in wb.iter("{%s}sheet" % NS["x"]):
        rid = sh.get("{%s}id" % NS["r"])
        target = rid_map.get(rid, "")
        target = target.lstrip("/")
        name = target if target.startswith("xl/") else "xl/" + target
        out.append((sh.get("name"), name))
    return out


def xlsx_units(path):
    units, meta, grids = [], {"figures": 0, "tables": 0}, {}
    with zipfile.ZipFile(path) as zf:
        shared = []
        sroot = read_xml(zf, "xl/sharedStrings.xml")
        if sroot is not None:
            for si in sroot.iter("{%s}si" % NS["x"]):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["x"])))
        for sheet_name, part in _xlsx_sheets(zf):
            root = read_xml(zf, part)
            if root is None:
                continue
            grid = {}
            for c in root.iter("{%s}c" % NS["x"]):
                ref = c.get("r") or ""
                t = c.get("t")
                v = c.find("{%s}v" % NS["x"])
                raw = v.text if (v is not None and v.text is not None) else None
                if raw is None:
                    is_ = c.find("{%s}is" % NS["x"])
                    if is_ is not None:
                        raw = "".join(tt.text or "" for tt in is_.iter("{%s}t" % NS["x"]))
                if raw is None:
                    continue
                if t == "s":
                    try:
                        text = shared[int(raw)]
                    except (ValueError, IndexError):
                        text = raw
                    val = None
                else:
                    text = raw
                    try:
                        val = float(raw)
                    except ValueError:
                        val = None
                units.append({"file": path, "where": "%s!%s" % (sheet_name, ref), "text": text, "value": val})
                m = re.match(r"([A-Z]+)(\d+)", ref)
                if m:
                    grid[(int(m.group(2)), m.group(1))] = (val, text)
            grids[sheet_name] = grid
    meta["tables"] = 0
    return units, meta, grids


# ---------------------------------------------------------------- number model


def normalise(token):
    """Return (value_in_base_units, dimension) or None."""
    num = float(token.group("num").replace(",", ""))
    unit = (token.group("unit") or "").lower()
    cur = token.group("cur") or ""
    if unit == "%":
        return num, "pct"
    if unit == "bps":
        return num / 100.0, "pct"
    if unit in UNIT_SCALE:
        return num * UNIT_SCALE[unit], "money" if cur else "count"
    return num, ("money" if cur else "count")


def key_from(label):
    if not label:
        return None
    words = [w.strip(".,:;").lower() for w in re.split(r"\s+", label.strip())]
    words = [w for w in words if w and w not in STOP]
    if len(words) < 1:
        return None
    return " ".join(words[-3:])


def collect_numbers(units):
    rows = []
    for u in units:
        for m in NUM.finditer(u["text"]):
            if YEAR.match(m.group("num")):
                continue
            if not m.group("label") and not m.group("unit") and not m.group("cur"):
                continue
            val = normalise(m)
            if val is None:
                continue
            key = key_from(m.group("label"))
            rows.append({
                "key": key,
                "raw": ((m.group("cur") or "") + m.group("num") + (m.group("unit") or "")).strip(),
                "value": val[0],
                "dim": val[1],
                "where": u["where"],
                "file": u["file"],
                "context": u["text"][:150],
                "decimals": len(m.group("num").split(".")[1]) if "." in m.group("num") else 0,
            })
    return rows


def conflicts(numbers):
    out = []
    groups = {}
    for r in numbers:
        if not r["key"] or len(r["key"]) < 2:
            continue
        groups.setdefault((r["key"], r["dim"]), []).append(r)
    for (key, dim), rows in groups.items():
        values = [r["value"] for r in rows]
        hi, lo = max(values), min(values)
        if hi == 0:
            continue
        rel = (hi - lo) / abs(hi)
        tol = max(0.005, 0.5 * 10 ** (-min(r["decimals"] for r in rows)))
        if rel > tol and hi - lo > tol:
            a, b = rows[0], max(rows, key=lambda r: abs(r["value"] - rows[0]["value"]))
            out.append({
                "code": "NUMBER_CONFLICT",
                "severity": "warning",
                "subject": key,
                "message": "\"%s\" carries %s and %s — a %.1f%% difference" % (key, a["raw"], b["raw"], rel * 100),
                "where": [a["where"], b["where"]],
                "files": [a["file"], b["file"]],
                "context": [a["context"], b["context"]],
            })
    return out


# ---------------------------------------------------------------- table checks


def table_checks(grids, path):
    out = []
    for sheet, grid in grids.items():
        rows = sorted({r for r, _ in grid.keys()})
        cols = sorted({c for _, c in grid.keys()})
        for r in rows:
            label = ""
            for c in cols[:2]:
                cell = grid.get((r, c))
                if cell and cell[1]:
                    label = str(cell[1]).strip().lower()
                    break
            if not label or not re.search(r"\b(total|totals|sum|subtotal|grand total)\b", label):
                continue
            for c in cols:
                total_cell = grid.get((r, c))
                if not total_cell or total_cell[0] is None:
                    continue
                parts = []
                for rr in range(r - 1, max(0, r - 6), -1):
                    cell = grid.get((rr, c))
                    if cell is None or cell[0] is None:
                        break
                    parts.append(cell[0])
                if len(parts) < 2:
                    continue
                s = sum(parts)
                total = total_cell[0]
                if abs(total - s) > max(0.01, abs(s) * 0.005):
                    out.append({
                        "code": "TOTAL_MISMATCH",
                        "severity": "blocker",
                        "subject": "%s!%s%d" % (sheet, c, r),
                        "message": "row labelled \"%s\" reads %g, the %d rows above sum to %g (difference %g)"
                                   % (label, total, len(parts), s, total - s),
                        "where": ["%s!%s%d" % (sheet, c, r)],
                        "files": [path],
                        "context": [label],
                    })
    return out


def pct_checks(grids, path):
    """A column of shares that is meant to total 100% and does not."""
    out = []
    for sheet, grid in grids.items():
        cols = sorted({c for _, c in grid.keys()})
        for c in cols:
            vals = [(r, grid[(r, c)][0]) for r in sorted({r for r, _ in grid.keys()}) if (r, c) in grid]
            nums = [v for _, v in vals if v is not None and 0 < v <= 100.5]
            if len(nums) < 3:
                continue
            s = sum(nums)
            if abs(s - 100.0) <= 0.6:
                continue
            if not (90 <= s <= 110):
                continue
            out.append({
                "code": "PCT_TOTAL",
                "severity": "warning",
                "subject": "%s column %s" % (sheet, c),
                "message": "the %d shares in column %s sum to %.1f%%, not 100%%" % (len(nums), c, s),
                "where": ["%s!%s%d:%s%d" % (sheet, c, vals[0][0], c, vals[-1][0])],
                "files": [path],
                "context": [", ".join("%g" % v for v in nums)],
            })
    return out


# ---------------------------------------------------------------- gate


def audit(paths):
    numbers, units_all, metas, findings = [], [], [], []
    grids_by_file = {}
    for p in paths:
        ext = os.path.splitext(p)[1].lower()
        if ext == ".pptx":
            u, m = pptx_units(p)
        elif ext == ".docx":
            u, m = docx_units(p)
        elif ext == ".xlsx":
            u, m, grids = xlsx_units(p)
            grids_by_file[p] = grids
            findings.extend(table_checks(grids, p))
            findings.extend(pct_checks(grids, p))
        else:
            with open(p, encoding="utf-8", errors="replace") as f:
                text = f.read()
            u, m = [], None
            for i, line in enumerate(text.splitlines(), 1):
                if line.strip():
                    u.append({"file": p, "where": "line %d" % i, "text": line.strip()})
        metas.append((p, m or {}))
        units_all.extend(u)
        numbers.extend(collect_numbers(u))

    findings.extend(conflicts(numbers))

    # placeholders and figure references
    for u in units_all:
        m = PLACEHOLDERS.search(u["text"])
        if m:
            findings.append({"code": "PLACEHOLDER", "severity": "blocker", "subject": m.group(0),
                             "message": "unfinished placeholder \"%s\" in the deliverable" % m.group(0),
                             "where": [u["where"]], "files": [u["file"]], "context": [u["text"][:150]]})

    figure_count = sum(m.get("figures", 0) for _, m in metas)
    table_count = sum(m.get("tables", 0) for _, m in metas)
    max_ref = {"figure": 0, "table": 0}
    for u in units_all:
        for m in FIG_REF.finditer(u["text"]):
            kind = "figure" if m.group(1).lower().startswith("fig") else "table"
            max_ref[kind] = max(max_ref[kind], int(m.group(2)))
    if max_ref["figure"] and max_ref["figure"] > max(1, figure_count):
        findings.append({"code": "FIGURE_REF", "severity": "warning", "subject": "Figure %d" % max_ref["figure"],
                         "message": "text refers to Figure %d but the set contains %d figure(s)"
                                    % (max_ref["figure"], figure_count),
                         "where": ["cross-file"], "files": [], "context": [""]})
    if max_ref["table"] and max_ref["table"] > max(1, table_count):
        findings.append({"code": "TABLE_REF", "severity": "warning", "subject": "Table %d" % max_ref["table"],
                         "message": "text refers to Table %d but the set contains %d table(s)"
                                    % (max_ref["table"], table_count),
                         "where": ["cross-file"], "files": [], "context": [""]})

    # duplicate headings in the same document
    for p, m in metas:
        heads = [h for h in (m.get("headings") or []) if h.strip()]
        seen = {}
        for h in heads:
            seen[h] = seen.get(h, 0) + 1
        for h, c in seen.items():
            if c > 1:
                findings.append({"code": "DUPLICATE_HEADING", "severity": "note", "subject": h,
                                 "message": "heading \"%s\" appears %d times" % (h, c),
                                 "where": [p], "files": [p], "context": [""]})

    return {
        "files": [p for p in paths],
        "counts": {"numbers": len(numbers), "figures": figure_count, "tables": table_count},
        "findings": findings,
    }


def rank(f):
    return {"blocker": 0, "warning": 1, "note": 2}.get(f["severity"], 3)


def render(payload):
    fs = sorted(payload["findings"], key=lambda f: (rank(f), f["code"]))
    blockers = sum(1 for f in fs if f["severity"] == "blocker")
    verdict = "BLOCK" if blockers else ("WARN" if fs else "PASS")
    out = ["# Consistency gate\n"]
    out.append("Files: %d · numbers parsed: %d · figures: %d · tables: %d\n"
               % (len(payload["files"]), payload["counts"]["numbers"],
                  payload["counts"]["figures"], payload["counts"]["tables"]))
    out.append("**Verdict: %s** — blockers: %d · warnings: %d · notes: %d\n"
               % (verdict, blockers, sum(1 for f in fs if f["severity"] == "warning"),
                  sum(1 for f in fs if f["severity"] == "note")))
    if not fs:
        out.append("No conflict above tolerance. Numbers, totals and references agree across the set.\n")
        return "\n".join(out)

    out.append("## Findings\n")
    out.append("| # | Severity | Code | Where | What is wrong |")
    out.append("|---|---|---|---|---|")
    for i, f in enumerate(fs, 1):
        out.append("| %d | %s | %s | %s | %s |" % (i, f["severity"], f["code"],
                                                    " · ".join(str(w) for w in f["where"]),
                                                    f["message"].replace("|", "/")))
    out.append("\n## Evidence\n")
    for i, f in enumerate(fs, 1):
        out.append("**%d. %s** — %s" % (i, f["code"], f["message"]))
        for w, c in zip(f["where"], f["context"]):
            out.append("- `%s` — %s" % (w, (c or "").strip()[:200]))
        out.append("")
    out.append("## Rule set\n")
    out.append("- NUMBER_CONFLICT: the same label carries two values whose relative difference exceeds the rounding tolerance.")
    out.append("- TOTAL_MISMATCH: a row labelled total does not equal the contiguous rows above it in the same column.")
    out.append("- PCT_TOTAL: a set of shares in one row misses 100%.")
    out.append("- FIGURE_REF / TABLE_REF: text cites an exhibit number the set does not contain.")
    out.append("- PLACEHOLDER: TODO/TBD/lorem survives into the deliverable.")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description="Find where a deck, memo and workbook disagree with each other.")
    ap.add_argument("--input", action="append", required=True, help="file or directory (repeatable)")
    ap.add_argument("--out", default="out")
    a = ap.parse_args()

    paths = []
    for i in a.input:
        if os.path.isdir(i):
            for root, _, names in os.walk(i):
                for n in names:
                    if n.lower().endswith((".pptx", ".docx", ".xlsx", ".md", ".txt")) and not n.startswith("~$"):
                        paths.append(os.path.join(root, n))
        elif os.path.exists(i):
            paths.append(i)
        else:
            print("missing input: %s" % i, file=sys.stderr)
    if not paths:
        raise SystemExit("no inputs")

    payload = audit(sorted(paths))
    os.makedirs(a.out, exist_ok=True)
    open(os.path.join(a.out, "consistency_gate.md"), "w", encoding="utf-8").write(render(payload))
    open(os.path.join(a.out, "consistency_gate.json"), "w", encoding="utf-8").write(
        json.dumps(payload, ensure_ascii=False, indent=2))
    print("files=%d numbers=%d findings=%d" % (len(payload["files"]), payload["counts"]["numbers"], len(payload["findings"])))
    for f in sorted(payload["findings"], key=rank)[:10]:
        print("  [%s] %s %s" % (f["severity"], f["code"], f["where"]))
    print("wrote %s" % os.path.join(a.out, "consistency_gate.md"))
    return 0


if __name__ == "__main__":
    sys.exit(main())



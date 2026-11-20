"""Build the NSDUH slice's metadata records from SAMHSA's codebooks, the data files and the crosswalk.

Usage (under the memory cap):
    python build/build_records.py

Reads build/slice.yaml; for each release, its codebook as text (data/interim/<codebook>.txt,
pdftotext -layout of the PDF in data/raw/) and its Stata file (data/interim/<dta>, unzipped
from data/raw/); and SAMHSA's comparability chart, data/raw/PUFComparability_2025.xlsx.
Writes metadata/releases/<release>/variables.jsonl, metadata/groups.jsonl,
metadata/links.jsonl and build/report.md.

Where each field comes from (STANDARD.md §3):
- label, question text and item numbers, codes and labels, valid ranges, notes: the codebook
  (publisher), with its page
- the name's spelling, which codes occur, counts for availability, storage type: the data file
- measure: build/slice.yaml (hand, Claude, Task 60)
- missing types: rule nsduh_missing_from_labels (the codebook marks no codes as missing;
  SAMHSA writes its reserved codes' labels in capitals, e.g. "94 = DON'T KNOW")
- links between releases and module groups: the comparability chart (publisher)
- storage type: rule storage_type_from_values
- scale signatures: the shared library's rule (pdt.derived)

Names are spelled as each release's data file spells them: the 2024 file writes most names in
lowercase, while its codebook and the other years use capitals.
"""

import collections
import json
import re
from pathlib import Path

import openpyxl
import pandas as pd
import pyreadstat
import yaml

# The shared library (PublicDataToolkit lib/pdt): installed, or on PYTHONPATH; otherwise a
# checkout next to this repo, or $PDT.
try:
    from pdt.derived import scale_signatures
except ImportError:
    import os
    import sys
    sys.path.insert(0, str(Path(os.environ.get("PDT", Path(__file__).resolve().parents[2] / "PublicDataToolkit"))
                           / "lib"))
    from pdt.derived import scale_signatures

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
INTERIM = ROOT / "data" / "interim"
META = ROOT / "metadata"
DATASET = "nsduh"
CROSSWALK = "nsduh-puf-comparability-2025"
RULE_VERSION = "0.1"

# ---------------------------------------------------------------- the codebook

ENTRY = re.compile(r"^(?P<token>[A-Za-z][A-Za-z0-9_]*)\s+Len\s*:\s*(?P<len>\d+)\s+(?P<label>\S.*?)\s*$")
REFS = re.compile(r"^\((?P<refs>[^()]*)\)\s*$")
VALUE = re.compile(r"^\s+(?P<code>RANGE|\.|-?\d+(?:\.\d+)?)\s*=\s*(?P<label>.*?)\s*\.{2,}\s*(?P<freq>\d+)\s+(?P<pct>[\d.]+)\s*$")
VALUE_START = re.compile(r"^\s{10,}(?P<code>RANGE|\.|-?\d+(?:\.\d+)?)\s*=\s*(?P<label>.*?)\s*$")
FOOTER = re.compile(r"^\s*Codebook Creation Date:.*?(?P<page>\d+)\s*$")


def clean_lines(text):
    """The codebook's lines without page footers and the running header after each one.

    Returns [(line, page)], page being the number printed in the footer that ends the page.
    """
    raw = text.split("\n")
    out, pending, i = [], [], 0
    while i < len(raw):
        line = raw[i].replace("\f", "")
        m = FOOTER.match(line)
        if m:
            page = int(m.group("page"))
            out += [(l, page) for l in pending]
            pending = []
            # The running header: the next non-blank line, indented far, in capitals.
            j = i + 1
            while j < len(raw) and not raw[j].strip():
                j += 1
            if j < len(raw) and raw[j].startswith(" " * 30) and raw[j].strip().isupper():
                i = j
            i += 1
            continue
        pending.append(line)
        i += 1
    out += [(l, None) for l in pending]
    return out


def paragraph_before(lines, k):
    """The contiguous non-blank lines ending just before index k (skipping blanks), and its start."""
    j = k - 1
    while j >= 0 and not lines[j][0].strip():
        j -= 1
    end = j
    while j >= 0 and lines[j][0].strip():
        j -= 1
    return [l for l, _ in lines[j + 1:end + 1]], j + 1


def join(lines):
    return " ".join(" ".join(l.split()) for l in lines).strip()


def parse_codebook(text, columns):
    """{NAME (upper): entry} for every entry in the codebook.

    columns: the data file's names, by upper case, so a trailing note marker ("DSTHOP301":
    the name DSTHOP30, marked "1" by a section note) can be told from a name ending in 1.
    """
    lines = clean_lines(text)
    entries = {}
    for k, (line, page) in enumerate(lines):
        m = ENTRY.match(line)
        if not m:
            continue
        token, mark = m.group("token"), None
        if token.upper() not in columns and token[-1].isdigit() and token[:-1].upper() in columns:
            token, mark = token[:-1], token[-1]
        name = token.upper()
        if name in entries:
            continue
        entry = {"name": name, "label": " ".join(m.group("label").split()), "page": page, "mark": mark}
        # Item numbers or source variables, in parentheses just above the entry.
        para, start = paragraph_before(lines, k)
        if len(para) == 1 and REFS.match(para[0].strip()):
            refs = [r.strip() for r in REFS.match(para[0].strip()).group("refs").split(",") if r.strip()]
            entry["refs"] = refs
            para, start = paragraph_before(lines, start)
        else:
            refs = []
        # A recode lists its source variables; a question lists its questionnaire items.
        entry["recode"] = bool(refs) and all(r.upper() in columns for r in refs)
        if para and not entry["recode"] and not para[0].strip().startswith("NOTE"):
            entry["question"] = join(para)
            para, start = paragraph_before(lines, start)
        if para and para[0].strip().startswith("NOTE"):
            entry["note"] = join(para)
        if mark:
            # The section note that says "marked by a<mark>".
            for back in range(k - 1, max(k - 400, 0), -1):
                if f"marked by a{mark}" in lines[back][0]:
                    note_para, _ = paragraph_before(lines, back + 1)
                    entry["section_note"] = join(note_para)
                    break
        entry["values"] = parse_values(lines, k + 1)
        if entry.get("page") is None:
            entry["page"] = next((p for _, p in lines[k:k + 200] if p), None)
        entries[name] = entry
    return entries


def parse_values(lines, k):
    """The value lines after an entry, until the next paragraph."""
    values, pending, blanks = [], None, 0
    while k < len(lines):
        line = lines[k][0]
        if not line.strip():
            blanks += 1
            if blanks >= 2 and values:
                break
            k += 1
            continue
        blanks = 0
        if "Freq" in line and "Pct" in line:
            k += 1
            continue
        m = VALUE.match(line if pending is None else pending + " " + line.strip())
        if m:
            values.append({"code": m.group("code"), "label": " ".join(m.group("label").split()),
                           "freq": int(m.group("freq"))})
            pending = None
        elif VALUE_START.match(line) and pending is None:
            pending = line.rstrip()          # a label that wraps onto the next line
        elif pending is not None:
            pending = pending + " " + line.strip()
        else:
            break
        k += 1
    return values


# ---------------------------------------------------------------- rules

# Rule nsduh_missing_from_labels: SAMHSA's reserved codes have labels in capitals.
RESERVED = [
    (re.compile(r"^BAD DATA"), ["uncodeable"]),
    (re.compile(r"^DON'?T KNOW|^DON T KNOW"), ["dk"]),
    (re.compile(r"^REFUSED"), ["refused"]),
    (re.compile(r"^BLANK"), ["no_answer"]),
    (re.compile(r"^LEGITIMATE SKIP"), ["inap"]),
    (re.compile(r"^NEVER USED"), ["inap"]),
    (re.compile(r"^DID NOT USE"), ["inap"]),
]
RULE_MISSING = {"origin": "rule", "rule": "nsduh_missing_from_labels", "rule_version": RULE_VERSION,
                "note": "The codebook marks no code as missing. SAMHSA writes the labels of its reserved codes "
                        "in capitals (85 BAD DATA, 94 DON'T KNOW, 97 REFUSED, 98 BLANK, 99 LEGITIMATE SKIP, "
                        "91 NEVER USED, 93 DID NOT USE ...); the rule types those, and '.' labeled "
                        "'Aged 12-17' (out of the universe)."}


def classify(code, label):
    """(kind, missing_types or None, note or None) for one code."""
    head = re.sub(r"\s*\(.*$", "", label).strip()
    if code == ".":
        if head.lower() in {"aged 12-17", "aged 12 to 17"}:
            return "missing", ["inap"], None
        return "unresolved", None, f"'.' labeled {label!r}: blank in the data; the label doesn't say why for every case"
    # The patterns are in capitals and case-sensitive, so a recode's "Never used" stays valid,
    # while "BAD DATA Logically assigned" (capitals, then lowercase) is a reserved code.
    for pat, types in RESERVED:
        if pat.search(head):
            return "missing", types, None
    if head and head == head.upper() and re.search(r"[A-Z]{3}", head):
        return "unresolved", None, f"label in capitals, like a reserved code, but the rule doesn't know {head!r}"
    return "valid", None, None


def code_value(code):
    if code == ".":
        return "."
    x = float(code)
    return int(x) if x.is_integer() else x


# ---------------------------------------------------------------- the data

def read_data(path, wanted):
    """(names by upper case -> as spelled, {NAME: Counter of codes}, {NAME: storage type}) for wanted names."""
    _, meta = pyreadstat.read_dta(str(path), metadataonly=True)
    spelled = {c.upper(): c for c in meta.column_names}
    use = [spelled[n] for n in wanted if n in spelled]
    frame, _ = pyreadstat.read_dta(str(path), usecols=use)
    counts, storage = {}, {}
    for col in use:
        s = frame[col]
        vc = s.value_counts(dropna=False)
        counts[col.upper()] = collections.Counter({("." if pd.isna(k) else code_value(k)): int(v)
                                                   for k, v in vc.items()})
        present = s.dropna()
        num = pd.to_numeric(present, errors="coerce")
        storage[col.upper()] = ("string" if num.isna().any() else
                                "integer" if len(num) == 0 or (num % 1 == 0).all() else "float")
    return spelled, counts, storage, meta.number_rows


# ---------------------------------------------------------------- build

def build_record(name, release, rel, entry, counts, storage, spelled, measure, spec):
    wave = int(release)
    rec = {"dataset": DATASET, "release": release, "name": spelled[name]}
    cb = rel["codebook_source"]
    loc = {"origin": "publisher", "source": cb, "location": f"p. {entry['page']}"}
    rec["label"] = entry["label"]
    if entry.get("question"):
        q = {"text": entry["question"]}
        if entry.get("refs"):
            q["number"] = ", ".join(entry["refs"])
        rec["questions"] = [q]
    rec["storage_type"] = storage[name]
    if measure:
        rec["measure"] = measure
    values, report, rng = [], [], None
    for v in entry["values"]:
        if v["code"] == "RANGE":
            m = re.match(r"(-?[\d.]+)\s*-\s*(-?[\d.]+)", v["label"])
            if m:
                lo, hi = (code_value(x) for x in m.groups())
                rng = {"min": lo, "max": hi}
            continue
        kind, types, note = classify(v["code"], v["label"])
        item = {"code": code_value(v["code"]), "label": v["label"], "kind": kind}
        if types:
            item["missing_types"] = types
        if kind == "unresolved":
            item["origin"] = {**RULE_MISSING, "note": note}
            report.append(("unresolved", f"{release} `{rec['name']}` code {v['code']}: {note}"))
        values.append(item)
    if values:
        rec["values"] = values
    if rng:
        rec["valid_range"] = rng
    notes = [x for x in (entry.get("note"), entry.get("section_note")) if x]
    if notes:
        rec["notes"] = " ".join(notes)
    missing_codes = {v["code"] for v in values if v["kind"] != "valid"}
    n = sum(c for code, c in counts[name].items() if code not in missing_codes and code != ".")
    rec["availability"] = [{"wave": wave, "n": n}] if n else []
    rec["source_location"] = {"source": cb, "page": entry["page"]}
    fields = {"storage_type": {"origin": "rule", "rule": "storage_type_from_values", "rule_version": RULE_VERSION,
                               "source": rel["data_source"]},
              "availability": {"origin": "data_file", "source": rel["data_source"],
                               "note": "n counts the cases with a code that isn't missing or unresolved."}}
    if measure:
        fields["measure"] = {"origin": "hand", "author": spec["author"], "date": spec["date"]}
    if entry.get("recode"):
        rec.setdefault("extensions", {})["nsduh"] = {"source_variables": entry["refs"]}
    record_origin = {**loc, "note": "The name is spelled as in this release's data file"
                     + (" (the codebook writes it in capitals)." if rec["name"] != name else ".")}
    rec["provenance"] = {"record": record_origin, "fields": fields}
    if any(v["kind"] != "valid" for v in values):
        rec["provenance"]["codes"] = RULE_MISSING
    # C4: codes in the data that the codebook doesn't list, or outside its range. An entry that
    # lists no values (an identifier, a weight) is skipped.
    listed = {v["code"] for v in values}
    if entry["values"]:
        extra = [(code, c) for code, c in sorted(counts[name].items(), key=lambda kv: str(kv[0]))
                 if code not in listed and not (rng and code != "." and rng["min"] <= code <= rng["max"])]
        if extra:
            report.append(("undocumented", f"{release} `{rec['name']}`: {len(extra)} codes in the data that the "
                                           f"codebook doesn't list, e.g. {extra[:5]}"))
    # C5: the codebook's frequencies against the data's.
    for v in entry["values"]:
        if v["code"] == "RANGE":
            continue
        got = counts[name].get(code_value(v["code"]), 0)
        if got != v["freq"]:
            report.append(("counts", f"{release} `{rec['name']}` code {v['code']}: codebook {v['freq']}, data {got}"))
    return rec, report


def load_crosswalk():
    wb = openpyxl.load_workbook(RAW / "PUFComparability_2025.xlsx", read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    head_i = next(i for i, r in enumerate(rows) if r and r[1] == "VARIABLE")
    head = rows[head_i]
    years = {int("20" + str(h).split()[-1]): j for j, h in enumerate(head) if h and str(h).startswith("PUF")}
    out = {}
    for r in rows[head_i + 1:]:
        if r[1]:
            out[str(r[1]).upper()] = {"row": r[0], "gen": {y: r[j] for y, j in years.items()},
                                      "comparable": r[head.index("Comparable Variable")],
                                      "module": r[head.index("MODULE")], "notes": r[head.index("NOTES")]}
    return out


def links_and_groups(records_by_release, crosswalk, releases):
    links, report = [], []
    order = sorted(releases)
    have = {rel: {r["name"].upper(): r["name"] for r in recs} for rel, recs in records_by_release.items()}
    for a, b in zip(order, order[1:]):
        for name in sorted(set(have[a]) & set(have[b])):
            cw = crosswalk.get(name)
            if not cw:
                report.append(("links", f"`{name}` isn't in the comparability chart"))
                continue
            ga, gb = cw["gen"].get(int(a)), cw["gen"].get(int(b))
            if ga is not None and ga == gb:
                note = "Same comparability number in the chart for both years."
                if have[a][name] != have[b][name]:
                    note += f" The data files spell the name {have[a][name]} and {have[b][name]}."
                links.append({"from": {"release": a, "name": have[a][name]}, "to": {"release": b, "name": have[b][name]},
                              "relations": ["same_item"],
                              "origin": {"origin": "publisher", "source": CROSSWALK, "location": f"row {cw['row']}"},
                              "note": note})
            else:
                report.append(("links", f"`{name}` {a}->{b}: comparability number {ga} -> {gb}, so no link"))
        # Renames: the chart's "Comparable Variable" names a predecessor.
        for name in sorted(set(have[b]) - set(have[a])):
            cw = crosswalk.get(name)
            prev = (cw or {}).get("comparable")
            if prev and str(prev).upper() in have[a]:
                links.append({"from": {"release": a, "name": have[a][str(prev).upper()]},
                              "to": {"release": b, "name": have[b][name]},
                              "relations": ["renamed"],
                              "origin": {"origin": "publisher", "source": CROSSWALK, "location": f"row {cw['row']}"},
                              "note": "The chart names it the comparable variable. Its note: "
                                      + " ".join(str(cw["notes"] or "").split())})
    groups = collections.defaultdict(list)
    for rel, recs in records_by_release.items():
        for r in recs:
            mod = (crosswalk.get(r["name"].upper()) or {}).get("module")
            if mod:
                groups[mod].append({"release": rel, "name": r["name"]})
    group_lines = []
    for mod in sorted(groups):
        slug = re.sub(r"[^a-z0-9]+", "-", mod.lower()).strip("-")
        group_lines.append({"id": f"module-{slug}", "kind": "section", "label": mod.title(),
                            "members": sorted(groups[mod], key=lambda m: (m["release"], m["name"])),
                            "ordered": False,
                            "origin": {"origin": "publisher", "source": CROSSWALK,
                                       "note": "The chart's MODULE column; members limited to this slice."}})
    return links, group_lines, report


def main():
    spec = yaml.safe_load((ROOT / "build" / "slice.yaml").read_text())
    wanted = [n.upper() for n in spec["variables"]]
    measures = {n.upper(): m for n, m in spec["variables"].items()}
    crosswalk = load_crosswalk()
    records_by_release, report, coverage = {}, [], []
    for release, rel in spec["releases"].items():
        spelled, counts, storage, nrows = read_data(INTERIM / rel["dta"], wanted)
        text = (INTERIM / (Path(rel["codebook"]).stem + ".txt")).read_text()
        entries = parse_codebook(text, set(spelled))
        undocumented = sorted(set(spelled) - set(entries))
        not_in_data = sorted(set(entries) - set(spelled))
        coverage.append(f"- {release}: {len(spelled)} variables in the data, {len(entries)} codebook entries; "
                        f"in the data only: {len(undocumented)} {undocumented[:8]}; "
                        f"in the codebook only: {len(not_in_data)} {not_in_data[:8]}")
        recs = []
        for name in wanted:
            if name not in spelled:
                continue
            if name not in entries:
                report.append(("undocumented", f"{release} `{spelled[name]}`: in the data, no codebook entry"))
                continue
            rec, rep = build_record(name, release, rel, entries[name], counts, storage, spelled,
                                    measures.get(name), spec)
            recs.append(rec)
            report += rep
        records_by_release[release] = recs
        write_jsonl(META / "releases" / release / "variables.jsonl", recs)
    links, groups, rep = links_and_groups(records_by_release, crosswalk, list(spec["releases"]))
    report += rep
    write_jsonl(META / "links.jsonl", links)
    # Scale signatures by the shared rule, one per record (names differ in case across releases).
    write_jsonl(META / "derived" / "scale_signature.jsonl",
                scale_signatures([r for recs in records_by_release.values() for r in recs], per_release=True))
    write_jsonl(META / "groups.jsonl", groups)
    write_report(records_by_release, links, groups, coverage, report)
    print(", ".join(f"{r}: {len(v)}" for r, v in records_by_release.items()),
          f"records; {len(links)} links; {len(groups)} groups; see build/report.md")


def write_jsonl(path, records):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def write_report(records_by_release, links, groups, coverage, report):
    by = collections.defaultdict(list)
    for kind, line in report:
        by[kind].append(line)
    lines = ["# Build report: nsduh", "", "Generated by `build/build_records.py`. Do not edit.", ""]
    for rel, recs in records_by_release.items():
        lines.append(f"- {rel}: {len(recs)} records, {sum('questions' in r for r in recs)} with question text")
    lines += [f"- Links: {len(links)}; module groups: {len(groups)}", "", "## C3: coverage", "", *coverage, ""]
    for kind, title in [("undocumented", "C4: codes or variables in the data that the codebook doesn't list"),
                        ("unresolved", "Codes left unresolved"),
                        ("counts", "C5: codebook frequencies vs the data"),
                        ("links", "Links not written")]:
        lines += [f"## {title}", "", *([f"- {x}" for x in by[kind]] or ["None."]), ""]
    (ROOT / "build" / "report.md").write_text("\n".join(lines))


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Write the NSDUH data in the common format (STANDARD.md §8).

Reads each year's Stata file, unzipped into data/interim/ (build/slice.yaml gives its name),
and writes data/processed/nsduh-<year>.parquet with the variables this repo has records for.
pdt.release.write_release reads the columns in batches, and checks each batch, then the whole
file, by a round trip against the cells read. The files have about 57,000 rows and 2,600
columns; only the records' columns are read. Run under the memory cap.

Stata missing values become Codes, as pdt.readers.read_stata gives them. Names are matched
case-insensitively: the 2024 file writes most names in lowercase.
"""

import datetime
import sys
from pathlib import Path

import pandas as pd
import pyreadstat
import yaml

try:
    from pdt.data import Code
    from pdt.release import data_source, write_release
except ImportError:
    sys.exit("Needs the PublicDataToolkit library: pip install -e ../PublicDataToolkit, "
             "or run through `make data`.")

ROOT = Path(__file__).resolve().parent.parent


def read_columns(path, names, encoding=None):
    """The named columns of a Stata file as a raw frame, as pdt.readers.read_stata gives them:
    extended missing .a-.z become Code(".a"), system missing Code("."), integers stay ints.
    Dates (UpdateDate) become ISO text, 2020-05-19: the common format has no date type."""
    kwargs = {"encoding": encoding} if encoding else {}
    _, meta = pyreadstat.read_dta(str(path), metadataonly=True, **kwargs)
    actual = {c.upper(): c for c in meta.column_names}
    if missing := [n for n in names if n.upper() not in actual]:
        raise KeyError(f"{path}: no column(s) {missing}")
    cols = [actual[n.upper()] for n in names]
    frame, meta = pyreadstat.read_dta(str(path), usecols=cols, user_missing=True, **kwargs)
    types = meta.readstat_variable_types
    out = {}
    for name, col in zip(names, cols):
        kind = types[col]
        numeric = kind in ("int8", "int16", "int32", "float", "double")
        integer = kind.startswith("int")
        cells = []
        for c in frame[col].tolist():
            if numeric and isinstance(c, str):
                c = Code("." + c)
            elif numeric and isinstance(c, float) and c != c:
                c = Code(".")
            elif integer and isinstance(c, float):
                c = int(c)
            elif isinstance(c, (datetime.date, pd.Timestamp)):
                c = c.isoformat()                      # dates as ISO text: the common format has no date type
            cells.append(c)
        out[name] = pd.Series(cells, dtype=object)
    return pd.DataFrame(out)


def read(release, table, names):
    spec = yaml.safe_load((ROOT / "build" / "slice.yaml").read_text())["releases"][release]
    path = ROOT / "data" / "interim" / spec["dta"]
    if not path.exists():
        sys.exit(f"{path} is missing: unzip data/raw/{spec['stata']} into data/interim/")
    return read_columns(path, names), data_source(ROOT, release, table)["id"]


if __name__ == "__main__":
    write_release(ROOT, read)

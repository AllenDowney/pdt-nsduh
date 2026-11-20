# Extensions

Every key under `extensions` in this repo's records, with a reason (STANDARD.md §4.4). `pdt check` fails a key that isn't listed here (C10).

| Key | What it holds | Source | Why it's an extension |
|---|---|---|---|
| `nsduh.source_variables` | For a recoded variable, the variables the codebook lists above its entry (e.g. `SPDPSTMON` from `KSSLR6MON`) | the codebook | The standard has no field for the inputs of a recode; `derived_from` links are for continuity across releases. A candidate for the standard if another dataset needs it |

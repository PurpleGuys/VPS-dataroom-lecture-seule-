# How to add a document to the dataroom

Two files, always together: the PDF and its companion note.

1. **Drop the PDF** in the right folder (`01_Financial`, `02_Permits_Regulatory`, ...).
   Keep the filename short, lowercase, with the year in it:
   `veolia-2024-universal-registration-document.pdf`.

2. **Create the companion note** beside it, named `<the exact pdf filename>.md`
   (so `veolia-2024-urd.pdf` → `veolia-2024-urd.pdf.md`), with this frontmatter:

```yaml
---
title: Veolia Universal Registration Document 2024
source_url: https://www.veolia.com/en/investors/...
date_consulted: 2026-09-22
publisher: Veolia Environnement SA
doc_type: annual_report      # annual_report | 10-K | permit | regulation | press_release | rating | other
target: Veolia
folder: 01_Financial
language: en                 # en | fr
added_by: ethan
ocr: false                   # true only for a scanned PDF with no text layer
---
```

3. **Commit and push.** The server pulls every five minutes and re-indexes what changed.

## The three rules the server enforces

- No companion note → the document is **not indexed** and appears in `dataroom_status()`
  under errors.
- No `source_url` or no `date_consulted` → same refusal. A document we cannot cite is
  worse than a document we do not have.
- `date_consulted` must be `YYYY-MM-DD`, and `source_url` must start with `http`.

## register.csv

Every figure that ends up in the report goes in `register.csv`, one row each:

`id,label,value,unit,period,file,page,source_url,date_consulted,added_by,checked_by`

`checked_by` must be a **different person** from `added_by`. `lookup_figure()` reads
this file, so a figure in the register is a figure anyone in the group can re-verify
in one step with `read_page(file, page)`.

# Form example

`form-template.docx` is a generic form with two separate `[...]` fields beside the Name and Summary labels, a two-row Item / Goal / Metric table, native Standard and Expedited checkboxes, and read-only guidance. `form-filled.docx` fills those fields and checks Standard without changing the guidance or labels.

From the repository root, regenerate the artifacts with:

```sh
PYTHONPATH=src .venv/bin/python examples/form/generate.py
PYTHONPATH=src .venv/bin/python examples/form/populate.py
```

`populate.py` extracts a temporary bundle, writes `field-map.json` and `patch.json`, renders the filled DOCX, and checks package/text/style fidelity. The map uses distinct paragraph IDs for the repeated markers and keyed table rows; it is reviewed for **this template only**, not a general discovery rule. The patch binds to the generated template hash. ZIP timestamps are normalized so repeated runs produce the same artifacts. Visual layout comparison is not part of the fidelity check.

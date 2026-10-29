# DocXUnravel

DocXUnravel reads DOCX files, extracts their text and structure, and applies checked edits without rebuilding the document from scratch. It is a Python library with a CLI. Your application or agent decides what a field means; DocXUnravel validates the patch and keeps untouched parts of the DOCX intact.

Requires Python 3.11 or newer.

## Install

```sh
python -m pip install docxunravel
```

The PyPI package contains the library and CLI. Download the sample files separately from the GitHub Release.

## Try the form example

The [form example](https://github.com/Vinayak-G87/DocXUnravel/blob/main/examples/form/README.md) has two `[...]` fields, a small table, and two checkboxes. Its patch is bound to the included template. The four example files are distributed as a [GitHub Release ZIP](https://github.com/Vinayak-G87/DocXUnravel/releases/download/v0.4.2/docxunravel-examples.zip), not in the PyPI package. The ZIP will be available after the `v0.4.2` release workflow finishes.

Run these steps in order, using the Python environment where you installed DocXUnravel:

1. Download and unzip the example into a new working directory. No repository clone is needed.

   ```sh
   mkdir docxunravel-example
   cd docxunravel-example
   curl -fL -o docxunravel-examples.zip https://github.com/Vinayak-G87/DocXUnravel/releases/download/v0.4.2/docxunravel-examples.zip
   python -m zipfile -e docxunravel-examples.zip .
   ```

2. Extract the template. This creates `form-bundle/profile.json`, which the next commands need. Use a new bundle directory if `form-bundle` already contains files.

   ```sh
   python -m docxunravel extract examples/form/form-template.docx form-bundle --explicit-placeholders-only
   ```

3. Preview the patch. The first argument is the bundle directory; the second is the patch file. This step does not write a DOCX.

   ```sh
   python -m docxunravel preview form-bundle examples/form/patch.json
   ```

4. Render the filled document as `filled.docx` in your working directory.

   ```sh
   python -m docxunravel render-bundle form-bundle examples/form/patch.json filled.docx
   ```

Compare [the template](https://github.com/Vinayak-G87/DocXUnravel/blob/main/examples/form/form-template.docx) and [the filled example](https://github.com/Vinayak-G87/DocXUnravel/blob/main/examples/form/form-filled.docx), or inspect the [field map](https://github.com/Vinayak-G87/DocXUnravel/blob/main/examples/form/field-map.json). The field map belongs to this form, not to other DOCX files.

## Work with your own DOCX

Choose any name for the extracted directory. This guide uses `my-bundle` throughout:

```sh
python -m docxunravel extract my-document.docx my-bundle --explicit-placeholders-only
```

Read `my-bundle/content.json` to find paragraph IDs and their current text, including paragraphs in table cells. Use nearby labels to decide which ID to edit; repeated `[...]` markers may mean different things. Read `my-bundle/profile.json` for table structure, styles, formatting, edit warnings, and the `patch_contract` that binds edits to this exact document. `my-bundle/package/` holds the original DOCX parts. These are **inputs to inspect**, not files to edit: changing `profile.json` makes its binding invalid, and editing `content.json` does not update the DOCX.

Create a separate JSON file for your proposed changes. You choose its filename too; this guide uses `my-patch.json`. This example asks you to pick one real paragraph ID from `content.json` and copies its current text into `expected`:

```python
import json
from pathlib import Path

bundle = Path("my-bundle")
content = json.loads((bundle / "content.json").read_text(encoding="utf-8"))
profile = json.loads((bundle / "profile.json").read_text(encoding="utf-8"))
element_id = "body.t0.r0.c1.p0"  # replace with an ID from your content.json
paragraph = next(p for p in content["paragraphs"] if p["id"] == element_id)
patch = {
    "contract": profile["patch_contract"],
    "elements": {
        element_id: {
            "expected": paragraph["text"],
            "value": "Alex Morgan",
            "format_policy": "reject_mixed_runs",
        }
    },
}
Path("my-patch.json").write_text(json.dumps(patch, indent=2), encoding="utf-8")
```

Run the Python snippet from the directory containing `my-bundle`. The ID is only an example; it must match your document. `expected` catches stale targets, and `reject_mixed_runs` refuses full replacement of paragraphs with mixed run formatting rather than silently flattening their emphasis. For other supported patches, `tokens` maps approved placeholder names to strings, `runs` targets individual runs, and `operations` handles structures such as table rows and checkboxes. See [the complete form patch](https://github.com/Vinayak-G87/DocXUnravel/blob/main/examples/form/patch.json) for `operations`. The allowed top-level keys are `contract`, `elements`, `tokens`, `runs`, `operations`, and `options`; `preview` validates the patch before writing a DOCX. The example's prebuilt `patch.json` works only with its included template, not your own document.

```sh
python -m docxunravel preview my-bundle my-patch.json
python -m docxunravel render-bundle my-bundle my-patch.json my-output.docx
python -m docxunravel fidelity-report my-document.docx my-output.docx fidelity.json \
  --bundle my-bundle --patch my-patch.json
```

Preview does not write a DOCX. Rendering writes a new file and checks the resulting package. Fidelity checks expected text and preserved package parts, styles, and numbering, but not page appearance. The output keeps the source's section settings; longer or shorter text can still change line breaks and pagination. For layout-sensitive documents, compare the files visually in Word or use the optional `visual` command with LibreOffice.

### Placeholders

Use `--placeholder NAME=LITERAL` during `extract` if you want to replace exact text by name:

```sh
python -m docxunravel extract my-document.docx my-bundle \
  --explicit-placeholders-only \
  --placeholder 'OWNER=[Enter owner]'
```

The profile records where `OWNER` occurs. A patch can then set `"tokens": {"OWNER": "Alex Morgan"}`. The same literal is replaced everywhere it appears, including in table cells. For repeated `[...]` fields with different meanings, use their paragraph IDs instead. `--explicit-placeholders-only` leaves unapproved `{{NAME}}` text alone. Without it, automatic brace-token detection remains available for existing templates. DOCX text is never evaluated as Python code.

### Python callers and agents

Callers can inspect the profile and choose targets with their own Python or agent logic. The snippet below assumes the release example ZIP has been extracted into the current directory and uses a different bundle name from the CLI example above:

```python
import json
from pathlib import Path
from docxunravel import export_bundle, preview_bundle, render_bundle

profile = export_bundle("examples/form/form-template.docx", "python-bundle",
                        auto_placeholders=False)
patch = json.loads(Path("examples/form/patch.json").read_text(encoding="utf-8"))
assert patch["contract"] == profile["patch_contract"]
plan = preview_bundle("python-bundle", patch)
if plan["valid"]:
    render_bundle("python-bundle", "filled.docx", patch)
```

For a different DOCX, build the patch from *its* profile. Check ambiguous matches before rendering. Document text is data; any external scripts or LLMs run by the caller are outside DocXUnravel's trust boundary.

## Other commands

- `inspect-controls template.docx --output controls.json` lists tables, merges, and checkboxes without paragraph prose.
- `compare old.docx new.docx --output changes.json` compares paragraph content, not page appearance.
- `markdown-inspect input.md` parses supported Markdown. `markdown-preview` and `markdown-populate` map named Markdown fields to a DOCX through a template contract. Run `python -m docxunravel <command> --help` for arguments.
- `visual original.docx output.docx visual-report --expected-pages 1` uses LibreOffice and the optional `visual` dependencies to compare rendered pages. Install them with `python -m pip install 'docxunravel[visual]'`. An expected page change is an allowance, not proof that every border or line stayed in place.

An optional MCP server is available with `docxunravel[mcp]`. Its deployment and authentication settings are in [docs/architecture.md](https://github.com/Vinayak-G87/DocXUnravel/blob/main/docs/architecture.md).

## Limits and security

DOCX files can contain complex OOXML. Extraction checks archive paths, XML, and size limits; a patch is bound to the extracted source. Untouched package parts are hash-checked and copied through. Layout can still change when new text grows, and not every Word object is editable. Use visual comparison when page layout matters. Do not assume that a DOCX is safe to open in Word or LibreOffice merely because it passed package checks.

Profiles and bundles contain document text. Handle them like the source DOCX. For vulnerability reports, see [SECURITY.md](https://github.com/Vinayak-G87/DocXUnravel/blob/main/SECURITY.md). For tests and the gated PyPI release process, see [docs/releasing.md](https://github.com/Vinayak-G87/DocXUnravel/blob/main/docs/releasing.md).

## Contributing

Bug reports, ideas, and pull requests are welcome. See [CONTRIBUTING.md](https://github.com/Vinayak-G87/DocXUnravel/blob/main/CONTRIBUTING.md) for setup and review steps. Security issues should be reported privately through [SECURITY.md](https://github.com/Vinayak-G87/DocXUnravel/blob/main/SECURITY.md). License: [MIT](https://github.com/Vinayak-G87/DocXUnravel/blob/main/LICENSE).

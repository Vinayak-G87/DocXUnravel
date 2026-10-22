# DocXUnravel

Current implementation: **v0.4**

DocXUnravel extracts, profiles, edits, and reconstructs DOCX documents using a hybrid representation:

1. A **lossless package** containing every original XML and binary DOCX part.
2. A **semantic JSON profile** that an AI can understand and edit.
3. A **deterministic renderer** that patches selected OOXML and rebuilds the DOCX.

A `.docx` is an Open Packaging Conventions ZIP archive, not one XML file. Keeping the complete package means images, themes, styles, relationships, numbering, embedded files, comments, and unsupported XML remain available even when they are not modeled semantically.

## Setup

```bash
python3 -m venv .venv
.venv/bin/pip install -e .
```

Install optional features only when needed:

```bash
pip install -e '.[visual]'  # visual regression
pip install -e '.[mcp]'     # secure MCP adapter
```

The package exposes `docxunravel`, `python -m docxunravel`, and—with the MCP extra—`docxunravel-mcp`.

## Repository layout

```text
src/docxunravel/       installable core package, CLI, and bundled schemas
tests/                unit/security/fidelity tests, support code, and fixtures
examples/hld/         fictional end-to-end template automation demonstration
examples/patches/     schema 1.1 patch examples
scripts/              repository maintenance and compatibility commands
reports/compatibility retained compatibility baseline
assets/               repository branding only
docs/                 architecture and deployment direction
```

The core package has no network or AI dependency. The optional OAuth-protected MCP server is a thin adapter around the same public Python API; see `docs/architecture.md`.

## v0.4 semantic template automation

DocXUnravel can populate an existing enterprise DOCX from canonical migration-discovery JSON without asking an agent to generate OOXML. The source DOCX remains authoritative for structure, styles, headers, footers, fields, drawings, relationships, and package assets.

The deterministic workflow is:

```text
migration-discovery JSON
→ validated semantic template contract
→ validated schema 1.1 DocXUnravel patch
→ existing OOXML renderer
→ updated DOCX
→ package, semantic, style, and visual fidelity reports
```

The three JSON surfaces are intentionally separate:

| Surface | Version | Purpose |
|---|---:|---|
| Patch contract | `1.1` | Existing low-level validated patch API; remains backward compatible. |
| Template contract | `0.1` | Maps business field names to IDs in one exact semantic profile. |
| Migration discovery | `0.1` | Canonical domain data independent of any particular DOCX. |

Versioned schemas are bundled under `src/docxunravel/schemas/`:

- `src/docxunravel/schemas/template-contract-0.1.schema.json`
- `src/docxunravel/schemas/asset-registry-0.1.schema.json`
- `src/docxunravel/schemas/migration-discovery-0.1.schema.json`

### Template contract

A template contract binds semantic names to profile IDs, never to raw OOXML paths:

```json
{
  "contract_version": "0.1",
  "document_type": "azure_migration_hld",
  "binding": {
    "schema_version": "1.1",
    "source_sha256": "copy from profile.patch_contract",
    "profile_sha256": "copy from profile.patch_contract"
  },
  "fields": {
    "executive_summary": {
      "type": "text",
      "target": "body.p12",
      "required": true,
      "data_path": "/project/scope"
    },
    "application_inventory": {
      "type": "table_rows",
      "target": "body.t2",
      "required": true,
      "template_row": 1,
      "columns": {"id": 0, "name": 1, "owner": 2},
      "data_path": "/application_inventory"
    }
  }
}
```

Supported field types are deliberately narrow:

- `text`: replace one existing paragraph using an explicit formatting policy.
- `table_cell`: replace the single editable paragraph in an existing cell.
- `table_rows`: clone one designated existing row and populate all columns.
- `placeholder`: use the existing schema 1.1 token replacement mechanism.

Validation rejects stale source/profile bindings, unknown IDs, wrong target types, duplicate targets, parent-table/descendant conflicts, invalid column mappings, non-unique placeholder locations, tracked revisions, unsafe control nodes, nested template-row tables, and merged template-row layouts. Optional fields omitted from input are no-ops; required fields must be present and non-null.

### Canonical migration discovery workflow

The included demonstration data is explicitly fictional:

```bash
.venv/bin/docxunravel extract \
  examples/hld/fictional-azure-migration-hld-template.docx demo-bundle

.venv/bin/docxunravel template-validate \
  demo-bundle examples/hld/fictional-hld-template-contract-0.1.json

.venv/bin/docxunravel template-preview \
  demo-bundle \
  examples/hld/fictional-hld-template-contract-0.1.json \
  examples/hld/fictional-migration-discovery-0.1.json \
  --output preview.json

.venv/bin/docxunravel template-patch \
  demo-bundle \
  examples/hld/fictional-hld-template-contract-0.1.json \
  examples/hld/fictional-migration-discovery-0.1.json \
  compiled-patch.json

.venv/bin/docxunravel template-populate \
  demo-bundle \
  examples/hld/fictional-hld-template-contract-0.1.json \
  examples/hld/fictional-migration-discovery-0.1.json \
  generated.docx \
  --patch-output applied-patch.json \
  --fidelity-output fidelity.json
```

`template-preview` never writes a DOCX. Its field-oriented output contains the source element, semantic field, original content, proposed content, operation type, validation warnings, unsupported operations, and the compiled patch.

### Read-only asset registry

Every semantic profile now includes `assets`, a read-only registry for available images, charts, drawing-backed content, embedded objects, media, and external relationships. Entries include content type, package location, relationship ID/type, owning part, SHA-256, byte size, and available drawing dimensions/placement. Binary data is never embedded in the AI-visible profile. Asset creation and replacement are not supported in v0.4; package preservation keeps unsupported assets intact.

### End-to-end fictional HLD demonstration

Run the reproducible demonstration (visual dependencies required) with:

```bash
.venv/bin/pip install -e '.[visual]'
.venv/bin/python examples/hld/generate_demo.py
```

Generated evidence is checked into `examples/hld/`:

- `fictional-azure-migration-hld-template.docx` — authoritative source template
- `fictional-azure-migration-hld-output.docx` — populated result
- `fictional-hld-template-contract-0.1.json` — exact source/profile binding
- `fictional-migration-discovery-0.1.json` — canonical fictional input
- `fictional-hld-dry-run-preview.json` — field-level dry run
- `fictional-hld-compiled-patch.json` — deterministic schema 1.1 patch
- `fictional-hld-package-integrity-report.json` — part and asset hash evidence
- `fictional-hld-fidelity-report.json` — unified fidelity evidence
- `fictional-hld-visual-report.json` — compact eight-page visual result (intermediate PDFs and page images are not retained)

The demonstration changes only `word/document.xml`; all 19 unrelated package parts and the existing architecture image remain hash-identical. The generated and source documents both render to eight pages, with no unexpected visual changes.

## Recommended lossless workflow

### 1. Extract the bundle

```bash
.venv/bin/docxunravel extract customer-template.docx customer-bundle
```

The result is:

```text
customer-bundle/
├── profile.json
└── package/
    ├── [Content_Types].xml
    ├── _rels/.rels
    ├── word/document.xml
    ├── word/styles.xml
    ├── word/_rels/document.xml.rels
    ├── word/media/...
    └── ...every other original package entry
```

`profile.json` contains:

- ordered paragraphs, tables, cells, and nested tables
- headers and footers
- structural IDs such as `body.p3` and `body.t0.r1.c0.p0`
- native Word paragraph IDs or content fingerprints as `stable_id` aliases
- exact OOXML part and XPath locators for paragraphs, tables, and cells
- tracked-revision and mixed-run-formatting risk markers
- runs, common font properties, paragraph formatting, styles, sections, and metadata
- discovered `{{TOKEN}}` placeholders
- SHA-256 hashes for the source and every package entry
- a ready-to-fill `replacement_contract`
- a read-only, relationship-aware asset registry with hashes and placement metadata
- a mandatory `patch_contract` binding generated patches to this source, semantic profile, and package manifest

The package entries are extracted without transformation. During rendering, every untouched entry is hash-verified and copied back without changing its contents.

### Bounded extraction and XML security

Every DOCX is inspected before `python-docx` opens it and streamed to disk under configurable limits. Defaults are:

| Limit | Default |
|---|---:|
| ZIP entries | 2,048 |
| Individual uncompressed entry | 100 MiB |
| Total uncompressed package | 500 MiB |
| Per-entry compression ratio | 200:1 |

Override them for trusted, unusually large documents without changing existing CLI usage:

```bash
.venv/bin/docxunravel extract input.docx bundle \
  --max-entries 4096 \
  --max-entry-size 209715200 \
  --max-total-size 1073741824 \
  --max-compression-ratio 300
```

Extraction rejects path traversal, duplicate names, encrypted entries, malformed ZIPs, oversized entries/packages, and suspicious compression ratios. Actual extracted bytes are counted while streaming, not trusted solely from ZIP metadata.

All package XML uses one hardened parser configuration: entity resolution, DTD loading, external networking, recovery mode, and huge-tree mode are disabled. OOXML parts containing `DOCTYPE` or `ENTITY` declarations are rejected before parsing. DocXUnravel never invokes XInclude or fetches XML resources.

### 2. Have AI return a patch

The AI reads `profile.json` but should return only requested changes:

```json
{
  "contract": {
    "schema_version": "1.1",
    "source_sha256": "copy from profile.patch_contract",
    "profile_sha256": "copy from profile.patch_contract"
  },
  "tokens": {
    "CUSTOMER": "Acme Corporation",
    "SYSTEM_NAME": "Order Platform"
  },
  "elements": {
    "body.p4": "This architecture provides an event-driven order workflow.",
    "body.t0.r1.c1.p0": "Amazon EKS"
  }
}
```

The `contract` object must exactly match `profile.json`'s `patch_contract`; this prevents a patch created from an older or different document from being applied. Token keys omit braces. Both `{{CUSTOMER}}` and `{{ CUSTOMER }}` are recognized, including placeholders split across Word runs. An `elements` value replaces a complete existing paragraph.

For optimistic element-level conflict detection and explicit formatting behavior, use an object:

```json
{
  "elements": {
    "body.p4": {
      "expected": "Original architecture description",
      "value": "New architecture description",
      "format_policy": "reject_mixed_formatting"
    }
  }
}
```

### Replacement policies

`format_policy` is explicit and validated:

| Policy | Behavior |
|---|---|
| `preserve_first_run` | Places all replacement text in the first existing text run. Paragraph properties and every run's formatting XML remain, but the replacement inherits the first run's character formatting. |
| `preserve_common_formatting` | Computes the exact run properties shared by all non-empty runs and applies only those properties to the replacement text. Unique bold/italic/color properties are not silently carried forward. |
| `replace_text_only` | Distributes replacement text over existing text nodes using their original lengths. It does not alter run properties or boundaries, but different text lengths can shift which characters receive emphasis. |
| `reject_mixed_runs` | Rejects full replacement when the profile identifies mixed run formatting. Use a run-level patch instead. |

Legacy names `first_run` and `reject_mixed_formatting` remain accepted as aliases. Mixed-run replacements return a warning for every policy except `reject_mixed_runs`, which returns a validation error. All policies leave paragraph style, numbering, indentation, spacing, table/cell properties, and section properties untouched.

Tracked-revision targets and paragraphs containing fields, hyperlinks, content controls, drawings, embedded objects, tabs, or explicit line/page-break nodes are rejected for full/run-level replacement rather than risking unsupported OOXML corruption. Token replacement remains the preferred targeted mechanism for ordinary placeholders.

### Run-level replacement

Use `runs` when one known run should change and neighboring formatting must remain semantically attached:

```json
{
  "contract": {"schema_version": "1.1", "source_sha256": "...", "profile_sha256": "..."},
  "runs": {
    "body.p4": [
      {
        "index": 1,
        "expected": "Kubernetes",
        "value": "Amazon EKS"
      }
    ]
  }
}
```

Run indexes and expected text are checked against the bound profile and again against live OOXML. A paragraph cannot receive both a full replacement and run-level replacements in one patch.

A useful AI instruction is:

> Read the supplied document profile and return JSON only. Preserve all existing IDs. Put placeholder values in `tokens`, full paragraph replacements in `elements`, and requested structural changes in `operations`. Do not reproduce or modify locators, formatting metadata, or package data.

### 3. Preview and rebuild the DOCX

Validate the patch and inspect a human-readable change plan first:

```bash
.venv/bin/docxunravel preview customer-bundle values.json
```

Then render:

```bash
.venv/bin/docxunravel render-bundle \
  customer-bundle values.json generated.docx
```

This directly edits the located XML nodes, retains all untouched package parts, verifies their hashes, and creates a new DOCX. Output is written to a temporary file, checked for required parts, XML validity, content-type coverage, relationship integrity, and `python-docx` openability, then atomically published with `os.replace()`.

The operation examples below show only the `operations` fragment; a complete patch must also include the matching `contract`.

## Structural generation

The bundle renderer supports five deliberately small operations:

- `insert_before`
- `insert_after`
- `append_to_cell`
- `delete`
- `repeat_table_rows` (contract-generated, existing-template-row cloning only)

Targets are resolved before structural mutation, so one insertion cannot invalidate the locator of a later operation. A patch is rejected when multiple structural operations target the same original element, an element is edited and deleted together, a modified target is inside a deleted ancestor, a cell is deleted directly, or a target contains tracked/unsupported OOXML structures.

### Add a paragraph

```json
{
  "operations": [
    {
      "op": "insert_after",
      "target": "body.p4",
      "element": {
        "type": "paragraph",
        "style": "Normal",
        "text": "The platform is deployed across three availability zones."
      }
    }
  ]
}
```

### Add a table

```json
{
  "operations": [
    {
      "op": "insert_after",
      "target": "body.p4",
      "element": {
        "type": "table",
        "style": "Table Grid",
        "rows": [
          ["Component", "Technology", "Purpose"],
          ["API", "FastAPI", "Request processing"],
          ["Compute", "Amazon EKS", "Container orchestration"]
        ]
      }
    }
  ]
}
```

### Add a table inside a table cell

Cells accept recursive `elements`, so nesting is not limited to one level:

```json
{
  "operations": [
    {
      "op": "append_to_cell",
      "target": "body.t0.r1.c1",
      "element": {
        "type": "table",
        "style": "Table Grid",
        "rows": [
          ["Property", "Value"],
          ["Replicas", "3"],
          ["Region", "us-east-1"]
        ]
      }
    }
  ]
}
```

A generated table cell can itself contain paragraphs and nested tables:

```json
{
  "type": "table",
  "rows": [
    [
      "Compute",
      {
        "elements": [
          {"type": "paragraph", "text": "Cluster configuration"},
          {
            "type": "table",
            "style": "Table Grid",
            "rows": [["Nodes", "3"]]
          }
        ]
      }
    ]
  ]
}
```

## Optional visual regression

Visual equivalence is not inferred from XML validity. The optional workflow renders both DOCX files through LibreOffice, rasterizes their PDFs, compares page dimensions and pixels, writes three images per comparable page, and emits a JSON report.

Install optional dependencies separately:

```bash
.venv/bin/pip install -e '.[visual]'
```

LibreOffice (`libreoffice` or `soffice`) must also be available on `PATH`; it is never required for extraction, profiling, patch validation, or rendering.

```bash
.venv/bin/docxunravel visual \
  original.docx generated.docx visual-report \
  --difference-threshold 0.002 \
  --expected-difference-threshold 0.10 \
  --pixel-threshold 12 \
  --dimension-tolerance-pt 0.5 \
  --expected-pages 1,3-4
```

The command exits `0` when no unexpected changed pages are detected and `2` otherwise. `--expected-pages` is not blanket approval: page-count changes and dimension changes always fail, and expected pages must remain below `--expected-difference-threshold`.

For a strict no-op round trip, expected pages are forbidden and every pixel difference above the configured noise threshold is unexpected:

```bash
.venv/bin/docxunravel visual \
  original.docx noop-output.docx visual-noop \
  --no-op \
  --difference-threshold 0
```

Example report excerpt:

```json
{
  "passed": false,
  "original_page_count": 3,
  "generated_page_count": 4,
  "page_count_match": false,
  "expected_changed_pages": [1],
  "unexpected_changed_pages": [2, 3, 4],
  "pages": [
    {
      "page": 1,
      "status": "expected_change",
      "difference_ratio": 0.0124,
      "dimension_match": true,
      "images": {
        "original": "page-001-original.png",
        "generated": "page-001-generated.png",
        "difference": "page-001-diff.png"
      }
    }
  ]
}
```

Results depend on the selected layout engine, installed fonts, and platform. For highest fidelity to Microsoft Word, a deployment may substitute a Word-based PDF renderer while retaining the PDF comparison function.

## Unified fidelity report

The core-only fidelity command combines evidence without inferring visual correctness from package integrity:

```bash
.venv/bin/docxunravel fidelity-report \
  original.docx generated.docx fidelity.json \
  --bundle template-bundle \
  --patch values.json \
  --visual-report visual-report/visual-report.json
```

The JSON independently reports:

- original and generated package integrity and enforced limits
- expected modified parts, changed parts, missing/added parts, and untouched hash preservation
- semantic text hashes, expected patch result, and unexpected added/removed text
- `styles.xml` and `numbering.xml` hashes
- whether visual verification was performed and its result
- unsupported OOXML structures encountered in the source profile

A report can pass without visual verification only with `visual_differences.performed: false`; it never claims visual equivalence in that case. For no-op reports, semantic text must match exactly. For patch reports, generated semantic text must match the patch-derived expected multiset.

## Regression fixtures and performance

`tests/fixtures/complex_sections_comments_revisions.docx` is reproducibly generated by `tests/support/fixture_factory.py` and covers nested/merged tables, three sections, portrait/landscape pages, distinct headers/footers, bookmarks, comments, tracked insertions/deletions, numbering, and images. See `tests/fixtures/README.md` for regeneration.

Run all behavior, security, fidelity, visual, and performance tests with:

```bash
.venv/bin/python -m unittest discover -v
```

The performance suite builds small, medium, and large synthetic documents and measures profiling, extraction, and no-op rendering with generous regression ceilings rather than brittle micro-benchmarks.

## Python API

```python
from docxunravel import export_bundle, render_bundle
from docxunravel.template_contract import (
    compile_migration_template_patch,
    populate_migration_template,
    preview_migration_template_population,
    validate_template_contract,
)

profile = export_bundle("template.docx", "template-bundle")
result = render_bundle(
    "template-bundle",
    "output.docx",
    {
        "contract": profile["patch_contract"],
        "tokens": {"CUSTOMER": "Acme"},
        "elements": {"body.p2": "Generated architecture summary"},
    },
)
```

The original simpler APIs remain available:

```bash
.venv/bin/docxunravel profile template.docx profile.json
.venv/bin/docxunravel render template.docx values.json output.docx --profile profile.json
```

Those use `python-docx` directly. Use the bundle workflow when package preservation and raw OOXML access matter.

## Secure MCP adapter

Install and run locally over stdio:

```bash
pip install -e '.[mcp]'
DOCXUNRAVEL_MCP_WORKSPACE=/srv/docxunravel docxunravel-mcp --transport stdio
```

For Streamable HTTP, DocXUnravel is an OAuth 2.1 resource server backed by an external OIDC provider. It does not implement users, passwords, login pages, or token issuance:

```bash
export DOCXUNRAVEL_MCP_WORKSPACE=/srv/docxunravel
export DOCXUNRAVEL_MCP_TRANSPORT=streamable-http
export DOCXUNRAVEL_MCP_PUBLIC_URL=https://documents.example.com/mcp
export DOCXUNRAVEL_MCP_OIDC_ISSUER=https://login.example.com/tenant/v2.0
export DOCXUNRAVEL_MCP_OIDC_AUDIENCE=api://docxunravel
export DOCXUNRAVEL_MCP_ALLOWED_CLIENT_IDS=approved-agent-client

docxunravel-mcp
```

Provision source templates as `/srv/docxunravel/documents/<opaque-id>.docx`. MCP clients use only the opaque ID and cannot provide filesystem paths. Available tools are:

- `list_documents`
- `profile_document`
- `preview_template_population`
- `populate_template`
- `validate_output`

HTTP tools enforce cumulative `documents:read`, `documents:preview`, and `documents:write` scopes. The server publishes RFC 9728 metadata, validates signed JWTs through OIDC discovery and rotating JWKS, accepts bearer tokens only through headers, returns no binary/Base64 DOCX content, and preserves preview-before-write plus atomic output behavior. Deployment and provider guidance is in `docs/architecture.md`.

## Guarantees and boundaries

### What is preserved

- Every original package entry is retained.
- Untouched entries are checked against their extraction-time SHA-256 hashes, and the complete package manifest is authenticated by the profile contract.
- Unknown XML and binary content can survive without semantic support.
- Only XML parts containing targeted changes are reserialized.
- The resulting file is functionally reconstructed, not promised to be byte-for-byte identical to the original ZIP.

### What still requires care

1. **Formatting inheritance:** effective appearance can come from styles, defaults, and themes. The profile provides direct formatting and useful inheritance hints, not a complete Word layout engine.
2. **Mixed inline formatting:** no policy can infer which new words should be emphasized. Use `reject_mixed_runs`, targeted `runs`, or placeholders when semantic formatting boundaries matter.
3. **Advanced structural objects:** package preservation keeps text boxes, SmartArt, charts, content controls, fields, comments, tracked changes, and embedded files, but the current semantic operations do not create or deeply edit all of them. Targets containing tracked revisions are rejected.
4. **Relationships:** adding a new image, chart, hyperlink, or embedded object requires allocating relationship IDs and updating relationships and content types. Existing related objects are preserved.
5. **Visual preservation:** package integrity is checked deterministically, but Word performs pagination. The optional visual command detects rendered changes; it cannot prove cross-renderer equivalence, and expected pages are a page-level allowance rather than semantic region masking.
6. **Merged cells:** Word can expose one merged cell through multiple grid positions. The profile marks duplicate merged cells; use the first corresponding cell ID for ordinary cell targeting. Repeatable template rows containing horizontal or vertical merges are rejected in v0.4.
7. **Security:** profiles and extracted XML contain customer text. Apply the same privacy controls used for the source document. Resource limits reduce denial-of-service risk but do not make untrusted documents safe to open in Microsoft Word or LibreOffice.
8. **Compression limits:** ratio checks use ZIP metadata before extraction and actual byte counters while streaming. They do not estimate CPU cost for every possible compressor implementation.
9. **Semantic comparison:** patch fidelity compares paragraph-text multisets to tolerate structural index movement. Documents containing many identical paragraphs can reduce localization precision even though count mismatches are still detected.
10. **Repeatable rows:** v0.4 requires one paragraph per template-row cell, maps every table column, and rejects nested tables, merged cells, revisions, fields, drawings, and other unsafe editable structures in that row. It clones existing row/cell/paragraph/run properties but does not infer or repair malformed layouts.
11. **Rich content and assets:** semantic population is text and row oriented. It does not generate unrestricted rich text or create/replace images, charts, SmartArt, OLE objects, hyperlinks, or embedded packages.
12. **Discovery prose:** canonical discovery values are mapped directly. DocXUnravel does not call an LLM or synthesize narrative prose; applications may generate approved text before invoking the engine.

The architecture is therefore extensible without asking the AI to edit raw XML: add deterministic operations for images, charts, numbering, or prototype cloning as those use cases become necessary.

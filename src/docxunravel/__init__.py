"""Public DocXUnravel Python API."""

from .core import (
    DEFAULT_PACKAGE_LIMITS,
    SCHEMA_VERSION,
    UNSAFE_EDIT_TAGS,
    PackageLimits,
    create_fidelity_report,
    export_bundle,
    inspect_docx_archive,
    preview_bundle,
    profile_document,
    render_bundle,
    render_document,
    validate_docx_package,
    validate_patch,
    write_fidelity_report,
)
from .migration_discovery import validate_migration_discovery
from .template_contract import (
    compile_migration_template_patch,
    compile_template_patch,
    populate_migration_template,
    preview_migration_template_population,
    preview_template_population,
    template_contract_binding,
    validate_template_contract,
)

__version__ = "0.4.0"

__all__ = [
    "DEFAULT_PACKAGE_LIMITS",
    "PackageLimits",
    "SCHEMA_VERSION",
    "UNSAFE_EDIT_TAGS",
    "compile_migration_template_patch",
    "compile_template_patch",
    "create_fidelity_report",
    "export_bundle",
    "inspect_docx_archive",
    "populate_migration_template",
    "preview_bundle",
    "preview_migration_template_population",
    "preview_template_population",
    "profile_document",
    "render_bundle",
    "render_document",
    "template_contract_binding",
    "validate_docx_package",
    "validate_migration_discovery",
    "validate_patch",
    "validate_template_contract",
    "write_fidelity_report",
]

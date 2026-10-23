"""Secure, optional MCP adapter for DocXUnravel's deterministic Python API."""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import os
import re
import secrets
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from mcp.server import MCPServer
from mcp.server.auth.middleware.auth_context import get_access_token
from mcp.server.auth.provider import TokenVerifier
from mcp.server.auth.settings import AuthSettings
from pydantic import AnyHttpUrl
from starlette.requests import Request
from starlette.responses import JSONResponse

from . import __version__
from .core import create_fidelity_report, export_bundle, validate_docx_package
from .mcp_auth import OIDCJWTVerifier
from .template_contract import populate_template, preview_template_population

READ_SCOPE = "documents:read"
PREVIEW_SCOPE = "documents:preview"
WRITE_SCOPE = "documents:write"
_DOCUMENT_ID = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}\Z")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _atomic_json(path: Path, value: dict[str, Any]) -> None:
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, ensure_ascii=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


class DocumentWorkspace:
    """Server-owned storage addressed only by opaque document identifiers."""

    def __init__(self, root: str | Path) -> None:
        configured_root = Path(root).expanduser()
        configured_root.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.root = configured_root.resolve(strict=True)
        if not self.root.is_dir():
            raise ValueError("MCP workspace must be a directory")
        self.documents = self._directory("documents")
        self.outputs = self._directory("outputs")
        self.bundles = self._directory("bundles")
        self.reports = self._directory("reports")

    def _directory(self, name: str) -> Path:
        path = self.root / name
        if path.is_symlink():
            raise ValueError(f"MCP workspace directory cannot be a symlink: {name}")
        path.mkdir(mode=0o700, exist_ok=True)
        resolved = path.resolve(strict=True)
        if not resolved.is_dir() or resolved.parent != self.root:
            raise ValueError(f"MCP workspace directory escapes its root: {name}")
        return resolved

    @staticmethod
    def validate_id(document_id: str) -> str:
        if not isinstance(document_id, str) or not _DOCUMENT_ID.fullmatch(document_id):
            raise ValueError(
                "Document ID must contain 1-64 ASCII letters, digits, hyphens, or underscores"
            )
        return document_id

    def _document_path(self, document_id: str, *, outputs_only: bool = False) -> Path:
        document_id = self.validate_id(document_id)
        directories = (
            (self.outputs,) if outputs_only else (self.documents, self.outputs)
        )
        for directory in directories:
            candidate = directory / f"{document_id}.docx"
            if candidate.is_symlink():
                raise ValueError("Symbolic-link documents are not allowed")
            try:
                resolved = candidate.resolve(strict=True)
            except FileNotFoundError:
                continue
            if resolved.parent != directory or not resolved.is_file():
                raise ValueError("Document path escapes the configured workspace")
            return resolved
        raise KeyError(f"Unknown document ID: {document_id}")

    def _source_path(self, document_id: str) -> Path:
        document_id = self.validate_id(document_id)
        candidate = self.documents / f"{document_id}.docx"
        if candidate.is_symlink():
            raise ValueError("Symbolic-link documents are not allowed")
        try:
            resolved = candidate.resolve(strict=True)
        except FileNotFoundError as error:
            raise KeyError(f"Unknown template ID: {document_id}") from error
        if resolved.parent != self.documents or not resolved.is_file():
            raise ValueError("Template path escapes the configured workspace")
        return resolved

    def list_documents(self) -> dict[str, Any]:
        items: list[dict[str, Any]] = []
        for kind, directory in (("source", self.documents), ("output", self.outputs)):
            for path in sorted(directory.glob("*.docx")):
                if path.is_symlink() or not path.is_file():
                    continue
                try:
                    document_id = self.validate_id(path.stem)
                except ValueError:
                    continue
                items.append(
                    {
                        "document_id": document_id,
                        "kind": kind,
                        "size": path.stat().st_size,
                        "sha256": _sha256(path),
                    }
                )
        return {"documents": items, "count": len(items)}

    @contextmanager
    def _temporary_bundle(self, source: Path) -> Iterator[tuple[Path, dict[str, Any]]]:
        with tempfile.TemporaryDirectory(
            prefix="bundle-", dir=self.bundles
        ) as temporary:
            bundle = Path(temporary) / "content"
            profile = export_bundle(source, bundle)
            yield bundle, profile

    def profile(self, document_id: str) -> dict[str, Any]:
        source = self._document_path(document_id)
        with self._temporary_bundle(source) as (_, profile):
            return profile

    def preview(
        self,
        template_id: str,
        contract: dict[str, Any],
        data: dict[str, Any],
    ) -> dict[str, Any]:
        source = self._source_path(template_id)
        with self._temporary_bundle(source) as (bundle, _):
            return preview_template_population(bundle, contract, data)

    def populate(
        self,
        template_id: str,
        contract: dict[str, Any],
        data: dict[str, Any],
    ) -> dict[str, Any]:
        source = self._source_path(template_id)
        while True:
            output_id = secrets.token_hex(16)
            output = self.outputs / f"{output_id}.docx"
            report = self.reports / f"{output_id}.json"
            source_collision = self.documents / f"{output_id}.docx"
            if (
                not output.exists()
                and not report.exists()
                and not source_collision.exists()
            ):
                break

        with (
            tempfile.TemporaryDirectory(
                prefix="render-", dir=self.outputs
            ) as temporary,
            self._temporary_bundle(source) as (bundle, _),
        ):
            temporary_root = Path(temporary)
            temporary_output = temporary_root / "output.docx"
            temporary_fidelity = temporary_root / "fidelity.json"
            result = populate_template(
                bundle,
                temporary_output,
                contract,
                data,
                fidelity_output=temporary_fidelity,
            )
            fidelity = json.loads(temporary_fidelity.read_text(encoding="utf-8"))
            patch = result["preview"]["patch"]
            record = {
                "schema_version": "1.0",
                "source_id": template_id,
                "source_sha256": _sha256(source),
                "output_sha256": _sha256(temporary_output),
                "patch": patch,
                "fidelity": fidelity,
            }
            try:
                os.replace(temporary_output, output)
                _atomic_json(report, record)
            except Exception:
                output.unlink(missing_ok=True)
                report.unlink(missing_ok=True)
                raise

        render = result["render"]
        return {
            "output_id": output_id,
            "source_id": template_id,
            "source_sha256": record["source_sha256"],
            "output_sha256": record["output_sha256"],
            "changes": result["preview"]["changes"],
            "validation_warnings": result["preview"]["validation_warnings"],
            "render": {
                "changed_elements": render["changed_elements"],
                "changed_runs": render["changed_runs"],
                "token_replacements": render["token_replacements"],
                "structural_changes": render["structural_changes"],
                "modified_parts": render["modified_parts"],
                "integrity": render["integrity"],
            },
            "fidelity_passed": fidelity["passed"],
        }

    def validate(self, document_id: str) -> dict[str, Any]:
        document = self._document_path(document_id)
        package = validate_docx_package(document)
        report_path = self.reports / f"{self.validate_id(document_id)}.json"
        if not report_path.exists():
            return {
                "document_id": document_id,
                "sha256": _sha256(document),
                "package": package,
                "fidelity": None,
            }
        if (
            report_path.is_symlink()
            or report_path.resolve(strict=True).parent != self.reports
        ):
            raise ValueError("Output report escapes the configured workspace")
        record = json.loads(report_path.read_text(encoding="utf-8"))
        if not isinstance(record, dict):
            raise TypeError("Output report is invalid")
        current_hash = _sha256(document)
        if current_hash != record.get("output_sha256"):
            return {
                "document_id": document_id,
                "sha256": current_hash,
                "package": package,
                "fidelity": {"passed": False, "stale": True},
            }
        source_id = record.get("source_id")
        patch = record.get("patch")
        if not isinstance(source_id, str) or not isinstance(patch, dict):
            raise TypeError("Output report has invalid provenance")
        source = self._source_path(source_id)
        if _sha256(source) != record.get("source_sha256"):
            return {
                "document_id": document_id,
                "sha256": current_hash,
                "package": package,
                "fidelity": {"passed": False, "stale": True},
            }
        with self._temporary_bundle(source) as (bundle, _):
            fidelity = create_fidelity_report(
                source, document, bundle=bundle, patch=patch
            )
        return {
            "document_id": document_id,
            "sha256": current_hash,
            "package": package,
            "fidelity": fidelity,
        }

    def ready(self) -> bool:
        return all(
            directory.is_dir() and not directory.is_symlink()
            for directory in (self.documents, self.outputs, self.bundles, self.reports)
        )


@dataclass(frozen=True)
class HTTPAuthConfig:
    issuer: str
    audience: str
    resource_server_url: str
    algorithms: tuple[str, ...] = ("RS256", "ES256")
    allowed_client_ids: tuple[str, ...] = ()
    allow_insecure_http: bool = False


def _require_http_url(url: str, label: str, allow_insecure_http: bool) -> str:
    parsed = urlparse(url)
    schemes = {"https"} | ({"http"} if allow_insecure_http else set())
    if parsed.scheme not in schemes or not parsed.netloc or parsed.username:
        requirement = "HTTP(S)" if allow_insecure_http else "HTTPS"
        raise ValueError(f"{label} must be an absolute {requirement} URL")
    if parsed.query or parsed.fragment:
        raise ValueError(f"{label} cannot contain a query string or fragment")
    return url


def _require_public_url(url: str, allow_insecure_http: bool) -> str:
    url = _require_http_url(url, "MCP public URL", allow_insecure_http)
    if not urlparse(url).path.endswith("/mcp"):
        raise ValueError("MCP public URL must end in /mcp")
    return url


def _require_scopes(scopes: set[str], auth_enabled: bool) -> None:
    if not auth_enabled:
        return
    token = get_access_token()
    if token is None:
        raise PermissionError("Authentication is required")
    missing = scopes - set(token.scopes)
    if missing:
        raise PermissionError(
            f"Missing required OAuth scopes: {', '.join(sorted(missing))}"
        )


def create_mcp_server(
    workspace: DocumentWorkspace,
    *,
    auth: HTTPAuthConfig | None = None,
    token_verifier: TokenVerifier | None = None,
) -> MCPServer:
    """Create the MCP server; HTTP auth is optional only for local stdio use."""

    auth_settings = None
    if auth is not None:
        issuer_url = _require_http_url(
            auth.issuer, "OIDC issuer", auth.allow_insecure_http
        )
        resource_url = _require_public_url(
            auth.resource_server_url, auth.allow_insecure_http
        )
        if not auth.audience or any(character.isspace() for character in auth.audience):
            raise ValueError("OIDC audience must be a non-empty identifier")
        verifier = token_verifier or OIDCJWTVerifier(
            issuer=auth.issuer,
            audience=auth.audience,
            algorithms=auth.algorithms,
            allowed_client_ids=auth.allowed_client_ids,
            allow_insecure_http=auth.allow_insecure_http,
        )
        token_verifier = verifier
        auth_settings = AuthSettings(
            issuer_url=AnyHttpUrl(issuer_url),
            resource_server_url=AnyHttpUrl(resource_url),
            required_scopes=[READ_SCOPE],
            validate_token_resource=False,
        )
    elif token_verifier is not None:
        raise ValueError("A token verifier requires HTTP auth configuration")

    server = MCPServer(
        name="docxunravel",
        title="DocXUnravel",
        description="Deterministic semantic DOCX template automation",
        instructions=(
            "Use preview_template_population before populate_template. "
            "Tools accept opaque document IDs, never filesystem paths or binary DOCX data."
        ),
        version=__version__,
        token_verifier=token_verifier,
        auth=auth_settings,
    )
    auth_enabled = auth is not None

    @server.tool()
    async def list_documents() -> dict[str, Any]:
        """List server-managed document metadata without returning DOCX binaries."""
        _require_scopes({READ_SCOPE}, auth_enabled)
        return await asyncio.to_thread(workspace.list_documents)

    @server.tool()
    async def profile_document(document_id: str) -> dict[str, Any]:
        """Return the semantic profile and read-only asset registry for a document."""
        _require_scopes({READ_SCOPE}, auth_enabled)
        return await asyncio.to_thread(workspace.profile, document_id)

    @server.tool()
    async def preview_template_population(
        template_id: str,
        contract: dict[str, Any],
        data: dict[str, Any],
    ) -> dict[str, Any]:
        """Validate and preview template data without creating output."""
        _require_scopes({READ_SCOPE, PREVIEW_SCOPE}, auth_enabled)
        return await asyncio.to_thread(workspace.preview, template_id, contract, data)

    @server.tool()
    async def populate_template(
        template_id: str,
        contract: dict[str, Any],
        data: dict[str, Any],
    ) -> dict[str, Any]:
        """Populate a template atomically and return a new opaque output ID."""
        _require_scopes({READ_SCOPE, PREVIEW_SCOPE, WRITE_SCOPE}, auth_enabled)
        return await asyncio.to_thread(workspace.populate, template_id, contract, data)

    @server.tool()
    async def validate_output(document_id: str) -> dict[str, Any]:
        """Validate package integrity and available fidelity provenance."""
        _require_scopes({READ_SCOPE}, auth_enabled)
        return await asyncio.to_thread(workspace.validate, document_id)

    @server.custom_route("/healthz", methods=["GET"], include_in_schema=False)
    async def health(_: Request) -> JSONResponse:
        return JSONResponse({"status": "ok"})

    @server.custom_route("/readyz", methods=["GET"], include_in_schema=False)
    async def readiness(_: Request) -> JSONResponse:
        ready = workspace.ready()
        return JSONResponse(
            {"status": "ready" if ready else "not-ready"},
            status_code=200 if ready else 503,
        )

    return server


def _environment_boolean(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"{name} must be a boolean")


def _csv_environment(name: str, default: str) -> tuple[str, ...]:
    return tuple(
        value.strip()
        for value in os.environ.get(name, default).split(",")
        if value.strip()
    )


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the optional DocXUnravel MCP server"
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "streamable-http"),
        default=os.environ.get("DOCXUNRAVEL_MCP_TRANSPORT", "stdio"),
    )
    parser.add_argument(
        "--workspace", default=os.environ.get("DOCXUNRAVEL_MCP_WORKSPACE")
    )
    parser.add_argument(
        "--host", default=os.environ.get("DOCXUNRAVEL_MCP_HOST", "127.0.0.1")
    )
    parser.add_argument(
        "--port", type=int, default=int(os.environ.get("DOCXUNRAVEL_MCP_PORT", "8000"))
    )
    parser.add_argument(
        "--public-url", default=os.environ.get("DOCXUNRAVEL_MCP_PUBLIC_URL")
    )
    parser.add_argument(
        "--issuer", default=os.environ.get("DOCXUNRAVEL_MCP_OIDC_ISSUER")
    )
    parser.add_argument(
        "--audience", default=os.environ.get("DOCXUNRAVEL_MCP_OIDC_AUDIENCE")
    )
    parser.add_argument(
        "--max-request-bytes",
        type=int,
        default=int(
            os.environ.get("DOCXUNRAVEL_MCP_MAX_REQUEST_BYTES", str(2 * 1024 * 1024))
        ),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if not args.workspace:
        raise SystemExit("DOCXUNRAVEL_MCP_WORKSPACE or --workspace is required")
    if args.port < 1 or args.port > 65535 or args.max_request_bytes < 1024:
        raise SystemExit("Invalid MCP port or request-size limit")

    workspace = DocumentWorkspace(args.workspace)
    if args.transport == "stdio":
        create_mcp_server(workspace).run(transport="stdio")
        return 0

    missing = [
        name
        for name, value in (
            ("public URL", args.public_url),
            ("OIDC issuer", args.issuer),
            ("OIDC audience", args.audience),
        )
        if not value
    ]
    if missing:
        raise SystemExit(
            "Streamable HTTP requires "
            + ", ".join(missing)
            + "; unauthenticated HTTP is disabled"
        )
    allow_insecure = _environment_boolean("DOCXUNRAVEL_MCP_ALLOW_INSECURE_HTTP")
    auth = HTTPAuthConfig(
        issuer=args.issuer,
        audience=args.audience,
        resource_server_url=args.public_url,
        algorithms=_csv_environment("DOCXUNRAVEL_MCP_OIDC_ALGORITHMS", "RS256,ES256"),
        allowed_client_ids=_csv_environment("DOCXUNRAVEL_MCP_ALLOWED_CLIENT_IDS", ""),
        allow_insecure_http=allow_insecure,
    )
    server = create_mcp_server(workspace, auth=auth)
    server.run(
        transport="streamable-http",
        host=args.host,
        port=args.port,
        streamable_http_path="/mcp",
        json_response=True,
        stateless_http=True,
        max_request_body_size=args.max_request_bytes,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

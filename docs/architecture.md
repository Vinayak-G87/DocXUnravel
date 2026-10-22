# Production architecture

## Package boundary

`docxunravel` is a deterministic Python library. It owns bounded DOCX inspection, semantic profiles, template contracts, patch compilation, OOXML rendering, and fidelity evidence. It does not call an LLM, issue OAuth tokens, or fetch document relationships.

The optional MCP adapter is limited to `mcp_server.py` and `mcp_auth.py`. It invokes the same public functions as local Python clients; authentication and transport behavior do not enter `core.py` or weaken document validation.

## MCP deployment modes

Install the adapter with:

```bash
pip install 'docxunravel[mcp]'
```

Two transports are supported:

- **stdio:** local-process trust. OAuth does not apply because stdio has no HTTP authorization header.
- **Streamable HTTP:** OAuth 2.1 resource server mode. Startup is refused unless an external OIDC issuer, audience, and public resource URL are configured.

The HTTP server never provides login, password, authorization, token, callback, or refresh-token endpoints. Use a standards-compliant authorization server such as Microsoft Entra ID, Auth0, or Keycloak. Interactive clients perform Authorization Code with PKCE with that provider; DocXUnravel validates the resulting access token.

## Security model

The OIDC verifier:

1. Requires HTTPS for the issuer, JWKS endpoint, and public MCP URL. Plain HTTP requires an explicit development-only override.
2. Uses OIDC discovery and bounded JWKS responses, caches keys, and refreshes for key rotation.
3. Accepts only configured RSA/ECDSA signature algorithms; `alg=none` and symmetric JWT algorithms are rejected.
4. Validates signature, exact issuer, audience, expiry, not-before when present, and authorized client identity (`azp`, `client_id`, or Entra `appid`). Multi-audience tokens require `azp`.
5. Can restrict tokens to an allowlist of OAuth client IDs.
6. Exposes only minimal claims to MCP request context and never writes access tokens to patches, bundles, outputs, or reports.

The SDK publishes RFC 9728 Protected Resource Metadata at `/.well-known/oauth-protected-resource/mcp`. Bearer tokens are accepted only from the `Authorization` header; query-string tokens do not authenticate.

Scopes are cumulative:

| Operation | Required scopes |
|---|---|
| `list_documents`, `profile_document`, `validate_output` | `documents:read` |
| `preview_template_population` | `documents:read documents:preview` |
| `populate_template` | `documents:read documents:preview documents:write` |

The SDK-level HTTP gate requires `documents:read`; the adapter performs the additional operation-level checks. Configure the authorization server to issue only the scopes each client needs.

## Workspace and document isolation

Set `DOCXUNRAVEL_MCP_WORKSPACE` to a dedicated server directory. At startup the adapter creates:

```text
workspace/
├── documents/  # operator-provisioned source templates: <opaque-id>.docx
├── outputs/    # generated documents
├── reports/    # internal patch/fidelity provenance
└── bundles/    # temporary extraction only; automatically removed
```

Clients submit only opaque IDs containing ASCII letters, digits, hyphens, and underscores. They cannot submit filesystem paths. Traversal, separators, extensions, symbolic-link documents, and workspace directories that escape through symlinks are rejected. DOCX binaries are not returned through MCP tools or embedded as Base64 in profiles.

Source templates should be mounted read-only at the infrastructure boundary where possible. The process needs write access to `outputs`, `reports`, and `bundles`. Generated documents are published with the existing atomic renderer and receive a random 128-bit output ID. A failed population does not publish an output or report.

The initial adapter deliberately has no upload/download API. Provision inputs and retrieve outputs through a separately authenticated storage or application boundary with its own size, malware-scanning, tenancy, retention, and audit policy.

## Configuration

Example production environment:

```bash
export DOCXUNRAVEL_MCP_WORKSPACE=/srv/docxunravel
export DOCXUNRAVEL_MCP_TRANSPORT=streamable-http
export DOCXUNRAVEL_MCP_HOST=127.0.0.1
export DOCXUNRAVEL_MCP_PORT=8000
export DOCXUNRAVEL_MCP_PUBLIC_URL=https://documents.example.com/mcp
export DOCXUNRAVEL_MCP_OIDC_ISSUER=https://login.example.com/tenant/v2.0
export DOCXUNRAVEL_MCP_OIDC_AUDIENCE=api://docxunravel
export DOCXUNRAVEL_MCP_OIDC_ALGORITHMS=RS256
export DOCXUNRAVEL_MCP_ALLOWED_CLIENT_IDS=approved-agent-client
export DOCXUNRAVEL_MCP_MAX_REQUEST_BYTES=2097152

docxunravel-mcp
```

Provider mapping:

- **Entra ID:** issuer is the tenant-specific v2 issuer; audience is the API application ID URI or application ID; scopes are exposed API scopes. `azp`/`appid` can be allowlisted.
- **Auth0:** issuer is the tenant URL and audience is the API identifier.
- **Keycloak:** issuer is the realm URL and audience is the configured resource/client audience.

Issuer strings are compared exactly. Confirm the discovery document's `issuer` value and use that exact value in configuration.

`DOCXUNRAVEL_MCP_ALLOW_INSECURE_HTTP=true` exists only for loopback integration testing. Do not enable it in production.

## Network and process deployment

- Terminate TLS at a trusted reverse proxy or ingress and preserve the public `/mcp` URL exactly.
- Bind the application to loopback or a private interface when TLS terminates upstream.
- Do not forward inbound access tokens to downstream services.
- Apply request-rate, connection, concurrency, and total body-size limits at the proxy. The application also defaults to a 2 MiB MCP request limit.
- Run as a dedicated non-root identity with a restrictive umask and minimal filesystem permissions.
- Keep outbound network access restricted to the configured OIDC discovery/JWKS hosts. OOXML external relationships are preserved but never fetched.
- Use `/healthz` for liveness and `/readyz` for workspace readiness. These endpoints disclose only status and intentionally do not require OAuth.
- Log tool names, principal/client IDs, opaque document IDs, hashes, outcomes, and durations where needed; never log authorization headers, raw JWT claims, discovery payloads, document text, contracts, patches, or binary content.
- Add tenant-specific workspace roots, quotas, encryption, retention, deletion, and audit sinks in the hosting application before serving multiple tenants.

## Python plugin deployment

The repository uses a `src/` layout and `pyproject.toml`. Consumers can install `docxunravel`, import its public API, and provide discovery/contract JSON from any application or model. Optional visual and MCP dependencies are isolated behind `visual` and `mcp` extras.

There is intentionally no custom plugin framework yet. If plugin discovery becomes necessary, use standard Python entry points with a small typed protocol and keep untrusted plugins outside the renderer process.

## Current limitations

- MCP does not upload or download DOCX binaries; storage integration belongs to the host application.
- Authorization is client/scope based. Multi-tenant subject-to-workspace policy is a hosting concern and is not inferred automatically.
- JWT access tokens are supported; opaque-token introspection is not implemented.
- JWKS/discovery caching is in-process. Large multi-replica deployments should use ordinary provider caching and rate controls.
- Visual regression is not run automatically by MCP population; package and semantic fidelity are reported independently.

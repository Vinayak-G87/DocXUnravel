"""External OIDC access-token verification for the optional MCP adapter."""

from __future__ import annotations

import asyncio
import json
import time
import urllib.request
from collections.abc import Mapping, Sequence
from typing import Any
from urllib.parse import urlparse

import jwt
from mcp.server.auth.provider import AccessToken

_SAFE_ALGORITHMS = frozenset({"RS256", "RS384", "RS512", "ES256", "ES384", "ES512"})
_MAX_METADATA_BYTES = 1024 * 1024


def _require_secure_url(value: str, label: str, allow_insecure_http: bool) -> str:
    parsed = urlparse(value)
    allowed_schemes = {"https"} | ({"http"} if allow_insecure_http else set())
    if parsed.scheme not in allowed_schemes or not parsed.netloc or parsed.username:
        requirement = "HTTP(S)" if allow_insecure_http else "HTTPS"
        raise ValueError(f"{label} must be an absolute {requirement} URL")
    if parsed.fragment or parsed.query:
        raise ValueError(f"{label} cannot contain a query string or fragment")
    return value.rstrip("/")


def _fetch_json(url: str, timeout_seconds: float) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/json", "User-Agent": "DocXUnravel-MCP/0.4"},
    )
    with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
        final_url = response.geturl()
        if urlparse(final_url).scheme != urlparse(url).scheme:
            raise ValueError("OIDC metadata redirected to a different URL scheme")
        payload = response.read(_MAX_METADATA_BYTES + 1)
    if len(payload) > _MAX_METADATA_BYTES:
        raise ValueError("OIDC metadata response exceeds the size limit")
    value = json.loads(payload)
    if not isinstance(value, dict):
        raise TypeError("OIDC metadata response must be a JSON object")
    return value


def _claim_scopes(claims: Mapping[str, Any]) -> list[str]:
    scopes: set[str] = set()
    for name in ("scope", "scp"):
        value = claims.get(name)
        if value is None:
            continue
        if isinstance(value, str):
            scopes.update(value.split())
        elif isinstance(value, list) and all(isinstance(item, str) for item in value):
            scopes.update(value)
        else:
            raise ValueError(f"JWT {name} claim must be a string or string array")
    return sorted(scopes)


class OIDCJWTVerifier:
    """Verify signed JWT access tokens from an external OIDC issuer.

    The verifier performs discovery and JWKS retrieval on demand, caches keys for a
    bounded interval, and refreshes once when a token references an unknown key ID.
    A static JWKS can be injected for deterministic tests without network access.
    """

    def __init__(
        self,
        *,
        issuer: str,
        audience: str,
        algorithms: Sequence[str] = ("RS256", "ES256"),
        allowed_client_ids: Sequence[str] = (),
        allow_insecure_http: bool = False,
        cache_seconds: int = 300,
        timeout_seconds: float = 5.0,
        clock_skew_seconds: int = 30,
        jwks: Mapping[str, Any] | None = None,
    ) -> None:
        self.issuer = _require_secure_url(issuer, "OIDC issuer", allow_insecure_http)
        if not audience or any(character.isspace() for character in audience):
            raise ValueError("OIDC audience must be a non-empty identifier")
        selected_algorithms = tuple(dict.fromkeys(algorithms))
        if not selected_algorithms or not set(selected_algorithms) <= _SAFE_ALGORITHMS:
            raise ValueError(
                "OIDC algorithms must be selected from approved RSA/ECDSA algorithms"
            )
        if cache_seconds < 30:
            raise ValueError("OIDC JWKS cache must be at least 30 seconds")
        if timeout_seconds <= 0 or clock_skew_seconds < 0:
            raise ValueError("OIDC timeout and clock skew must be non-negative")

        self.audience = audience
        self.algorithms = selected_algorithms
        self.allowed_client_ids = frozenset(allowed_client_ids)
        self.allow_insecure_http = allow_insecure_http
        self.cache_seconds = cache_seconds
        self.timeout_seconds = timeout_seconds
        self.clock_skew_seconds = clock_skew_seconds
        self._static_jwks = dict(jwks) if jwks is not None else None
        self._cached_jwks: dict[str, Any] | None = None
        self._cached_at = 0.0
        self._lock = asyncio.Lock()

    async def _load_jwks(self, *, refresh: bool = False) -> dict[str, Any]:
        if self._static_jwks is not None:
            return self._static_jwks
        now = time.monotonic()
        if (
            not refresh
            and self._cached_jwks is not None
            and now - self._cached_at < self.cache_seconds
        ):
            return self._cached_jwks

        async with self._lock:
            now = time.monotonic()
            if (
                not refresh
                and self._cached_jwks is not None
                and now - self._cached_at < self.cache_seconds
            ):
                return self._cached_jwks
            discovery_url = self.issuer + "/.well-known/openid-configuration"
            metadata = await asyncio.to_thread(
                _fetch_json, discovery_url, self.timeout_seconds
            )
            if metadata.get("issuer") != self.issuer:
                raise ValueError(
                    "OIDC discovery issuer does not match configured issuer"
                )
            jwks_url = metadata.get("jwks_uri")
            if not isinstance(jwks_url, str):
                raise TypeError("OIDC discovery does not provide jwks_uri")
            jwks_url = _require_secure_url(
                jwks_url, "OIDC JWKS URI", self.allow_insecure_http
            )
            jwks = await asyncio.to_thread(_fetch_json, jwks_url, self.timeout_seconds)
            if not isinstance(jwks.get("keys"), list):
                raise TypeError("OIDC JWKS must contain a keys array")
            self._cached_jwks = jwks
            self._cached_at = time.monotonic()
            return jwks

    def _decode(self, token: str, jwks: Mapping[str, Any]) -> dict[str, Any]:
        header = jwt.get_unverified_header(token)
        algorithm = header.get("alg")
        key_id = header.get("kid")
        if algorithm not in self.algorithms or algorithm == "none":
            raise jwt.InvalidAlgorithmError("JWT algorithm is not allowed")
        if not isinstance(key_id, str) or not key_id:
            raise jwt.InvalidKeyError("JWT header must contain a key ID")

        keys = jwks.get("keys")
        if not isinstance(keys, list):
            raise jwt.InvalidKeyError("JWKS keys are unavailable")
        matching = [
            value
            for value in keys
            if isinstance(value, dict)
            and value.get("kid") == key_id
            and value.get("use", "sig") == "sig"
            and value.get("alg", algorithm) == algorithm
        ]
        if len(matching) != 1:
            raise jwt.InvalidKeyError("JWT signing key is unknown or ambiguous")
        signing_key = jwt.PyJWK.from_dict(matching[0], algorithm=algorithm).key
        claims = jwt.decode(
            token,
            key=signing_key,
            algorithms=list(self.algorithms),
            audience=self.audience,
            issuer=self.issuer,
            leeway=self.clock_skew_seconds,
            options={"require": ["exp", "iss", "aud"]},
        )
        if not isinstance(claims, dict):
            raise jwt.InvalidTokenError("JWT claims must be an object")
        return claims

    async def verify_token(self, token: str) -> AccessToken | None:
        """Return MCP access metadata for a valid token, otherwise ``None``."""

        if not token or len(token) > 16_384:
            return None
        try:
            jwks = await self._load_jwks()
            try:
                claims = self._decode(token, jwks)
            except (jwt.InvalidKeyError, jwt.InvalidSignatureError):
                if self._static_jwks is not None:
                    raise
                claims = self._decode(token, await self._load_jwks(refresh=True))

            audiences = claims.get("aud")
            audience_values = audiences if isinstance(audiences, list) else [audiences]
            authorized_party = claims.get("azp")
            if len(audience_values) > 1 and not isinstance(authorized_party, str):
                raise jwt.InvalidTokenError("Multi-audience JWT must contain azp")
            client_id = (
                authorized_party or claims.get("client_id") or claims.get("appid")
            )
            if not isinstance(client_id, str) or not client_id:
                raise jwt.InvalidTokenError(
                    "JWT does not identify an authorized client"
                )
            if self.allowed_client_ids and client_id not in self.allowed_client_ids:
                raise jwt.InvalidTokenError("JWT client is not authorized")

            subject = claims.get("sub")
            if subject is not None and not isinstance(subject, str):
                raise jwt.InvalidTokenError("JWT subject must be a string")
            expires_at = claims.get("exp")
            if not isinstance(expires_at, int) or isinstance(expires_at, bool):
                raise jwt.InvalidTokenError("JWT expiry must be an integer")
            return AccessToken(
                token=token,
                client_id=client_id,
                scopes=_claim_scopes(claims),
                expires_at=expires_at,
                resource=self.audience,
                subject=subject,
                claims={"iss": self.issuer, "aud": audiences, "azp": client_id},
            )
        except (jwt.PyJWTError, ValueError, TypeError, OSError, json.JSONDecodeError):
            return None

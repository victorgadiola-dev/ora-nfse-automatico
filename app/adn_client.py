from __future__ import annotations

import base64
import gzip
import json
import os
import zlib
from dataclasses import dataclass
from io import BytesIO
from typing import Any
import zipfile

try:
    import requests_pkcs12
except Exception:  # pragma: no cover
    requests_pkcs12 = None


@dataclass(slots=True)
class AdnResponse:
    status_code: int
    content_type: str
    content: bytes
    headers: dict[str, str]
    json_body: Any | None = None
    text_preview: str | None = None


class AdnClientError(RuntimeError):
    pass


class AdnClient:
    """Cliente da API ADN/contribuintes da NFS-e Nacional.

    Usa certificado A1 em formato PFX/P12 e consulta GET /DFe/{NSU}.
    """

    def __init__(self, base_url: str, pfx_path: str, pfx_password: str | None, timeout: int = 60) -> None:
        if requests_pkcs12 is None:
            raise AdnClientError("Instale requests-pkcs12 para usar certificado A1: pip install requests-pkcs12")
        self.base_url = base_url.rstrip("/")
        self.pfx_path = pfx_path
        self.pfx_password = pfx_password
        self.timeout = timeout

    def get_dfe(self, nsu: int, cnpj_consulta: str | None = None) -> AdnResponse:
        url = f"{self.base_url}/DFe/{nsu}"
        params = {"cnpjConsulta": cnpj_consulta} if cnpj_consulta else None
        response = requests_pkcs12.get(
            url,
            params=params,
            pkcs12_filename=self.pfx_path,
            pkcs12_password=self.pfx_password,
            timeout=self.timeout,
            headers={"Accept": "application/json, application/xml, text/xml, application/zip, */*"},
        )
        return _to_adn_response(response)

    def get_eventos(self, chave_acesso: str) -> AdnResponse:
        url = f"{self.base_url}/NFSe/{chave_acesso}/Eventos"
        response = requests_pkcs12.get(
            url,
            pkcs12_filename=self.pfx_path,
            pkcs12_password=self.pfx_password,
            timeout=self.timeout,
            headers={"Accept": "application/json, application/xml, text/xml, application/zip, */*"},
        )
        return _to_adn_response(response)


def _to_adn_response(response: Any) -> AdnResponse:
    content_type = response.headers.get("content-type", "")
    json_body = None
    if "json" in content_type.lower():
        try:
            json_body = response.json()
        except json.JSONDecodeError:
            json_body = None
    preview = response.text[:600] if response.content else None
    return AdnResponse(
        status_code=response.status_code,
        content_type=content_type,
        content=response.content,
        headers={str(k): str(v) for k, v in response.headers.items()},
        json_body=json_body,
        text_preview=preview,
    )


def password_from_env_or_value(env_name: str | None, plain_value: str | None) -> str | None:
    if env_name:
        return os.environ.get(env_name)
    return plain_value


def _looks_like_xml(blob: bytes) -> bool:
    return blob.lstrip().startswith(b"<")


def _maybe_decompress(blob: bytes) -> bytes:
    if not blob:
        return blob
    if blob[:2] == b"\x1f\x8b":
        return gzip.decompress(blob)
    try:
        if blob[:1] in {b"x", b"\x78"}:
            return zlib.decompress(blob)
    except Exception:
        pass
    return blob


def _try_base64_to_bytes(value: str) -> bytes | None:
    cleaned = value.strip()
    if not cleaned:
        return None
    if "," in cleaned and "base64" in cleaned[:120].lower():
        cleaned = cleaned.split(",", 1)[1]
    compact = "".join(cleaned.split())
    if len(compact) < 24:
        return None
    try:
        return base64.b64decode(compact, validate=True)
    except Exception:
        try:
            return base64.b64decode(compact + "=" * (-len(compact) % 4), validate=False)
        except Exception:
            return None


def _extract_xml_from_bytes(blob: bytes, label: str = "documento") -> list[tuple[str, bytes]]:
    candidates = [blob]
    try:
        decompressed = _maybe_decompress(blob)
        if decompressed != blob:
            candidates.insert(0, decompressed)
    except Exception:
        pass

    for candidate in candidates:
        if _looks_like_xml(candidate):
            return [(f"{label}.xml", candidate)]

        bio = BytesIO(candidate)
        if zipfile.is_zipfile(bio):
            bio.seek(0)
            result: list[tuple[str, bytes]] = []
            with zipfile.ZipFile(bio) as zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    data = zf.read(info)
                    nested = _extract_xml_from_bytes(data, info.filename.rsplit(".", 1)[0] or label)
                    result.extend(nested)
            return result
    return []


def extract_xml_payloads(response: AdnResponse) -> list[tuple[str, bytes]]:
    """Extrai XMLs de respostas comuns do ADN.

    Suporta XML direto, ZIP, JSON com XML em texto, base64, gzip+base64 ou zip+base64.
    """
    direct = _extract_xml_from_bytes(response.content, "resposta_adn")
    if direct:
        return direct

    if response.json_body is None:
        try:
            body = json.loads(response.content.decode("utf-8"))
        except Exception:
            return []
    else:
        body = response.json_body

    result: list[tuple[str, bytes]] = []
    seen_hashes: set[int] = set()

    def add_extracted(items: list[tuple[str, bytes]]) -> None:
        for label, data in items:
            key = hash(data)
            if key in seen_hashes:
                continue
            seen_hashes.add(key)
            result.append((label, data))

    def visit(value: Any, label: str = "documento") -> None:
        if value is None:
            return
        if isinstance(value, dict):
            for key, child in value.items():
                visit(child, str(key) if key else label)
            return
        if isinstance(value, list):
            for index, child in enumerate(value):
                visit(child, f"{label}_{index + 1}")
            return
        if isinstance(value, str):
            raw_text = value.strip()
            if raw_text.startswith("<"):
                add_extracted([(f"{label}.xml", raw_text.encode("utf-8"))])
                return
            decoded = _try_base64_to_bytes(raw_text)
            if decoded:
                add_extracted(_extract_xml_from_bytes(decoded, label))

    visit(body)
    return result

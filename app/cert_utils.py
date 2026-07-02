from __future__ import annotations

import re
from pathlib import Path

from cryptography.hazmat.primitives.serialization import pkcs12


class CertificadoErro(RuntimeError):
    pass


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_.-]+", "_", value.strip())
    return cleaned.strip("._") or "certificado"


def extract_cnpj_root_from_pfx(path: str | Path, password: str | None) -> str | None:
    """Valida o PFX/P12 e tenta localizar a raiz do CNPJ no certificado.

    Essa extração é por melhor esforço porque certificados de autoridades
    diferentes podem guardar o documento em campos distintos.
    """
    data = Path(path).read_bytes()
    pwd = password.encode("utf-8") if password else None
    try:
        key, cert, _cas = pkcs12.load_key_and_certificates(data, pwd)
    except Exception as exc:  # noqa: BLE001
        raise CertificadoErro("Não consegui abrir o certificado. Confirme se é A1 .pfx/.p12 e se a senha está correta.") from exc
    if cert is None or key is None:
        raise CertificadoErro("O arquivo informado não contém certificado e chave privada.")

    text = " ".join([
        cert.subject.rfc4514_string(),
        cert.issuer.rfc4514_string(),
        str(cert.serial_number),
    ])
    for match in re.finditer(r"\d{14}", text):
        return match.group(0)[:8]
    return None

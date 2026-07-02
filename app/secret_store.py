from __future__ import annotations

from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken

from app.config import get_settings


class SecretStoreError(RuntimeError):
    pass


def get_or_create_secret_key() -> bytes:
    settings = get_settings()
    path = Path(settings.secret_key_file)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        value = path.read_bytes().strip()
        if value:
            return value
    key = Fernet.generate_key()
    path.write_bytes(key)
    return key


def _fernet() -> Fernet:
    return Fernet(get_or_create_secret_key())


def encrypt_secret(value: str | None) -> str | None:
    if value is None:
        return None
    return _fernet().encrypt(value.encode("utf-8")).decode("ascii")


def decrypt_secret(value: str | None) -> str | None:
    if not value:
        return ""
    try:
        return _fernet().decrypt(value.encode("ascii")).decode("utf-8")
    except InvalidToken as exc:
        raise SecretStoreError("Não consegui descriptografar a senha local. Verifique se a chave em data/.ora_nfse_secret.key é a mesma usada quando o certificado foi cadastrado.") from exc

from __future__ import annotations

from functools import lru_cache
import os
from pathlib import Path

from pydantic import model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Configuração da aplicação.

    A aplicação pode rodar em servidor Render ou, apenas para desenvolvimento,
    localmente. Em produção no Render, use DATA_DIR apontando para o caminho do disco persistente para que
    JSON, XMLs e certificados criptografados sobrevivam a redeploys.
    """

    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    app_name: str = "ORA NFS-e Automático"
    app_version: str = "0.15.0"
    app_env: str = "auto"

    # Autenticação simples para publicação em URL pública.
    # Em produção no Render, deixe REQUIRE_AUTH=true e configure APP_ACCESS_PASSWORD.
    require_auth: bool = False
    app_access_password: str = ""
    app_session_secret: str = ""
    session_max_age_seconds: int = 12 * 60 * 60
    secure_cookies: bool = False
    app_public_url: str = ""

    data_dir: str = "./data"
    storage_file: str | None = None
    cert_store_dir: str | None = None
    xml_store_dir: str | None = None
    secret_key_file: str | None = None

    nfse_adn_base_url: str = "https://adn.nfse.gov.br/contribuintes"
    request_timeout_seconds: int = 60
    request_delay_seconds: float = 0.4
    rate_limit_pause_seconds: int = 45
    rate_limit_max_pause_seconds: int = 300
    max_rate_limit_retries: int = 3

    default_max_consultas: int = 500
    default_stop_after_empty: int = 3

    allowed_origins: str = ""

    @model_validator(mode="after")
    def derive_storage_paths(self) -> "Settings":
        running_on_render = bool(os.getenv("RENDER") or os.getenv("RENDER_SERVICE_ID") or os.getenv("RENDER_SERVICE_NAME"))
        declared_render = self.app_env.lower() == "render"
        render_like = running_on_render or declared_render

        if self.app_env == "auto" or (running_on_render and self.app_env.lower() == "local"):
            self.app_env = "render" if running_on_render else "local"
            declared_render = self.app_env.lower() == "render"
            render_like = running_on_render or declared_render

        if render_like and self.data_dir.strip() in {"./data", "data", ""}:
            self.data_dir = "/opt/render/project/src/data"

        # URL pública e cookies seguros são padrão no Render.
        if render_like and not self.secure_cookies:
            self.secure_cookies = True

        # Em ambiente público, o sistema deve exigir senha. Se a senha não estiver
        # definida, a tela de login mostra instrução de configuração em vez de abrir dados fiscais.
        if render_like and not self.require_auth:
            self.require_auth = True

        if render_like and not self.app_public_url:
            service_name = os.getenv("RENDER_SERVICE_NAME") or ""
            if service_name:
                self.app_public_url = f"https://{service_name}.onrender.com"

        base = Path(self.data_dir)
        if not self.storage_file:
            self.storage_file = str(base / "ora_nfse_storage.json")
        if not self.cert_store_dir:
            self.cert_store_dir = str(base / "certificados")
        if not self.xml_store_dir:
            self.xml_store_dir = str(base / "xmls")
        if not self.secret_key_file:
            self.secret_key_file = str(base / ".ora_nfse_secret.key")
        return self

    def ensure_dirs(self) -> None:
        for value in [self.data_dir, self.cert_store_dir, self.xml_store_dir]:
            if value:
                Path(value).mkdir(parents=True, exist_ok=True)


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    settings = Settings()
    settings.ensure_dirs()
    return settings

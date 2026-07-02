from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Configuração local.

    O projeto foi desenhado para funcionar sem banco de dados. Tudo que é dado
    de uso fica dentro de DATA_DIR e não deve ser enviado para repositórios.
    """

    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    app_name: str = "ORA NFS-e Automático"
    app_version: str = "0.10.0"

    data_dir: str = "./data"
    storage_file: str = "./data/ora_nfse_storage.json"
    cert_store_dir: str = "./data/certificados"
    xml_store_dir: str = "./data/xmls"
    secret_key_file: str = "./data/.ora_nfse_secret.key"

    nfse_adn_base_url: str = "https://adn.nfse.gov.br/contribuintes"
    request_timeout_seconds: int = 60
    request_delay_seconds: float = 0.4
    rate_limit_pause_seconds: int = 45
    rate_limit_max_pause_seconds: int = 300
    max_rate_limit_retries: int = 3

    default_max_consultas: int = 500
    default_stop_after_empty: int = 3

    def ensure_dirs(self) -> None:
        for value in [self.data_dir, self.cert_store_dir, self.xml_store_dir]:
            Path(value).mkdir(parents=True, exist_ok=True)


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    settings = Settings()
    settings.ensure_dirs()
    return settings

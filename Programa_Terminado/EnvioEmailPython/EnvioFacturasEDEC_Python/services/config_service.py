from configparser import ConfigParser
from pathlib import Path
from typing import Optional
import sys


class ConfigService:
    """Carrega e valida o ficheiro config.ini."""

    def __init__(self, caminho: Optional[str] = None):
        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
        candidatos = []
        if caminho:
            candidatos.append(Path(caminho))
        candidatos.extend([Path.cwd() / "config.ini", base / "config.ini"])

        self.path = next((p for p in candidatos if p.exists()), None)
        if self.path is None:
            raise FileNotFoundError("O ficheiro config.ini não foi encontrado.")

        self.config = ConfigParser()
        if not self.config.read(self.path, encoding="utf-8"):
            raise RuntimeError("Não foi possível ler o ficheiro config.ini.")
        self._validar()

    def get(self, secao: str, chave: str, fallback: str = "") -> str:
        if not self.config.has_section(secao):
            return fallback
        return self.config.get(secao, chave, fallback=fallback).strip()

    def getint(self, secao: str, chave: str, fallback: int = 0) -> int:
        try:
            return self.config.getint(secao, chave, fallback=fallback)
        except (TypeError, ValueError):
            return fallback

    def _validar(self) -> None:
        obrigatorios = {
            "DATABASE": ["HOST", "PORT", "DATABASE", "USERNAME", "PASSWORD"],
            "EMAIL": [
                "REMETENTE", "USERNAME", "PASSWORD", "ASSUNTO",
                "SMTP_HOST", "SMTP_PORT", "SMTP_SEGURANCA"
            ],
        }
        for secao, chaves in obrigatorios.items():
            if not self.config.has_section(secao):
                raise ValueError("A secção [{}] não existe no config.ini.".format(secao))
            for chave in chaves:
                if not self.get(secao, chave):
                    raise ValueError(
                        "Configuração obrigatória vazia: [{}] {}".format(secao, chave)
                    )

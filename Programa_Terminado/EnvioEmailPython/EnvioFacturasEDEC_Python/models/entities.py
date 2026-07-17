from dataclasses import dataclass
from datetime import datetime


@dataclass
class Cliente:
    cil: str
    nome: str
    email: str
    arquivo_anexo: str = ""


@dataclass
class Usuario:
    id: int
    username: str
    password: str
    nivel: str

    @property
    def is_admin(self) -> bool:
        return self.nivel.lower() in {"admin", "administrador"}

    @property
    def is_gerente(self) -> bool:
        return self.nivel.lower() == "gerente"


@dataclass
class RelatorioEnvio:
    nome: str
    email: str
    cil: str
    status: str
    mensagem: str
    data_envio: datetime

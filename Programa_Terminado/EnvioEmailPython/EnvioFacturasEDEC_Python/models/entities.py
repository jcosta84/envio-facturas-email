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
    email: str = ""
    nome: str = ""
    estado: str = "ativo"

    @property
    def is_admin(self):
        return self.nivel == "admin"

    @property
    def is_gerente(self):
        return self.nivel == "gerente"


@dataclass
class RelatorioEnvio:
    nome: str
    email: str
    cil: str
    status: str
    mensagem: str
    data_envio: datetime
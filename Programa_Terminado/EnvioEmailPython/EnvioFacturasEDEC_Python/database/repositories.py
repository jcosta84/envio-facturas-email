from datetime import date
from typing import Optional
import platform
import socket

from models.entities import Cliente, Usuario, RelatorioEnvio
from models.lista_dupla import ListaDupla


class ClienteRepository:
    def __init__(self, db):
        self.db = db

    def inserir(self, cliente: Cliente):
        arquivo_anexo = (cliente.arquivo_anexo or "").strip()
        if not arquivo_anexo:
            arquivo_anexo = f"{cliente.cil}.pdf"

        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO clientes (cil, nome, email, arquivo_anexo)
                VALUES (%s, %s, %s, %s)
                """,
                (cliente.cil, cliente.nome, cliente.email, arquivo_anexo),
            )
            conn.commit()
            return cur.lastrowid

    def listar(self):
        lista = ListaDupla()
        with self.db.conectar() as conn:
            cur = conn.cursor(dictionary=True)
            cur.execute(
                """
                SELECT cil, nome, email, arquivo_anexo
                FROM clientes
                ORDER BY nome
                """
            )
            for registo in cur:
                lista.adicionar(
                    Cliente(
                        registo["cil"],
                        registo["nome"],
                        registo["email"],
                        registo.get("arquivo_anexo") or "",
                    )
                )
        return lista

    def atualizar(self, cil_original, cliente: Cliente):
        arquivo_anexo = (cliente.arquivo_anexo or "").strip()
        if not arquivo_anexo:
            arquivo_anexo = f"{cliente.cil}.pdf"

        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                UPDATE clientes
                SET cil = %s, nome = %s, email = %s, arquivo_anexo = %s
                WHERE cil = %s
                """,
                (cliente.cil, cliente.nome, cliente.email, arquivo_anexo, cil_original),
            )
            conn.commit()
            return cur.rowcount > 0

    def eliminar(self, cil):
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM clientes WHERE cil = %s", (cil,))
            conn.commit()
            return cur.rowcount > 0


class CcRepository:
    def __init__(self, db):
        self.db = db

    def listar(self):
        lista = ListaDupla()
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute("SELECT email_cc FROM cc_email ORDER BY email_cc")
            for (email,) in cur:
                lista.adicionar(email)
        return lista

    def inserir(self, email):
        email = email.strip()
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute("INSERT INTO cc_email(email_cc) VALUES(%s)", (email,))
            conn.commit()
            return cur.lastrowid

    def eliminar(self, email):
        email = email.strip()
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM cc_email WHERE email_cc = %s", (email,))
            conn.commit()
            return cur.rowcount > 0


class ConfiguracaoRepository:
    def __init__(self, db):
        self.db = db

    def obter_corpo_email(self):
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute("SELECT conteudo FROM corpo_email WHERE id = 1")
            registo = cur.fetchone()
            return registo[0] if registo else ""

    def guardar_corpo_email(self, texto):
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO corpo_email(id, conteudo)
                VALUES(1, %s)
                ON DUPLICATE KEY UPDATE conteudo = VALUES(conteudo)
                """,
                (texto,),
            )
            conn.commit()


class AuditoriaRepository:
    def __init__(self, db):
        self.db = db

    @staticmethod
    def ip():
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "DESCONHECIDO"

    def acesso(self, id_usuario, username, resultado, observacao):
        try:
            with self.db.conectar() as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    INSERT INTO acessos(
                        id_usuario,
                        username_informado,
                        resultado,
                        ip,
                        sistema_operativo,
                        navegador,
                        observacao,
                        data_hora
                    )
                    VALUES(%s, %s, %s, %s, %s, %s, %s, NOW())
                    """,
                    (
                        id_usuario,
                        username,
                        resultado,
                        self.ip(),
                        platform.system(),
                        "Aplicação Python",
                        observacao,
                    ),
                )
                conn.commit()
        except Exception as exc:
            print("Erro ao registar acesso:", exc)

    def log(self, id_usuario, operacao, modulo, descricao):
        try:
            with self.db.conectar() as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    INSERT INTO logs(
                        id_usuario,
                        operacao,
                        modulo,
                        descricao,
                        ip,
                        data_hora
                    )
                    VALUES(%s, %s, %s, %s, %s, NOW())
                    """,
                    (id_usuario, operacao, modulo, descricao, self.ip()),
                )
                conn.commit()
        except Exception as exc:
            print("Erro ao registar log:", exc)


class UsuarioRepository:
    def __init__(self, db):
        self.db = db
        self.aud = AuditoriaRepository(db)

    @staticmethod
    def _criar_usuario(registo):
        return Usuario(
            id=registo["id"],
            username=registo["username"],
            password=registo["password"],
            nivel=registo["nivel"],
            email=registo.get("email") or "",
            nome=registo.get("nome") or "",
            estado=registo.get("estado") or "ativo",
        )

    def autenticar(self, username, password):
        username = username.strip()
        with self.db.conectar() as conn:
            cur = conn.cursor(dictionary=True)
            cur.execute(
                """
                SELECT id, username, password, nome, email, nivel,
                       estado, data_inicio, data_fim
                FROM usuarios
                WHERE username = %s
                LIMIT 1
                """,
                (username,),
            )
            registo = cur.fetchone()

            if not registo:
                self.aud.acesso(None, username, "UTILIZADOR_INEXISTENTE", "Utilizador não existe.")
                return None

            if registo["password"] != password:
                self.aud.acesso(registo["id"], username, "SENHA_INCORRETA", "Palavra-passe incorreta.")
                return None

            estado = registo.get("estado") or "ativo"
            if estado != "ativo":
                self.aud.acesso(registo["id"], username, estado.upper(), "Conta indisponível.")
                return None

            hoje = date.today()
            data_inicio = registo.get("data_inicio")
            data_fim = registo.get("data_fim")

            if data_inicio and hoje < data_inicio:
                self.aud.acesso(registo["id"], username, "CONTA_NAO_INICIADA", "A conta ainda não está válida.")
                return None

            if data_fim and hoje > data_fim:
                self.aud.acesso(registo["id"], username, "CONTA_EXPIRADA", "A validade da conta terminou.")
                return None

            cur.execute(
                """
                UPDATE usuarios
                SET ultimo_acesso = NOW(), ultimo_ip = %s, tentativas_login = 0
                WHERE id = %s
                """,
                (self.aud.ip(), registo["id"]),
            )
            conn.commit()

            usuario = self._criar_usuario(registo)
            self.aud.acesso(usuario.id, usuario.username, "SUCESSO", "Login realizado.")
            return usuario

    def listar(self):
        lista = ListaDupla()
        with self.db.conectar() as conn:
            cur = conn.cursor(dictionary=True)
            cur.execute(
                """
                SELECT id, username, password, nome, email, nivel, estado
                FROM usuarios
                ORDER BY username
                """
            )
            for registo in cur:
                lista.adicionar(self._criar_usuario(registo))
        return lista

    def inserir(self, username, email, password, nivel):
        username = username.strip()
        email = email.strip()
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO usuarios(username, email, password, nivel, estado)
                VALUES(%s, %s, %s, %s, 'ativo')
                """,
                (username, email, password, nivel),
            )
            conn.commit()
            return cur.lastrowid

    def atualizar(self, id_, username, email, password, nivel):
        username = username.strip()
        email = email.strip()
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                UPDATE usuarios
                SET username = %s, email = %s, password = %s, nivel = %s
                WHERE id = %s
                """,
                (username, email, password, nivel, id_),
            )
            conn.commit()
            return cur.rowcount > 0

    def eliminar(self, id_):
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM usuarios WHERE id = %s", (id_,))
            conn.commit()
            return cur.rowcount > 0


class RelatorioRepository:
    def __init__(self, db):
        self.db = db

    def inserir(self, relatorio: RelatorioEnvio):
        with self.db.conectar() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO relatorio(nome, email, cil, status, mensagem, data_envio)
                VALUES(%s, %s, %s, %s, %s, %s)
                """,
                (
                    relatorio.nome,
                    relatorio.email,
                    relatorio.cil,
                    relatorio.status,
                    relatorio.mensagem,
                    relatorio.data_envio,
                ),
            )
            conn.commit()
            return cur.lastrowid

    def listar(self, inicio: Optional[date] = None, fim: Optional[date] = None):
        sql = """
            SELECT nome, email, cil, status, mensagem, data_envio
            FROM relatorio
            WHERE 1 = 1
        """
        parametros = []

        if inicio:
            sql += " AND DATE(data_envio) >= %s"
            parametros.append(inicio)

        if fim:
            sql += " AND DATE(data_envio) <= %s"
            parametros.append(fim)

        sql += " ORDER BY data_envio DESC"
        lista = ListaDupla()

        with self.db.conectar() as conn:
            cur = conn.cursor(dictionary=True)
            cur.execute(sql, parametros)
            for registo in cur:
                lista.adicionar(RelatorioEnvio(**registo))

        return lista
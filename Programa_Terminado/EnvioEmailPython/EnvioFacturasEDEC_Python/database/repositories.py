from datetime import date, datetime
from typing import Optional
import platform, socket
from models.entities import Cliente, Usuario, RelatorioEnvio
from models.lista_dupla import ListaDupla

class ClienteRepository:
    def __init__(self, db): self.db=db
    def inserir(self, c: Cliente):
        arquivo_anexo = c.arquivo_anexo.strip()

        if not arquivo_anexo:
            arquivo_anexo = "{}.pdf".format(c.cil)

        with self.db.conectar() as conn:
            cur = conn.cursor()

            cur.execute(
                """
                INSERT INTO clientes(
                    cil,
                    nome,
                    email,
                    arquivo_anexo
                )
                VALUES(%s, %s, %s, %s)
                """,
                (
                    c.cil,
                    c.nome,
                    c.email,
                    arquivo_anexo
                )
            )

            conn.commit()
    def listar(self):
        lista=ListaDupla()
        with self.db.conectar() as conn:
            cur=conn.cursor(dictionary=True); cur.execute("SELECT cil,nome,email,arquivo_anexo FROM clientes ORDER BY nome")
            for r in cur: lista.adicionar(Cliente(r['cil'],r['nome'],r['email'],r.get('arquivo_anexo') or ''))
        return lista
    def atualizar(self, cil_original, c: Cliente):
        arquivo_anexo = "{}.pdf".format(c.cil)

        with self.db.conectar() as conn:
            cur = conn.cursor()

            cur.execute(
                """
                UPDATE clientes
                SET
                    cil = %s,
                    nome = %s,
                    email = %s,
                    arquivo_anexo = %s
                WHERE cil = %s
                """,
                (
                    c.cil,
                    c.nome,
                    c.email,
                    arquivo_anexo,
                    cil_original
                )
            )

            conn.commit()

            return cur.rowcount > 0;

    def eliminar(self, cil):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("DELETE FROM clientes WHERE cil=%s",(cil,)); conn.commit(); return cur.rowcount>0

class CcRepository:
    def __init__(self, db): self.db=db
    def listar(self):
        lista=ListaDupla()
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("SELECT email_cc FROM cc_email ORDER BY email_cc")
            for (email,) in cur: lista.adicionar(email)
        return lista
    def inserir(self,email):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("INSERT INTO cc_email(email_cc) VALUES(%s)",(email.strip(),)); conn.commit()
    def eliminar(self,email):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("DELETE FROM cc_email WHERE email_cc=%s",(email.strip(),)); conn.commit()

class ConfiguracaoRepository:
    def __init__(self, db): self.db=db
    def obter_corpo_email(self):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("SELECT conteudo FROM corpo_email WHERE id=1"); r=cur.fetchone(); return r[0] if r else ""
    def guardar_corpo_email(self, texto):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("INSERT INTO corpo_email(id,conteudo) VALUES(1,%s) ON DUPLICATE KEY UPDATE conteudo=VALUES(conteudo)",(texto,)); conn.commit()

class AuditoriaRepository:
    def __init__(self, db): self.db=db
    @staticmethod
    def ip():
        try:return socket.gethostbyname(socket.gethostname())
        except:return "DESCONHECIDO"
    def acesso(self,id_usuario,username,resultado,observacao):
        try:
            with self.db.conectar() as conn:
                cur=conn.cursor(); cur.execute("INSERT INTO acessos(id_usuario,username_informado,resultado,ip,sistema_operativo,navegador,observacao,data_hora) VALUES(%s,%s,%s,%s,%s,%s,%s,NOW())",(id_usuario,username,resultado,self.ip(),platform.system(),"Aplicação Python",observacao)); conn.commit()
        except Exception: pass
    def log(self,id_usuario,operacao,modulo,descricao):
        try:
            with self.db.conectar() as conn:
                cur=conn.cursor(); cur.execute("INSERT INTO logs(id_usuario,operacao,modulo,descricao,ip,data_hora) VALUES(%s,%s,%s,%s,%s,NOW())",(id_usuario,operacao,modulo,descricao,self.ip())); conn.commit()
        except Exception: pass

class UsuarioRepository:
    def __init__(self, db): self.db=db; self.aud=AuditoriaRepository(db)
    def autenticar(self, username, password):
        with self.db.conectar() as conn:
            cur=conn.cursor(dictionary=True); cur.execute("SELECT * FROM usuarios WHERE username=%s LIMIT 1",(username.strip(),)); r=cur.fetchone()
            if not r: self.aud.acesso(None,username,"UTILIZADOR_INEXISTENTE","Utilizador não existe."); return None
            if r['password'] != password: self.aud.acesso(r['id'],username,"SENHA_INCORRETA","Palavra-passe incorreta."); return None
            if r.get('estado','ativo') != 'ativo': self.aud.acesso(r['id'],username,str(r.get('estado')).upper(),"Conta indisponível."); return None
            hoje=date.today(); inicio=r.get('data_inicio'); fim=r.get('data_fim')
            if inicio and hoje<inicio: return None
            if fim and hoje>fim: return None
            cur.execute("UPDATE usuarios SET ultimo_acesso=NOW(),ultimo_ip=%s,tentativas_login=0 WHERE id=%s",(self.aud.ip(),r['id'])); conn.commit()
            u=Usuario(r['id'],r['username'],r['password'],r['nivel']); self.aud.acesso(u.id,u.username,"SUCESSO","Login realizado."); return u
    def listar(self):
        lista=ListaDupla()
        with self.db.conectar() as conn:
            cur=conn.cursor(dictionary=True); cur.execute("SELECT id,username,password,nivel FROM usuarios ORDER BY username")
            for r in cur: lista.adicionar(Usuario(r['id'],r['username'],r['password'],r['nivel']))
        return lista
    def inserir(self,username,password,nivel):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("INSERT INTO usuarios(username,password,nivel) VALUES(%s,%s,%s)",(username.strip(),password,nivel)); conn.commit()
    def atualizar(self,id_,username,password,nivel):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("UPDATE usuarios SET username=%s,password=%s,nivel=%s WHERE id=%s",(username.strip(),password,nivel,id_)); conn.commit()
    def eliminar(self,id_):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("DELETE FROM usuarios WHERE id=%s",(id_,)); conn.commit()

class RelatorioRepository:
    def __init__(self, db): self.db=db
    def inserir(self,r:RelatorioEnvio):
        with self.db.conectar() as conn:
            cur=conn.cursor(); cur.execute("INSERT INTO relatorio(nome,email,cil,status,mensagem,data_envio) VALUES(%s,%s,%s,%s,%s,%s)",(r.nome,r.email,r.cil,r.status,r.mensagem,r.data_envio)); conn.commit()
    def listar(self, inicio: Optional[date] = None, fim: Optional[date] = None):
        sql="SELECT nome,email,cil,status,mensagem,data_envio FROM relatorio WHERE 1=1"; p=[]
        if inicio: sql+=" AND DATE(data_envio)>=%s"; p.append(inicio)
        if fim: sql+=" AND DATE(data_envio)<=%s"; p.append(fim)
        sql+=" ORDER BY data_envio DESC"; lista=ListaDupla()
        with self.db.conectar() as conn:
            cur=conn.cursor(dictionary=True); cur.execute(sql,p)
            for r in cur: lista.adicionar(RelatorioEnvio(**r))
        return lista

import customtkinter as ctk
from tkinter import TclError, messagebox

from services.config_service import ConfigService
from database.connection import Database
from database.repositories import UsuarioRepository


class LoginView(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Login — Sistema de Envio de E-mails")
        self.geometry("1000x650")
        self.minsize(900, 600)
        self.configure(fg_color="#0F141F")

        self.main_view = None
        self._fechando = False
        self._focus_after_id = None
        self.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)

        try:
            self.config_service = ConfigService()
            self.db = Database(self.config_service)
            self.db.testar()
            self.repo = UsuarioRepository(self.db)
        except Exception as exc:
            messagebox.showerror("Erro de inicialização", str(exc), parent=self)
            self.after(100, self.fechar_aplicacao)
            return

        self.grid_columnconfigure(0, weight=43)
        self.grid_columnconfigure(1, weight=57)
        self.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(self, corner_radius=0, fg_color="#121824")
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(0, weight=1)

        bloco = ctk.CTkFrame(left, fg_color="transparent")
        bloco.grid(row=0, column=0, padx=46, sticky="w")

        ctk.CTkLabel(bloco, text="✉", text_color="#7868FF", font=ctk.CTkFont(size=34)).pack(anchor="w")
        ctk.CTkLabel(bloco, text="Envio de E-mails", font=ctk.CTkFont(size=31, weight="bold")).pack(anchor="w", pady=(25, 10))
        ctk.CTkLabel(
            bloco,
            text="Gestão de clientes, documentos PDF,\nenvios e relatórios numa única aplicação.",
            justify="left",
            text_color="#B9C4D6",
            font=ctk.CTkFont(size=15),
        ).pack(anchor="w")

        for indice, texto in enumerate(("Gestão de clientes", "Envio automático de PDFs", "Relatórios em Excel e PDF")):
            linha = ctk.CTkFrame(bloco, fg_color="transparent")
            linha.pack(anchor="w", pady=(22 if indice == 0 else 8, 0))
            ctk.CTkLabel(linha, text="✓", width=28, height=28, corner_radius=14, fg_color="#27265D", text_color="#8B83FF").pack(side="left")
            ctk.CTkLabel(linha, text=texto, font=ctk.CTkFont(size=14)).pack(side="left", padx=12)

        ctk.CTkLabel(left, text="Sistema v1.0", text_color="#8290A8", font=ctk.CTkFont(size=11)).grid(row=1, column=0, padx=46, pady=28, sticky="sw")

        right = ctk.CTkFrame(self, corner_radius=0, fg_color="#0D121C")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(0, weight=1)

        card = ctk.CTkFrame(right, width=390, height=560, corner_radius=18, fg_color="#171E2B", border_width=1, border_color="#2B3548")
        card.grid(row=0, column=0, padx=70, pady=55)
        card.grid_propagate(False)

        ctk.CTkLabel(card, text="Bem-vindo", font=ctk.CTkFont(size=31, weight="bold")).pack(anchor="w", padx=40, pady=(42, 6))
        ctk.CTkLabel(card, text="Introduza os seus dados para continuar.", text_color="#9CACCA").pack(anchor="w", padx=40, pady=(0, 34))
        ctk.CTkLabel(card, text="Utilizador", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=40)

        self.user = ctk.CTkEntry(card, placeholder_text="Nome de utilizador", width=308, height=44, corner_radius=9)
        self.user.pack(padx=40, pady=(8, 22))

        ctk.CTkLabel(card, text="Palavra-passe", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=40)
        self.pwd = ctk.CTkEntry(card, placeholder_text="Palavra-passe", show="•", width=308, height=44, corner_radius=9)
        self.pwd.pack(padx=40, pady=(8, 12))

        self.msg = ctk.CTkLabel(card, text="", text_color="#FF7B7B")
        self.msg.pack(padx=40, anchor="w")

        ctk.CTkButton(card, text="Entrar", width=274, height=45, font=ctk.CTkFont(size=14, weight="bold"), command=self.entrar).pack(pady=(25, 0))

        self.bind("<Return>", lambda _event: self.entrar())
        self._focus_after_id = self.after(200, self._focar_utilizador)

    def _focar_utilizador(self):
        self._focus_after_id = None
        try:
            if self.winfo_exists():
                self.user.focus_set()
        except (TclError, RuntimeError):
            pass

    def entrar(self):
        username = self.user.get().strip()
        password = self.pwd.get()

        if not username or not password:
            self.msg.configure(text="Preencha o utilizador e a palavra-passe.")
            return

        self.msg.configure(text="")

        try:
            usuario = self.repo.autenticar(username, password)
        except Exception as exc:
            messagebox.showerror("Erro", str(exc), parent=self)
            return

        if not usuario:
            self.msg.configure(text="Dados inválidos ou conta indisponível.")
            return

        self.abrir_aplicacao(usuario)

    def abrir_aplicacao(self, usuario):
        from views.main_view import MainView

        self.withdraw()

        try:
            self.main_view = MainView(
                self.config_service,
                self.db,
                usuario,
                master=self,
                ao_terminar_sessao=self.voltar_login,
                ao_sair=self.fechar_aplicacao,
            )
            self.main_view.focus_force()
        except Exception as exc:
            self.main_view = None
            self.deiconify()
            messagebox.showerror("Erro", f"Não foi possível abrir a aplicação.\n\n{exc}", parent=self)

    def voltar_login(self):
        if self._fechando:
            return

        self.main_view = None
        self.user.delete(0, "end")
        self.pwd.delete(0, "end")
        self.msg.configure(text="")
        self.deiconify()
        self.lift()
        self.focus_force()
        self._focus_after_id = self.after(100, self._focar_utilizador)

    def fechar_aplicacao(self):
        if self._fechando:
            return

        self._fechando = True

        if self._focus_after_id is not None:
            try:
                self.after_cancel(self._focus_after_id)
            except (TclError, RuntimeError):
                pass
            self._focus_after_id = None

        try:
            if self.main_view is not None:
                self.main_view.cancelar_relogio()
                self.main_view.destroy()
        except (TclError, RuntimeError, AttributeError):
            pass

        self.main_view = None

        try:
            self.destroy()
        except (TclError, RuntimeError):
            pass